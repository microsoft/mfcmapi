#include <core/stdafx.h>
#include <core/mapi/version.h>
#include <core/utility/import.h>
#include <mapistub/library/stubutils.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/utility/error.h>
#include <appmodel.h>
#include <AppxPackaging.h>

namespace version
{
	std::wstring GetMSIVersion()
	{
		std::wstring szOut;

		for (const auto component : mapistub::g_pszOutlookQualifiedComponents)
		{
			auto b64 = false;
			auto lpszTempPath = mapistub::GetOLMAPI32Path(component);

			if (!lpszTempPath.empty())
			{
				const auto lpszTempVer = std::wstring(MAX_PATH, '\0');
				const auto lpszTempLang = std::wstring(MAX_PATH, '\0');
				DWORD dwValueBuf = MAX_PATH;
				const auto hRes = WC_W32(import::pfnMsiGetFileVersion(
					lpszTempPath.c_str(),
					const_cast<wchar_t*>(lpszTempVer.c_str()),
					&dwValueBuf,
					const_cast<wchar_t*>(lpszTempLang.c_str()),
					&dwValueBuf));
				if (SUCCEEDED(hRes))
				{
					szOut = strings::formatmessage(
						IDS_OUTLOOKVERSIONSTRING, lpszTempPath.c_str(), lpszTempVer.c_str(), lpszTempLang.c_str());
					b64 = Is64BitModule(lpszTempPath);

					szOut += strings::formatmessage(b64 ? IDS_TRUE : IDS_FALSE);
					szOut += L"\n"; // STRING_OK
				}
			}
		}

		return szOut;
	}

	LPCWSTR AppArchitectureToString(const APPX_PACKAGE_ARCHITECTURE a)
	{
		switch (a)
		{
		case APPX_PACKAGE_ARCHITECTURE_X86:
			return L"x86";
		case APPX_PACKAGE_ARCHITECTURE_ARM:
			return L"arm";
		case APPX_PACKAGE_ARCHITECTURE_X64:
			return L"x64";
		case APPX_PACKAGE_ARCHITECTURE_NEUTRAL:
			return L"neutral";
		default:
			return L"?";
		}
	}

	// Most of this function is from https://msdn.microsoft.com/en-us/library/windows/desktop/dn270601(v=vs.85).aspx
	// Return empty strings so the logic that calls this Lookup function can be handled accordingly
	// Only if the API call works do we want to return anything other than an empty string
	std::wstring GetFullName(_In_ LPCWSTR familyName, _In_ const UINT32 filter)
	{
		if (!import::pfnFindPackagesByPackageFamily)
		{
			output::DebugPrint(output::dbgLevel::Generic, L"LookupFamilyName: FindPackagesByPackageFamily not found\n");
			return L"";
		}

		UINT32 count = 0;
		UINT32 length = 0;

		auto rc =
			import::pfnFindPackagesByPackageFamily(familyName, filter, &count, nullptr, &length, nullptr, nullptr);
		if (rc == ERROR_SUCCESS)
		{
			output::DebugPrint(output::dbgLevel::Generic, L"LookupFamilyName: No packages found\n");
			return L"";
		}

		if (rc != ERROR_INSUFFICIENT_BUFFER)
		{
			output::DebugPrint(
				output::dbgLevel::Generic, L"LookupFamilyName: Error %ld in FindPackagesByPackageFamily\n", rc);
			return L"";
		}

		const auto fullNames = static_cast<LPWSTR*>(malloc(count * sizeof(LPWSTR)));
		const auto buffer = static_cast<PWSTR>(malloc(length * sizeof(WCHAR)));

		if (fullNames && buffer)
		{
			rc =
				import::pfnFindPackagesByPackageFamily(familyName, filter, &count, fullNames, &length, buffer, nullptr);
			if (rc != ERROR_SUCCESS)
			{
				output::DebugPrint(
					output::dbgLevel::Generic,
					L"LookupFamilyName: Error %d looking up Full Names from Family Names\n",
					rc);
			}
		}

		std::wstring builds;
		if (count && fullNames)
		{
			// We just take the first one.
			builds = fullNames[0];
		}

		free(buffer);
		free(fullNames);

		return builds;
	}

	std::wstring GetPackageVersion(LPCWSTR fullname)
	{
		if (!import::pfnPackageIdFromFullName) return strings::emptystring;
		UINT32 length = 0;
		auto rc = import::pfnPackageIdFromFullName(fullname, 0, &length, nullptr);
		if (rc == ERROR_SUCCESS)
		{
			output::DebugPrint(output::dbgLevel::Generic, L"GetPackageId: Package not found\n");
			return strings::emptystring;
		}

		if (rc != ERROR_INSUFFICIENT_BUFFER)
		{
			output::DebugPrint(output::dbgLevel::Generic, L"GetPackageId: Error %ld in PackageIdFromFullName\n", rc);
			return strings::emptystring;
		}

		std::wstring build;
		const auto package_id = static_cast<PACKAGE_ID*>(malloc(length));
		if (package_id)
		{
			rc = import::pfnPackageIdFromFullName(fullname, 0, &length, reinterpret_cast<BYTE*>(package_id));
			if (rc != ERROR_SUCCESS)
			{
				output::DebugPrint(
					output::dbgLevel::Generic, L"PackageIdFromFullName: Error %d looking up ID from full name\n", rc);
			}
			else
			{
				build = strings::format(
					L"%d.%d.%d.%d %ws",
					package_id->version.Major,
					package_id->version.Minor,
					package_id->version.Build,
					package_id->version.Revision,
					AppArchitectureToString(static_cast<APPX_PACKAGE_ARCHITECTURE>(package_id->processorArchitecture)));
			}
		}

		free(package_id);
		return build;
	}

	std::wstring GetCentennialVersion()
	{
		// Check for Centennial Office
		const auto familyName = L"Microsoft.Office.Desktop_8wekyb3d8bbwe";

		constexpr UINT32 filter =
			PACKAGE_FILTER_BUNDLE | PACKAGE_FILTER_HEAD | PACKAGE_PROPERTY_BUNDLE | PACKAGE_PROPERTY_RESOURCE;
		auto fullName = GetFullName(familyName, filter);

		if (!fullName.empty())
		{
			return L"Microsoft Store Version: " + GetPackageVersion(fullName.c_str());
		}

		return strings::emptystring;
	}

	std::wstring GetOutlookVersionString()
	{
		auto szVersionString = GetMSIVersion();

		if (szVersionString.empty())
		{
			szVersionString = GetCentennialVersion();
		}

		if (szVersionString.empty())
		{
			szVersionString = strings::loadstring(IDS_NOOUTLOOK);
		}

		return szVersionString;
	}

	bool Is64BitModule(const std::wstring& modulePath)
	{
		if (modulePath.empty()) return false;

		const auto hFile = CreateFileW(
			modulePath.c_str(), GENERIC_READ, FILE_SHARE_READ, nullptr, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, nullptr);

		if (hFile == INVALID_HANDLE_VALUE)
		{
			output::DebugPrint(
				output::dbgLevel::Generic, L"Is64BitModule: Failed to open file %ws\n", modulePath.c_str());
			return false;
		}

		// Read DOS header
		IMAGE_DOS_HEADER dosHeader = {};
		DWORD bytesRead = 0;
		if (!ReadFile(hFile, &dosHeader, sizeof(dosHeader), &bytesRead, nullptr) || bytesRead != sizeof(dosHeader) ||
			dosHeader.e_magic != IMAGE_DOS_SIGNATURE)
		{
			CloseHandle(hFile);
			return false;
		}

		// Seek to PE header
		if (SetFilePointer(hFile, dosHeader.e_lfanew, nullptr, FILE_BEGIN) == INVALID_SET_FILE_POINTER)
		{
			CloseHandle(hFile);
			return false;
		}

		// Read PE signature
		DWORD peSignature = 0;
		if (!ReadFile(hFile, &peSignature, sizeof(peSignature), &bytesRead, nullptr) ||
			bytesRead != sizeof(peSignature) || peSignature != IMAGE_NT_SIGNATURE)
		{
			CloseHandle(hFile);
			return false;
		}

		// Read file header to get machine type
		IMAGE_FILE_HEADER fileHeader = {};
		if (!ReadFile(hFile, &fileHeader, sizeof(fileHeader), &bytesRead, nullptr) || bytesRead != sizeof(fileHeader))
		{
			CloseHandle(hFile);
			return false;
		}

		CloseHandle(hFile);

		// Check machine type to determine if it's 64-bit
		switch (fileHeader.Machine)
		{
		case IMAGE_FILE_MACHINE_AMD64:
		case IMAGE_FILE_MACHINE_IA64:
		case IMAGE_FILE_MACHINE_ARM64:
			return true;
		case IMAGE_FILE_MACHINE_I386:
		case IMAGE_FILE_MACHINE_ARM:
		case IMAGE_FILE_MACHINE_ARMNT:
			return false;
		default:
			// Unknown architecture, assume 32-bit
			return false;
		}
	}
} // namespace version
