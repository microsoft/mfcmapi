#include <core/stdafx.h>
#include <core/mapi/version.h>
#include <core/utility/import.h>
#include <core/mapi/stubutils.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/utility/error.h>
#include <appmodel.h>
#include <AppxPackaging.h>

namespace version
{
	std::wstring GetMSIVersion()
	{
		if (!import::pfnMsiProvideQualifiedComponent || !import::pfnMsiGetFileVersion) return strings::emptystring;

		std::wstring szOut;

		for (auto i = oqcOfficeBegin; i < oqcOfficeEnd; i++)
		{
			auto b64 = false;
			auto lpszTempPath = mapistub::GetOutlookPath(mapistub::g_pszOutlookQualifiedComponents[i], &b64);

			if (!lpszTempPath.empty())
			{
				const auto lpszTempVer = new WCHAR[MAX_PATH];
				const auto lpszTempLang = new WCHAR[MAX_PATH];
				if (lpszTempVer && lpszTempLang)
				{
					DWORD dwValueBuf = MAX_PATH;
					const auto hRes = WC_W32(import::pfnMsiGetFileVersion(
						lpszTempPath.c_str(), lpszTempVer, &dwValueBuf, lpszTempLang, &dwValueBuf));
					if (SUCCEEDED(hRes))
					{
						szOut = strings::formatmessage(
							IDS_OUTLOOKVERSIONSTRING, lpszTempPath.c_str(), lpszTempVer, lpszTempLang);
						szOut += strings::formatmessage(b64 ? IDS_TRUE : IDS_FALSE);
						szOut += L"\n"; // STRING_OK
					}

					delete[] lpszTempLang;
					delete[] lpszTempVer;
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
			output::DebugPrint(DBGGeneric, L"LookupFamilyName: FindPackagesByPackageFamily not found\n");
			return L"";
		}

		UINT32 count = 0;
		UINT32 length = 0;

		auto rc =
			import::pfnFindPackagesByPackageFamily(familyName, filter, &count, nullptr, &length, nullptr, nullptr);
		if (rc == ERROR_SUCCESS)
		{
			output::DebugPrint(DBGGeneric, L"LookupFamilyName: No packages found\n");
			return L"";
		}

		if (rc != ERROR_INSUFFICIENT_BUFFER)
		{
			output::DebugPrint(DBGGeneric, L"LookupFamilyName: Error %ld in FindPackagesByPackageFamily\n", rc);
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
					DBGGeneric, L"LookupFamilyName: Error %d looking up Full Names from Family Names\n", rc);
			}
		}

		std::wstring builds;
		if (count)
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
			output::DebugPrint(DBGGeneric, L"GetPackageId: Package not found\n");
			return strings::emptystring;
		}

		if (rc != ERROR_INSUFFICIENT_BUFFER)
		{
			output::DebugPrint(DBGGeneric, L"GetPackageId: Error %ld in PackageIdFromFullName\n", rc);
			return strings::emptystring;
		}

		std::wstring build;
		const auto package_id = static_cast<PACKAGE_ID*>(malloc(length));
		if (package_id)
		{
			rc = import::pfnPackageIdFromFullName(fullname, 0, &length, reinterpret_cast<BYTE*>(package_id));
			if (rc != ERROR_SUCCESS)
			{
				output::DebugPrint(DBGGeneric, L"PackageIdFromFullName: Error %d looking up ID from full name\n", rc);
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

		const UINT32 filter =
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
} // namespace version
