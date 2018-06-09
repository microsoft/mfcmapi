#include <StdAfx.h>
#include <MAPI/Version.h>
#include <ImportProcs.h>
#include <MAPI/StubUtils.h>
#include <appmodel.h>

namespace version
{
	std::wstring GetMSIVersion()
	{
		auto hRes = S_OK;

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
					UINT ret = 0;
					DWORD dwValueBuf = MAX_PATH;
					WC_D(ret, import::pfnMsiGetFileVersion(lpszTempPath.c_str(),
						lpszTempVer,
						&dwValueBuf,
						lpszTempLang,
						&dwValueBuf));
					if (ERROR_SUCCESS == ret)
					{
						szOut = strings::formatmessage(IDS_OUTLOOKVERSIONSTRING, lpszTempPath.c_str(), lpszTempVer, lpszTempLang);
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

	LPCWSTR PackagePropertiesToString(const UINT32 properties)
	{
		const UINT32 PACKAGE_PROPERTY_APPLICATION = 0;
		const auto mask = PACKAGE_PROPERTY_APPLICATION | PACKAGE_PROPERTY_FRAMEWORK |
			PACKAGE_PROPERTY_RESOURCE | PACKAGE_PROPERTY_BUNDLE;
		switch (properties & mask)
		{
		case PACKAGE_PROPERTY_APPLICATION: return L"Application";
		case PACKAGE_PROPERTY_FRAMEWORK:   return L"Framework";
		case PACKAGE_PROPERTY_RESOURCE:    return L"Resource";
		case PACKAGE_PROPERTY_BUNDLE:      return L"Bundle";
		default:                           return L"?";
		}
	}

	// Most of this function is from https://msdn.microsoft.com/en-us/library/windows/desktop/dn270601(v=vs.85).aspx
	// Return empty strings so the logic that calls this Lookup function can be handled accordingly
	// Only if the API call works do we want to return anything other than an empty string
	std::wstring LookupFamilyName(_In_ LPCWSTR familyName, _In_ const UINT32 filter)
	{
		if (!import::pfnFindPackagesByPackageFamily)
		{
			output::DebugPrint(DBGGeneric, L"LookupFamilyName: FindPackagesByPackageFamily not found\n");
			return L"";
		}

		UINT32 count = 0;
		UINT32 length = 0;

		auto rc = import::pfnFindPackagesByPackageFamily(familyName, filter, &count, nullptr, &length, nullptr, nullptr);
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

		const auto fullNames = static_cast<LPWSTR *>(malloc(count * sizeof(LPWSTR)));
		const auto buffer = static_cast<PWSTR>(malloc(length * sizeof(WCHAR)));

		if (fullNames && buffer)
		{
			rc = import::pfnFindPackagesByPackageFamily(familyName, filter, &count, fullNames, &length, buffer, nullptr);
			if (rc != ERROR_SUCCESS)
			{
				output::DebugPrint(DBGGeneric, L"LookupFamilyName: Error %d looking up Full Names from Family Names\n", rc);
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

	std::wstring GetCentennialVersion()
	{
		// If the string is empty, check for Centennial office first
		//const auto familyName = L"Microsoft.Office.Desktop_8wekyb3d8bbwe";
		const auto familyName = L"Microsoft.MicrosoftOfficeHub_8wekyb3d8bbwe";

		const UINT32 filter = PACKAGE_FILTER_BUNDLE | PACKAGE_FILTER_HEAD | PACKAGE_PROPERTY_BUNDLE | PACKAGE_PROPERTY_RESOURCE;
		auto packageName = LookupFamilyName(familyName, filter);

		if (!packageName.empty())
		{
			// Lookup should return something like this:
			// Microsoft.Office.Desktop_16010.10222.20010.0_x86__8wekyb3d8bbwe

			const auto first = packageName.find('_');
			const auto last = packageName.find_last_of('_');

			// Trim the string between the first and last underscore character
			auto buildNum = packageName.substr(first, last - first);

			// We should be left with _16010.10222.20010.0_x86_
			// Replace the underscores with spaces
			std::replace(buildNum.begin(), buildNum.end(), '_', ' ');

			// We now should have the build num and bitness
			return L"Microsoft Store Version: " + buildNum;
		}

		return strings::emptystring;
	}

	std::wstring GetOutlookVersionString()
	{
		auto szVersionString = GetMSIVersion();

		//		if (szVersionString.empty())
		{
			szVersionString = GetCentennialVersion();
		}

		if (szVersionString.empty())
		{
			szVersionString = strings::loadstring(IDS_NOOUTLOOK);
		}

		return szVersionString;
	}
}
