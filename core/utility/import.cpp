#include <core/stdafx.h>
#include <core/utility/import.h>
#include <mapistub/library/stubutils.h>
#include <core/utility/strings.h>
#include <core/utility/file.h>
#include <core/utility/registry.h>
#include <core/utility/output.h>
#include <core/utility/error.h>

namespace import
{
	HMODULE hModAclui = nullptr;
	HMODULE hModOle32 = nullptr;
	HMODULE hModUxTheme = nullptr;
	HMODULE hModInetComm = nullptr;
	HMODULE hModShell32 = nullptr;
	HMODULE hModCrypt32 = nullptr;

	typedef HTHEME(STDMETHODCALLTYPE CLOSETHEMEDATA)(HTHEME hTheme);
	typedef CLOSETHEMEDATA* LPCLOSETHEMEDATA;

	typedef bool(WINAPI HEAPSETINFORMATION)(
		HANDLE HeapHandle,
		HEAP_INFORMATION_CLASS HeapInformationClass,
		PVOID HeapInformation,
		SIZE_T HeapInformationLength);
	typedef HEAPSETINFORMATION* LPHEAPSETINFORMATION;

	typedef HRESULT(STDMETHODCALLTYPE MIMEOLEGETCODEPAGECHARSET)(
		CODEPAGEID cpiCodePage,
		CHARSETTYPE ctCsetType,
		LPHCHARSET phCharset);
	typedef MIMEOLEGETCODEPAGECHARSET* LPMIMEOLEGETCODEPAGECHARSET;

	LPEDITSECURITY pfnEditSecurity = nullptr;

	LPSTGCREATESTORAGEEX pfnStgCreateStorageEx = nullptr;

	LPOPENTHEMEDATA pfnOpenThemeData = nullptr;
	LPCLOSETHEMEDATA pfnCloseThemeData = nullptr;
	LPGETTHEMEMARGINS pfnGetThemeMargins = nullptr;
	LPSETWINDOWTHEME pfnSetWindowTheme = nullptr;
	LPGETTHEMESYSSIZE pfnGetThemeSysSize = nullptr;
	LPSHGETPROPERTYSTOREFORWINDOW pfnSHGetPropertyStoreForWindow = nullptr;

	// From inetcomm.dll
	LPMIMEOLEGETCODEPAGECHARSET pfnMimeOleGetCodePageCharset = nullptr;

	// From MSI.dll
	LPMSIGETFILEVERSION pfnMsiGetFileVersion = nullptr;

	// From kernel32.dll
	LPHEAPSETINFORMATION pfnHeapSetInformation = nullptr;
	LPFINDPACKAGESBYPACKAGEFAMILY pfnFindPackagesByPackageFamily = nullptr;
	LPPACKAGEIDFROMFULLNAME pfnPackageIdFromFullName = nullptr;

	// From crypt32.dll
	LPCRYPTPROTECTDATA pfnCryptProtectData = nullptr;
	LPCRYPTUNPROTECTDATA pfnCryptUnprotectData = nullptr;

	_Check_return_ HMODULE LoadFromSystemDir(_In_ const std::wstring& szDLLName)
	{
		return mapistub::LoadFromSystemDir(szDLLName);
	}

	// Exists to allow some logging
	_Check_return_ HMODULE MyLoadLibraryW(_In_ const std::wstring& lpszLibFileName)
	{
		output::DebugPrint(
			output::dbgLevel::LoadLibrary, L"MyLoadLibraryW - loading \"%ws\"\n", lpszLibFileName.c_str());
		const auto hMod = WC_D(HMODULE, LoadLibraryW(lpszLibFileName.c_str()));
		if (hMod)
		{
			output::DebugPrint(
				output::dbgLevel::LoadLibrary,
				L"MyLoadLibraryW - \"%ws\" loaded at %p\n",
				lpszLibFileName.c_str(),
				hMod);
		}
		else
		{
			output::DebugPrint(
				output::dbgLevel::LoadLibrary, L"MyLoadLibraryW - \"%ws\" failed to load\n", lpszLibFileName.c_str());
		}

		return hMod;
	}

	template <class T> void LoadProc(_In_ const std::wstring& szModule, HMODULE& hModule, LPCSTR szEntryPoint, T& lpfn)
	{
		FARPROC lpfnFP = {};
		mapistub::LoadProc(szModule, hModule, szEntryPoint, lpfnFP);
		lpfn = reinterpret_cast<T>(lpfnFP);
	}

	void ImportProcs()
	{
		// clang-format off
		LoadProc(L"aclui.dll", hModAclui, "EditSecurity", pfnEditSecurity); // STRING_OK;
		LoadProc(L"ole32.dll", hModOle32, "StgCreateStorageEx", pfnStgCreateStorageEx); // STRING_OK;
		LoadProc(L"uxtheme.dll", hModUxTheme, "OpenThemeData", pfnOpenThemeData); // STRING_OK;
		LoadProc(L"uxtheme.dll", hModUxTheme, "CloseThemeData", pfnCloseThemeData); // STRING_OK;
		LoadProc(L"uxtheme.dll", hModUxTheme, "GetThemeMargins", pfnGetThemeMargins); // STRING_OK;
		LoadProc(L"uxtheme.dll", hModUxTheme, "SetWindowTheme", pfnSetWindowTheme); // STRING_OK;
		LoadProc(L"uxtheme.dll", hModUxTheme, "GetThemeSysSize", pfnGetThemeSysSize); // STRING_OK;
		LoadProc(L"msi.dll", mapistub::hModMSI, "MsiGetFileVersionW", pfnMsiGetFileVersion); // STRING_OK;
		LoadProc(L"shell32.dll", hModShell32, "SHGetPropertyStoreForWindow", pfnSHGetPropertyStoreForWindow); // STRING_OK;
		LoadProc(L"kernel32.dll", mapistub::hModKernel32, "FindPackagesByPackageFamily", pfnFindPackagesByPackageFamily); // STRING_OK;
		LoadProc(L"kernel32.dll", mapistub::hModKernel32, "PackageIdFromFullName", pfnPackageIdFromFullName); // STRING_OK;
		LoadProc(L"crypt32.dll", hModCrypt32, "CryptProtectData", pfnCryptProtectData); // STRING_OK;
		LoadProc(L"crypt32.dll", hModCrypt32, "CryptUnprotectData", pfnCryptUnprotectData ); // STRING_OK;
		// clang-format on
	}

	// Opens the mail key for the specified MAPI client, such as 'Microsoft Outlook' or 'ExchangeMAPI'
	// Pass empty string to open the mail key for the default MAPI client
	_Check_return_ HKEY GetMailKey(_In_ const std::wstring& szClient)
	{
		std::wstring lpszClient = L"Default";
		if (!szClient.empty()) lpszClient = szClient;
		output::DebugPrint(output::dbgLevel::LoadLibrary, L"Enter GetMailKey(%ws)\n", lpszClient.c_str());

		// If szClient is empty, we need to read the name of the default MAPI client
		if (szClient.empty())
		{
			HKEY hDefaultMailKey = nullptr;
			WC_W32_S(RegOpenKeyExW(
				HKEY_LOCAL_MACHINE,
				L"Software\\Clients\\Mail", // STRING_OK
				NULL,
				KEY_READ,
				&hDefaultMailKey));
			if (hDefaultMailKey)
			{
				auto lpszReg = registry::ReadStringFromRegistry(hDefaultMailKey,
																L""); // get the default value
				if (!lpszReg.empty())
				{
					lpszClient = lpszReg;
					output::DebugPrint(output::dbgLevel::LoadLibrary, L"Default MAPI = %ws\n", lpszClient.c_str());
				}

				EC_W32_S(RegCloseKey(hDefaultMailKey));
			}
		}

		HKEY hMailKey = nullptr;
		if (!szClient.empty())
		{
			auto szMailKey = std::wstring(L"Software\\Clients\\Mail\\") + szClient; // STRING_OK

			WC_W32_S(RegOpenKeyExW(HKEY_LOCAL_MACHINE, szMailKey.c_str(), NULL, KEY_READ, &hMailKey));
		}

		return hMailKey;
	}

	// Gets MSI IDs for the specified MAPI client, such as 'Microsoft Outlook' or 'ExchangeMAPI'
	// Pass empty string to get the IDs for the default MAPI client
	void GetMapiMsiIds(
		_In_ const std::wstring& szClient,
		_In_ std::wstring& lpszComponentID,
		_In_ std::wstring& lpszAppLCID,
		_In_ std::wstring& lpszOfficeLCID)
	{
		output::DebugPrint(output::dbgLevel::LoadLibrary, L"GetMapiMsiIds(%ws)\n", szClient.c_str());

		const auto hKey = GetMailKey(szClient);
		if (hKey)
		{
			lpszComponentID = registry::ReadStringFromRegistry(hKey, L"MSIComponentID"); // STRING_OK
			output::DebugPrint(
				output::dbgLevel::LoadLibrary,
				L"MSIComponentID = %ws\n",
				!lpszComponentID.empty() ? lpszComponentID.c_str() : L"<not found>");

			lpszAppLCID = registry::ReadStringFromRegistry(hKey, L"MSIApplicationLCID"); // STRING_OK
			output::DebugPrint(
				output::dbgLevel::LoadLibrary,
				L"MSIApplicationLCID = %ws\n",
				!lpszAppLCID.empty() ? lpszAppLCID.c_str() : L"<not found>");

			lpszOfficeLCID = registry::ReadStringFromRegistry(hKey, L"MSIOfficeLCID"); // STRING_OK
			output::DebugPrint(
				output::dbgLevel::LoadLibrary,
				L"MSIOfficeLCID = %ws\n",
				!lpszOfficeLCID.empty() ? lpszOfficeLCID.c_str() : L"<not found>");

			EC_W32_S(RegCloseKey(hKey));
		}
	}

	std::wstring GetMAPIPath(const std::wstring& szClient)
	{
		std::wstring lpszPath;

		// Find some strings:
		std::wstring szComponentID;
		std::wstring szAppLCID;
		std::wstring szOfficeLCID;

		GetMapiMsiIds(szClient, szComponentID, szAppLCID, szOfficeLCID);

		if (!szComponentID.empty())
		{
			if (!szAppLCID.empty())
			{
				lpszPath = mapistub::GetComponentPath(szComponentID, szAppLCID, false);
			}

			if (lpszPath.empty() && !szOfficeLCID.empty())
			{
				lpszPath = mapistub::GetComponentPath(szComponentID, szOfficeLCID, false);
			}

			if (lpszPath.empty())
			{
				lpszPath = mapistub::GetComponentPath(szComponentID, strings::emptystring, false);
			}
		}

		return lpszPath;
	}

	void WINAPI MyHeapSetInformation(
		_In_opt_ HANDLE HeapHandle,
		_In_ HEAP_INFORMATION_CLASS HeapInformationClass,
		_In_opt_count_(HeapInformationLength) PVOID HeapInformation,
		_In_ SIZE_T HeapInformationLength)
	{
		if (!pfnHeapSetInformation)
		{
			LoadProc(
				L"kernel32.dll",
				mapistub::hModKernel32,
				"HeapSetInformation",
				pfnHeapSetInformation); // STRING_OK;
		}

		if (pfnHeapSetInformation)
			pfnHeapSetInformation(HeapHandle, HeapInformationClass, HeapInformation, HeapInformationLength);
	}

	HRESULT WINAPI MyMimeOleGetCodePageCharset(CODEPAGEID cpiCodePage, CHARSETTYPE ctCsetType, LPHCHARSET phCharset)
	{
		if (!pfnMimeOleGetCodePageCharset)
		{
			LoadProc(
				L"inetcomm.dll",
				hModInetComm,
				"MimeOleGetCodePageCharset",
				pfnMimeOleGetCodePageCharset); // STRING_OK;
		}

		if (pfnMimeOleGetCodePageCharset) return pfnMimeOleGetCodePageCharset(cpiCodePage, ctCsetType, phCharset);
		return MAPI_E_CALL_FAILED;
	}
} // namespace import