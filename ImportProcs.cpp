#include <StdAfx.h>
#include <ImportProcs.h>
#include <Interpret/String.h>
#include <MAPI/StubUtils.h>

namespace import
{
	HMODULE hModAclui = nullptr;
	HMODULE hModOle32 = nullptr;
	HMODULE hModUxTheme = nullptr;
	HMODULE hModInetComm = nullptr;
	HMODULE hModMSI = nullptr;
	HMODULE hModKernel32 = nullptr;
	HMODULE hModShell32 = nullptr;

	typedef HTHEME(STDMETHODCALLTYPE CLOSETHEMEDATA)(HTHEME hTheme);
	typedef CLOSETHEMEDATA* LPCLOSETHEMEDATA;

	typedef bool(WINAPI HEAPSETINFORMATION)(
		HANDLE HeapHandle,
		HEAP_INFORMATION_CLASS HeapInformationClass,
		PVOID HeapInformation,
		SIZE_T HeapInformationLength);
	typedef HEAPSETINFORMATION* LPHEAPSETINFORMATION;

	typedef bool(WINAPI GETMODULEHANDLEEXW)(DWORD dwFlags, LPCWSTR lpModuleName, HMODULE* phModule);
	typedef GETMODULEHANDLEEXW* LPGETMODULEHANDLEEXW;

	typedef HRESULT(STDMETHODCALLTYPE MIMEOLEGETCODEPAGECHARSET)(
		CODEPAGEID cpiCodePage,
		CHARSETTYPE ctCsetType,
		LPHCHARSET phCharset);
	typedef MIMEOLEGETCODEPAGECHARSET* LPMIMEOLEGETCODEPAGECHARSET;

	typedef UINT(WINAPI MSIPROVIDECOMPONENTW)(
		LPCWSTR szProduct,
		LPCWSTR szFeature,
		LPCWSTR szComponent,
		DWORD dwInstallMode,
		LPWSTR lpPathBuf,
		LPDWORD pcchPathBuf);
	typedef MSIPROVIDECOMPONENTW FAR* LPMSIPROVIDECOMPONENTW;

	typedef UINT(WINAPI MSIPROVIDEQUALIFIEDCOMPONENTW)(
		LPCWSTR szCategory,
		LPCWSTR szQualifier,
		DWORD dwInstallMode,
		LPWSTR lpPathBuf,
		LPDWORD pcchPathBuf);
	typedef MSIPROVIDEQUALIFIEDCOMPONENTW FAR* LPMSIPROVIDEQUALIFIEDCOMPONENTW;

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
	LPMSIPROVIDEQUALIFIEDCOMPONENT pfnMsiProvideQualifiedComponent = nullptr;
	LPMSIGETFILEVERSION pfnMsiGetFileVersion = nullptr;
	LPMSIPROVIDECOMPONENTW pfnMsiProvideComponentW = nullptr;
	LPMSIPROVIDEQUALIFIEDCOMPONENTW pfnMsiProvideQualifiedComponentW = nullptr;

	// From kernel32.dll
	LPHEAPSETINFORMATION pfnHeapSetInformation = nullptr;
	LPGETMODULEHANDLEEXW pfnGetModuleHandleExW = nullptr;
	LPFINDPACKAGESBYPACKAGEFAMILY pfnFindPackagesByPackageFamily = nullptr;
	LPPACKAGEIDFROMFULLNAME pfnPackageIdFromFullName = nullptr;

	// Exists to allow some logging
	_Check_return_ HMODULE MyLoadLibraryW(_In_ const std::wstring& lpszLibFileName)
	{
		HMODULE hMod = nullptr;
		auto hRes = S_OK;
		output::DebugPrint(DBGLoadLibrary, L"MyLoadLibraryW - loading \"%ws\"\n", lpszLibFileName.c_str());
		WC_D(hMod, LoadLibraryW(lpszLibFileName.c_str()));
		if (hMod)
		{
			output::DebugPrint(
				DBGLoadLibrary, L"MyLoadLibraryW - \"%ws\" loaded at %p\n", lpszLibFileName.c_str(), hMod);
		}
		else
		{
			output::DebugPrint(DBGLoadLibrary, L"MyLoadLibraryW - \"%ws\" failed to load\n", lpszLibFileName.c_str());
		}

		return hMod;
	}

	// Loads szModule at the handle given by lphModule, then looks for szEntryPoint.
	// Will not load a module or entry point twice
	void LoadProc(const std::wstring& szModule, HMODULE* lphModule, LPCSTR szEntryPoint, FARPROC* lpfn)
	{
		if (!szEntryPoint || !lpfn || !lphModule) return;
		if (*lpfn) return;
		if (!*lphModule && !szModule.empty())
		{
			*lphModule = LoadFromSystemDir(szModule);
		}

		if (!*lphModule) return;

		auto hRes = S_OK;
		WC_D(*lpfn, GetProcAddress(*lphModule, szEntryPoint));
	}

	_Check_return_ HMODULE LoadFromSystemDir(_In_ const std::wstring& szDLLName)
	{
		if (szDLLName.empty()) return nullptr;

		auto hRes = S_OK;
		HMODULE hModRet = nullptr;
		std::wstring szDLLPath;
		UINT uiRet = NULL;

		static WCHAR szSystemDir[MAX_PATH] = {0};
		static auto bSystemDirLoaded = false;

		output::DebugPrint(DBGLoadLibrary, L"LoadFromSystemDir - loading \"%ws\"\n", szDLLName.c_str());

		if (!bSystemDirLoaded)
		{
			WC_D(uiRet, GetSystemDirectoryW(szSystemDir, MAX_PATH));
			bSystemDirLoaded = true;
		}

		szDLLPath = std::wstring(szSystemDir) + L"\\" + szDLLName;
		output::DebugPrint(DBGLoadLibrary, L"LoadFromSystemDir - loading from \"%ws\"\n", szDLLPath.c_str());
		hModRet = MyLoadLibraryW(szDLLPath);

		return hModRet;
	}

	_Check_return_ HMODULE LoadFromOLMAPIDir(_In_ const std::wstring& szDLLName)
	{
		HMODULE hModRet = nullptr;

		output::DebugPrint(DBGLoadLibrary, L"LoadFromOLMAPIDir - loading \"%ws\"\n", szDLLName.c_str());

		for (auto i = oqcOfficeBegin; i < oqcOfficeEnd; i++)
		{
			auto szOutlookMAPIPath = mapistub::GetInstalledOutlookMAPI(i);
			if (!szOutlookMAPIPath.empty())
			{
				auto hRes = S_OK;
				UINT ret = 0;
				WCHAR szDrive[_MAX_DRIVE] = {0};
				WCHAR szMAPIPath[MAX_PATH] = {0};
				WC_D(
					ret,
					_wsplitpath_s(
						szOutlookMAPIPath.c_str(),
						szDrive,
						_MAX_DRIVE,
						szMAPIPath,
						MAX_PATH,
						nullptr,
						NULL,
						nullptr,
						NULL));

				if (SUCCEEDED(hRes))
				{
					auto szFullPath = std::wstring(szDrive) + std::wstring(szMAPIPath) + szDLLName;

					output::DebugPrint(
						DBGLoadLibrary, L"LoadFromOLMAPIDir - loading from \"%ws\"\n", szFullPath.c_str());
					WC_D(hModRet, MyLoadLibraryW(szFullPath));
				}
			}

			if (hModRet) break;
		}

		return hModRet;
	}

	void ImportProcs()
	{
		// clang-format off
		LoadProc(L"aclui.dll", &hModAclui, "EditSecurity", reinterpret_cast<FARPROC*>(&pfnEditSecurity)); // STRING_OK;
		LoadProc(L"ole32.dll", &hModOle32, "StgCreateStorageEx", reinterpret_cast<FARPROC*>(&pfnStgCreateStorageEx)); // STRING_OK;
		LoadProc(L"uxtheme.dll", &hModUxTheme, "OpenThemeData", reinterpret_cast<FARPROC*>(&pfnOpenThemeData)); // STRING_OK;
		LoadProc(L"uxtheme.dll", &hModUxTheme, "CloseThemeData", reinterpret_cast<FARPROC*>(&pfnCloseThemeData)); // STRING_OK;
		LoadProc(L"uxtheme.dll", &hModUxTheme, "GetThemeMargins", reinterpret_cast<FARPROC*>(&pfnGetThemeMargins)); // STRING_OK;
		LoadProc(L"uxtheme.dll", &hModUxTheme, "SetWindowTheme", reinterpret_cast<FARPROC*>(&pfnSetWindowTheme)); // STRING_OK;
		LoadProc(L"uxtheme.dll", &hModUxTheme, "GetThemeSysSize", reinterpret_cast<FARPROC*>(&pfnGetThemeSysSize)); // STRING_OK;
		LoadProc(L"msi.dll", &hModMSI, "MsiGetFileVersionW", reinterpret_cast<FARPROC*>(&pfnMsiGetFileVersion)); // STRING_OK;
		LoadProc(L"msi.dll", &hModMSI, "MsiProvideQualifiedComponentW", reinterpret_cast<FARPROC*>(&pfnMsiProvideQualifiedComponent)); // STRING_OK;
		LoadProc(L"shell32.dll", &hModShell32, "SHGetPropertyStoreForWindow", reinterpret_cast<FARPROC*>(&pfnSHGetPropertyStoreForWindow)); // STRING_OK;
		LoadProc(L"kernel32.dll", &hModKernel32, "FindPackagesByPackageFamily", reinterpret_cast<FARPROC*>(&pfnFindPackagesByPackageFamily)); // STRING_OK;
		LoadProc(L"kernel32.dll", &hModKernel32, "PackageIdFromFullName", reinterpret_cast<FARPROC*>(&pfnPackageIdFromFullName)); // STRING_OK;
		// clang-format on
	}

	// Opens the mail key for the specified MAPI client, such as 'Microsoft Outlook' or 'ExchangeMAPI'
	// Pass empty string to open the mail key for the default MAPI client
	_Check_return_ HKEY GetMailKey(_In_ const std::wstring& szClient)
	{
		std::wstring lpszClient = L"Default";
		if (!szClient.empty()) lpszClient = szClient;
		output::DebugPrint(DBGLoadLibrary, L"Enter GetMailKey(%ws)\n", lpszClient.c_str());

		// If szClient is empty, we need to read the name of the default MAPI client
		if (szClient.empty())
		{
			HKEY hDefaultMailKey = nullptr;
			WC_W32S(RegOpenKeyExW(
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
					output::DebugPrint(DBGLoadLibrary, L"Default MAPI = %ws\n", lpszClient.c_str());
				}

				EC_W32S(RegCloseKey(hDefaultMailKey));
			}
		}

		HKEY hMailKey = nullptr;
		if (!szClient.empty())
		{
			auto szMailKey = std::wstring(L"Software\\Clients\\Mail\\") + szClient; // STRING_OK

			WC_W32S(RegOpenKeyExW(HKEY_LOCAL_MACHINE, szMailKey.c_str(), NULL, KEY_READ, &hMailKey));
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
		output::DebugPrint(DBGLoadLibrary, L"GetMapiMsiIds(%ws)\n", szClient.c_str());

		const auto hKey = GetMailKey(szClient);
		if (hKey)
		{
			lpszComponentID = registry::ReadStringFromRegistry(hKey, L"MSIComponentID"); // STRING_OK
			output::DebugPrint(
				DBGLoadLibrary,
				L"MSIComponentID = %ws\n",
				!lpszComponentID.empty() ? lpszComponentID.c_str() : L"<not found>");

			lpszAppLCID = registry::ReadStringFromRegistry(hKey, L"MSIApplicationLCID"); // STRING_OK
			output::DebugPrint(
				DBGLoadLibrary,
				L"MSIApplicationLCID = %ws\n",
				!lpszAppLCID.empty() ? lpszAppLCID.c_str() : L"<not found>");

			lpszOfficeLCID = registry::ReadStringFromRegistry(hKey, L"MSIOfficeLCID"); // STRING_OK
			output::DebugPrint(
				DBGLoadLibrary,
				L"MSIOfficeLCID = %ws\n",
				!lpszOfficeLCID.empty() ? lpszOfficeLCID.c_str() : L"<not found>");

			EC_W32S(RegCloseKey(hKey));
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
				&hModKernel32,
				"HeapSetInformation",
				reinterpret_cast<FARPROC*>(&pfnHeapSetInformation)); // STRING_OK;
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
				&hModInetComm,
				"MimeOleGetCodePageCharset",
				reinterpret_cast<FARPROC*>(&pfnMimeOleGetCodePageCharset)); // STRING_OK;
		}

		if (pfnMimeOleGetCodePageCharset) return pfnMimeOleGetCodePageCharset(cpiCodePage, ctCsetType, phCharset);
		return MAPI_E_CALL_FAILED;
	}

	STDAPI_(UINT)
	MsiProvideComponentW(
		LPCWSTR szProduct,
		LPCWSTR szFeature,
		LPCWSTR szComponent,
		DWORD dwInstallMode,
		LPWSTR lpPathBuf,
		LPDWORD pcchPathBuf)
	{
		if (!pfnMsiProvideComponentW)
		{
			LoadProc(
				L"msi.dll",
				&hModMSI,
				"MimeOleGetCodePageCharset",
				reinterpret_cast<FARPROC*>(&pfnMsiProvideComponentW)); // STRING_OK;
		}

		if (pfnMsiProvideComponentW)
			return pfnMsiProvideComponentW(szProduct, szFeature, szComponent, dwInstallMode, lpPathBuf, pcchPathBuf);
		return ERROR_NOT_SUPPORTED;
	}

	STDAPI_(UINT)
	MsiProvideQualifiedComponentW(
		LPCWSTR szCategory,
		LPCWSTR szQualifier,
		DWORD dwInstallMode,
		LPWSTR lpPathBuf,
		LPDWORD pcchPathBuf)
	{
		if (!pfnMsiProvideQualifiedComponentW)
		{
			LoadProc(
				L"msi.dll",
				&hModMSI,
				"MsiProvideQualifiedComponentW",
				reinterpret_cast<FARPROC*>(&pfnMsiProvideQualifiedComponentW)); // STRING_OK;
		}

		if (pfnMsiProvideQualifiedComponentW)
			return pfnMsiProvideQualifiedComponentW(szCategory, szQualifier, dwInstallMode, lpPathBuf, pcchPathBuf);
		return ERROR_NOT_SUPPORTED;
	}

	BOOL WINAPI MyGetModuleHandleExW(DWORD dwFlags, LPCWSTR lpModuleName, HMODULE* phModule)
	{
		if (!pfnGetModuleHandleExW)
		{
			LoadProc(
				L"kernel32.dll",
				&hModKernel32,
				"GetModuleHandleExW",
				reinterpret_cast<FARPROC*>(&pfnGetModuleHandleExW)); // STRING_OK;
		}

		if (pfnGetModuleHandleExW) return pfnGetModuleHandleExW(dwFlags, lpModuleName, phModule);
		*phModule = GetModuleHandleW(lpModuleName);
		return *phModule != nullptr;
	}
}