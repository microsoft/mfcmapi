#include <core/stdafx.h>
#include <Windows.h>
#include <Msi.h>
#include <winreg.h>
// deletethis

// Included for MFCMAPI tracing
#include <core/utility/import.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/utility/error.h>

namespace mapistub
{
	/*
	 * MAPI Stub Utilities
	 *
	 * Public Functions:
	 *
	 * GetPrivateMAPI()
	 * Obtain a handle to the MAPI DLL. This function will load the MAPI DLL
	 * if it hasn't already been loaded
	 *
	 * UnloadPrivateMAPI()
	 * Forces the MAPI DLL to be unloaded. This can cause problems if the code
	 * still has outstanding allocated MAPI memory, or unmatched calls to
	 * MAPIInitialize/MAPIUninitialize
	 *
	 * ForceOutlookMAPI()
	 * Instructs the stub code to always try loading the Outlook version of MAPI
	 * on the system, instead of respecting the system MAPI registration
	 * (HKLM\Software\Clients\Mail). This call must be made prior to any MAPI
	 * function calls.
	 */

	const WCHAR WszKeyNameMailClient[] = L"Software\\Clients\\Mail";
	const WCHAR WszValueNameDllPathEx[] = L"DllPathEx";
	const WCHAR WszValueNameDllPath[] = L"DllPath";

	const WCHAR WszValueNameMSI[] = L"MSIComponentID";
	const WCHAR WszValueNameLCID[] = L"MSIApplicationLCID";

	const WCHAR WszOutlookMapiClientName[] = L"Microsoft Outlook";

	const WCHAR WszMAPISystemPath[] = L"%s\\%s";
	const WCHAR WszMAPISystemDrivePath[] = L"%s%s%s";
	const WCHAR szMAPISystemDrivePath[] = L"%hs%hs%ws";

	static const WCHAR WszOlMAPI32DLL[] = L"olmapi32.dll";
	static const WCHAR WszMSMAPI32DLL[] = L"msmapi32.dll";
	static const WCHAR WszMapi32[] = L"mapi32.dll";
	static const WCHAR WszMapiStub[] = L"mapistub.dll";

	static const CHAR SzFGetComponentPath[] = "FGetComponentPath";

	// Sequence number which is incremented every time we set our MAPI handle which will
	// cause a re-fetch of all stored function pointers
	volatile ULONG g_ulDllSequenceNum = 1;

	// Whether or not we should ignore the system MAPI registration and always try to find
	// Outlook and its MAPI DLLs
	static bool s_fForceOutlookMAPI = false;

	// Whether or not we should ignore the registry and load MAPI from the system directory
	static bool s_fForceSystemMAPI = false;

	static volatile HMODULE g_hinstMAPI = nullptr;
	HMODULE g_hModPstPrx32 = nullptr;

	HMODULE GetMAPIHandle() { return g_hinstMAPI; }

	void SetMAPIHandle(HMODULE hinstMAPI)
	{
		output::DebugPrint(DBGLoadMAPI, L"Enter SetMAPIHandle: hinstMAPI = %p\n", hinstMAPI);
		const HMODULE hinstNULL = nullptr;
		HMODULE hinstToFree = nullptr;

		if (hinstMAPI == nullptr)
		{
			// If we've preloaded pstprx32.dll, unload it before MAPI is unloaded to prevent dependency problems
			if (g_hModPstPrx32)
			{
				FreeLibrary(g_hModPstPrx32);
				g_hModPstPrx32 = nullptr;
			}

			hinstToFree = static_cast<HMODULE>(InterlockedExchangePointer(
				const_cast<PVOID*>(reinterpret_cast<PVOID volatile*>(&g_hinstMAPI)), static_cast<PVOID>(hinstNULL)));
		}
		else
		{
			// Preload pstprx32 to prevent crash when using autodiscover to build a new profile
			if (!g_hModPstPrx32)
			{
				g_hModPstPrx32 = import::LoadFromOLMAPIDir(L"pstprx32.dll"); // STRING_OK
			}

			// Code Analysis gives us a C28112 error when we use InterlockedCompareExchangePointer, so we instead exchange, check and exchange back
			//hinstPrev = (HMODULE)InterlockedCompareExchangePointer(reinterpret_cast<volatile PVOID*>(&g_hinstMAPI), hinstMAPI, hinstNULL);
			const auto hinstPrev = static_cast<HMODULE>(InterlockedExchangePointer(
				const_cast<PVOID*>(reinterpret_cast<PVOID volatile*>(&g_hinstMAPI)), static_cast<PVOID>(hinstMAPI)));
			if (nullptr != hinstPrev)
			{
				(void) InterlockedExchangePointer(
					const_cast<PVOID*>(reinterpret_cast<PVOID volatile*>(&g_hinstMAPI)), static_cast<PVOID>(hinstPrev));
				hinstToFree = hinstMAPI;
			}

			// If we've updated our MAPI handle, any previous addressed fetched via GetProcAddress are invalid, so we
			// have to increment a sequence number to signal that they need to be re-fetched
			InterlockedIncrement(reinterpret_cast<volatile LONG*>(&g_ulDllSequenceNum));
		}

		if (nullptr != hinstToFree)
		{
			FreeLibrary(hinstToFree);
		}

		output::DebugPrint(DBGLoadMAPI, L"Exit SetMAPIHandle\n");
	}

	/*
	 * RegQueryWszExpand
	 * Wrapper for RegQueryValueExW which automatically expands REG_EXPAND_SZ values
	 */
	std::wstring RegQueryWszExpand(HKEY hKey, const std::wstring& lpValueName)
	{
		output::DebugPrint(
			DBGLoadMAPI, L"Enter RegQueryWszExpand: hKey = %p, lpValueName = %ws\n", hKey, lpValueName.c_str());
		DWORD dwType = 0;

		std::wstring ret;
		WCHAR rgchValue[MAX_PATH] = {0};
		DWORD dwSize = sizeof rgchValue;

		const auto dwErr = RegQueryValueExW(
			hKey, lpValueName.c_str(), nullptr, &dwType, reinterpret_cast<LPBYTE>(&rgchValue), &dwSize);

		if (dwErr == ERROR_SUCCESS)
		{
			output::DebugPrint(DBGLoadMAPI, L"RegQueryWszExpand: rgchValue = %ws\n", rgchValue);
			if (dwType == REG_EXPAND_SZ)
			{
				const auto szPath = new WCHAR[MAX_PATH];
				// Expand the strings
				const auto cch = ExpandEnvironmentStringsW(rgchValue, szPath, MAX_PATH);
				if (0 != cch && cch < MAX_PATH)
				{
					output::DebugPrint(DBGLoadMAPI, L"RegQueryWszExpand: rgchValue(expanded) = %ws\n", szPath);
					ret = szPath;
				}

				delete[] szPath;
			}
			else if (dwType == REG_SZ)
			{
				ret = std::wstring(rgchValue);
			}
		}

		output::DebugPrint(DBGLoadMAPI, L"Exit RegQueryWszExpand: dwErr = 0x%08X\n", dwErr);
		return ret;
	}

	/*
	 * GetComponentPath
	 * Wrapper around mapi32.dll->FGetComponentPath which maps an MSI component ID to
	 * a DLL location from the default MAPI client registration values
	 */
	std::wstring GetComponentPath(const std::wstring& szComponent, const std::wstring& szQualifier, bool fInstall)
	{
		output::DebugPrint(
			DBGLoadMAPI,
			L"Enter GetComponentPath: szComponent = %ws, szQualifier = %ws, fInstall = 0x%08X\n",
			szComponent.c_str(),
			szQualifier.c_str(),
			fInstall);
		auto fReturn = false;
		std::wstring path;

		typedef bool(STDAPICALLTYPE * FGetComponentPathType)(LPCSTR, LPSTR, LPSTR, DWORD, bool);

		auto hMapiStub = import::MyLoadLibraryW(WszMapi32);
		if (!hMapiStub) hMapiStub = import::MyLoadLibraryW(WszMapiStub);

		if (hMapiStub)
		{
			const auto pFGetCompPath =
				reinterpret_cast<FGetComponentPathType>(GetProcAddress(hMapiStub, SzFGetComponentPath));

			if (pFGetCompPath)
			{
				CHAR lpszPath[MAX_PATH] = {0};
				const ULONG cchPath = _countof(lpszPath);

				fReturn = pFGetCompPath(
					strings::wstringTostring(szComponent).c_str(),
					const_cast<LPSTR>(strings::wstringTostring(szQualifier).c_str()),
					lpszPath,
					cchPath,
					fInstall);
				if (fReturn) path = strings::LPCSTRToWstring(lpszPath);
				output::DebugPrint(DBGLoadMAPI, L"GetComponentPath: path = %ws\n", path.c_str());
			}

			FreeLibrary(hMapiStub);
		}

		output::DebugPrint(DBGLoadMAPI, L"Exit GetComponentPath: fReturn = 0x%08X\n", fReturn);
		return path;
	}

	/*
	 * GetMailClientFromMSIData
	 * Attempt to locate the MAPI provider DLL via HKLM\Software\Clients\Mail\(provider)\MSIComponentID
	 */
	std::wstring GetMailClientFromMSIData(HKEY hkeyMapiClient)
	{
		output::DebugPrint(DBGLoadMAPI, L"Enter GetMailClientFromMSIData\n");
		if (!hkeyMapiClient) return strings::emptystring;
		WCHAR rgchMSIComponentID[MAX_PATH] = {0};
		WCHAR rgchMSIApplicationLCID[MAX_PATH] = {0};
		DWORD dwType = 0;
		std::wstring szPath;

		DWORD dwSizeComponentID = _countof(rgchMSIComponentID);
		DWORD dwSizeLCID = _countof(rgchMSIApplicationLCID);

		if (ERROR_SUCCESS == RegQueryValueExW(
								 hkeyMapiClient,
								 WszValueNameMSI,
								 nullptr,
								 &dwType,
								 reinterpret_cast<LPBYTE>(&rgchMSIComponentID),
								 &dwSizeComponentID) &&
			ERROR_SUCCESS == RegQueryValueExW(
								 hkeyMapiClient,
								 WszValueNameLCID,
								 nullptr,
								 &dwType,
								 reinterpret_cast<LPBYTE>(&rgchMSIApplicationLCID),
								 &dwSizeLCID))
		{
			const auto componentID = std::wstring(rgchMSIComponentID, dwSizeComponentID);
			const auto applicationID = std::wstring(rgchMSIApplicationLCID, dwSizeLCID);
			szPath = GetComponentPath(componentID, applicationID, false);
		}

		output::DebugPrint(DBGLoadMAPI, L"Exit GetMailClientFromMSIData: szPath = %ws\n", szPath.c_str());
		return szPath;
	}

	/*
	* GetMAPISystemDir
	* Fall back for loading System32\Mapi32.dll if all else fails
	*/
	std::wstring GetMAPISystemDir()
	{
		output::DebugPrint(DBGLoadMAPI, L"Enter GetMAPISystemDir\n");
		auto buf = std::vector<wchar_t>();
		auto copied = DWORD();
		do
		{
			buf.resize(buf.size() + MAX_PATH);
			copied = EC_D(DWORD, ::GetSystemDirectoryW(&buf[0], static_cast<DWORD>(buf.size())));
		} while (copied >= buf.size());

		buf.resize(copied);

		const auto path = std::wstring(buf.begin(), buf.end());
		const auto szDLLPath = path + L"\\" + std::wstring(WszMapi32);

		output::DebugPrint(DBGLoadMAPI, L"Exit GetMAPISystemDir: found %ws\n", szDLLPath.c_str());
		return szDLLPath;
	}

	HKEY GetHKeyMapiClient(const std::wstring& pwzProviderOverride)
	{
		output::DebugPrint(DBGLoadMAPI, L"Enter GetHKeyMapiClient (%ws)\n", pwzProviderOverride.c_str());
		HKEY hMailKey = nullptr;

		// Open HKLM\Software\Clients\Mail
		auto hRes = WC_W32(RegOpenKeyExW(HKEY_LOCAL_MACHINE, WszKeyNameMailClient, 0, KEY_READ, &hMailKey));
		if (FAILED(hRes))
		{
			hMailKey = nullptr;
		}

		// If a specific provider wasn't specified, load the name of the default MAPI provider
		std::wstring defaultClient;
		auto pwzProvider = pwzProviderOverride;
		if (hMailKey && pwzProvider.empty())
		{
			const auto rgchMailClient = new (std::nothrow) WCHAR[MAX_PATH];
			if (rgchMailClient)
			{
				// Get Outlook application path registry value
				DWORD dwSize = MAX_PATH;
				DWORD dwType = 0;
				hRes = WC_W32(RegQueryValueExW(
					hMailKey, nullptr, nullptr, &dwType, reinterpret_cast<LPBYTE>(rgchMailClient), &dwSize));
				if (SUCCEEDED(hRes))
				{
					defaultClient = rgchMailClient;
					output::DebugPrint(
						DBGLoadMAPI,
						L"GetHKeyMapiClient: HKLM\\%ws = %ws\n",
						WszKeyNameMailClient,
						defaultClient.c_str());
				}

				delete[] rgchMailClient;
			}
		}

		if (pwzProvider.empty()) pwzProvider = defaultClient;

		HKEY hkeyMapiClient = nullptr;
		if (hMailKey && !pwzProvider.empty())
		{
			output::DebugPrint(DBGLoadMAPI, L"GetHKeyMapiClient: pwzProvider = %ws\n", pwzProvider.c_str());
			hRes = WC_W32(RegOpenKeyExW(hMailKey, pwzProvider.c_str(), 0, KEY_READ, &hkeyMapiClient));
			if (FAILED(hRes))
			{
				hkeyMapiClient = nullptr;
			}
		}

		output::DebugPrint(
			DBGLoadMAPI, L"Exit GetHKeyMapiClient.hkeyMapiClient found (%ws)\n", hkeyMapiClient ? L"true" : L"false");

		if (hMailKey) RegCloseKey(hMailKey);
		return hkeyMapiClient;
	}

	// Looks up Outlook's path given its qualified component guid
	std::wstring GetOutlookPath(_In_ const std::wstring& szCategory, _Out_opt_ bool* lpb64)
	{
		output::DebugPrint(DBGLoadMAPI, L"Enter GetOutlookPath: szCategory = %ws\n", szCategory.c_str());
		DWORD dwValueBuf = 0;
		std::wstring path;

		if (lpb64) *lpb64 = false;

		auto hRes = WC_W32(import::pfnMsiProvideQualifiedComponent(
			szCategory.c_str(),
			L"outlook.x64.exe", // STRING_OK
			static_cast<DWORD>(INSTALLMODE_DEFAULT),
			nullptr,
			&dwValueBuf));
		if (SUCCEEDED(hRes))
		{
			if (lpb64) *lpb64 = true;
		}
		else
		{
			hRes = WC_W32(import::pfnMsiProvideQualifiedComponent(
				szCategory.c_str(),
				L"outlook.exe", // STRING_OK
				static_cast<DWORD>(INSTALLMODE_DEFAULT),
				nullptr,
				&dwValueBuf));
		}

		if (SUCCEEDED(hRes))
		{
			dwValueBuf += 1;
			const auto lpszTempPath = new (std::nothrow) WCHAR[dwValueBuf];

			if (lpszTempPath != nullptr)
			{
				hRes = WC_W32(import::pfnMsiProvideQualifiedComponent(
					szCategory.c_str(),
					L"outlook.x64.exe", // STRING_OK
					static_cast<DWORD>(INSTALLMODE_DEFAULT),
					lpszTempPath,
					&dwValueBuf));
				if (FAILED(hRes))
				{
					hRes = WC_W32(import::pfnMsiProvideQualifiedComponent(
						szCategory.c_str(),
						L"outlook.exe", // STRING_OK
						static_cast<DWORD>(INSTALLMODE_DEFAULT),
						lpszTempPath,
						&dwValueBuf));
				}

				if (SUCCEEDED(hRes))
				{
					path = lpszTempPath;
					output::DebugPrint(DBGLoadMAPI, L"Exit GetOutlookPath: Path = %ws\n", path.c_str());
				}

				delete[] lpszTempPath;
			}
		}

		if (path.empty())
		{
			output::DebugPrint(DBGLoadMAPI, L"Exit GetOutlookPath: nothing found\n");
		}

		return path;
	}

	WCHAR g_pszOutlookQualifiedComponents[][MAX_PATH] = {
		L"{5812C571-53F0-4467-BEFA-0A4F47A9437C}", // O16_CATEGORY_GUID_CORE_OFFICE (retail) // STRING_OK
		L"{E83B4360-C208-4325-9504-0D23003A74A5}", // O15_CATEGORY_GUID_CORE_OFFICE (retail) // STRING_OK
		L"{1E77DE88-BCAB-4C37-B9E5-073AF52DFD7A}", // O14_CATEGORY_GUID_CORE_OFFICE (retail) // STRING_OK
		L"{24AAE126-0911-478F-A019-07B875EB9996}", // O12_CATEGORY_GUID_CORE_OFFICE (retail) // STRING_OK
		L"{BC174BAD-2F53-4855-A1D5-0D575C19B1EA}", // O11_CATEGORY_GUID_CORE_OFFICE (retail) // STRING_OK
		L"{BC174BAD-2F53-4855-A1D5-1D575C19B1EA}", // O11_CATEGORY_GUID_CORE_OFFICE (debug) // STRING_OK
	};

	std::wstring GetInstalledOutlookMAPI(int iOutlook)
	{
		output::DebugPrint(DBGLoadMAPI, L"Enter GetInstalledOutlookMAPI(%d)\n", iOutlook);

		if (!import::pfnMsiProvideQualifiedComponent || !import::pfnMsiGetFileVersion) return strings::emptystring;

		auto lpszTempPath = GetOutlookPath(g_pszOutlookQualifiedComponents[iOutlook], nullptr);

		if (!lpszTempPath.empty())
		{
			WCHAR szDrive[_MAX_DRIVE] = {0};
			WCHAR szOutlookPath[MAX_PATH] = {0};
			const auto hRes = WC_W32(_wsplitpath_s(
				lpszTempPath.c_str(), szDrive, _MAX_DRIVE, szOutlookPath, MAX_PATH, nullptr, NULL, nullptr, NULL));

			if (SUCCEEDED(hRes))
			{
				auto szPath = std::wstring(szDrive) + std::wstring(szOutlookPath) + WszOlMAPI32DLL;

				output::DebugPrint(DBGLoadMAPI, L"GetInstalledOutlookMAPI: found %ws\n", szPath.c_str());
				return szPath;
			}
		}

		output::DebugPrint(DBGLoadMAPI, L"Exit GetInstalledOutlookMAPI: found nothing\n");
		return strings::emptystring;
	}

	std::vector<std::wstring> GetInstalledOutlookMAPI()
	{
		output::DebugPrint(DBGLoadMAPI, L"Enter GetInstalledOutlookMAPI\n");
		auto paths = std::vector<std::wstring>();
		if (!import::pfnMsiProvideQualifiedComponent || !import::pfnMsiGetFileVersion) return paths;

		for (auto iCurrentOutlook = oqcOfficeBegin; iCurrentOutlook < oqcOfficeEnd; iCurrentOutlook++)
		{
			auto szPath = GetInstalledOutlookMAPI(iCurrentOutlook);
			if (!szPath.empty()) paths.push_back(szPath);
		}

		output::DebugPrint(DBGLoadMAPI, L"Exit GetInstalledOutlookMAPI: found nothing\n");
		return paths;
	}

	std::vector<std::wstring> GetMAPIPaths()
	{
		auto paths = std::vector<std::wstring>();
		std::wstring szPath;
		if (s_fForceSystemMAPI)
		{
			szPath = GetMAPISystemDir();
			if (!szPath.empty()) paths.push_back(szPath);
			return paths;
		}

		auto hkeyMapiClient = HKEY{};
		if (s_fForceOutlookMAPI)
			hkeyMapiClient = GetHKeyMapiClient(WszOutlookMapiClientName);
		else
			hkeyMapiClient = GetHKeyMapiClient(strings::emptystring);

		szPath = RegQueryWszExpand(hkeyMapiClient, WszValueNameDllPathEx);
		if (!szPath.empty()) paths.push_back(szPath);

		auto outlookPaths = GetInstalledOutlookMAPI();
		paths.insert(end(paths), std::begin(outlookPaths), std::end(outlookPaths));

		szPath = RegQueryWszExpand(hkeyMapiClient, WszValueNameDllPath);
		if (!szPath.empty()) paths.push_back(szPath);

		szPath = GetMailClientFromMSIData(hkeyMapiClient);
		if (!szPath.empty()) paths.push_back(szPath);

		if (!s_fForceOutlookMAPI)
		{
			szPath = GetMAPISystemDir();
			if (!szPath.empty()) paths.push_back(szPath);
		}

		if (hkeyMapiClient) RegCloseKey(hkeyMapiClient);
		return paths;
	}

	HMODULE GetDefaultMapiHandle()
	{
		output::DebugPrint(DBGLoadMAPI, L"Enter GetDefaultMapiHandle\n");
		HMODULE hinstMapi = nullptr;

		auto paths = GetMAPIPaths();
		for (const auto& szPath : paths)
		{
			output::DebugPrint(DBGLoadMAPI, L"Trying %ws\n", szPath.c_str());
			hinstMapi = import::MyLoadLibraryW(szPath);
			if (hinstMapi) break;
		}

		output::DebugPrint(DBGLoadMAPI, L"Exit GetDefaultMapiHandle: hinstMapi = %p\n", hinstMapi);
		return hinstMapi;
	}

	/*------------------------------------------------------------------------------
	 Attach to wzMapiDll(olmapi32.dll/msmapi32.dll) if it is already loaded in the
	 current process.
	 ------------------------------------------------------------------------------*/
	HMODULE AttachToMAPIDll(const WCHAR* wzMapiDll)
	{
		output::DebugPrint(DBGLoadMAPI, L"Enter AttachToMAPIDll: wzMapiDll = %ws\n", wzMapiDll);
		HMODULE hinstPrivateMAPI = nullptr;
		import::MyGetModuleHandleExW(0UL, wzMapiDll, &hinstPrivateMAPI);
		output::DebugPrint(DBGLoadMAPI, L"Exit AttachToMAPIDll: hinstPrivateMAPI = %p\n", hinstPrivateMAPI);
		return hinstPrivateMAPI;
	}

	void UnloadPrivateMAPI()
	{
		output::DebugPrint(DBGLoadMAPI, L"Enter UnloadPrivateMAPI\n");
		const auto hinstPrivateMAPI = GetMAPIHandle();
		if (nullptr != hinstPrivateMAPI)
		{
			SetMAPIHandle(nullptr);
		}

		output::DebugPrint(DBGLoadMAPI, L"Exit UnloadPrivateMAPI\n");
	}

	void ForceOutlookMAPI(bool fForce)
	{
		output::DebugPrint(DBGLoadMAPI, L"ForceOutlookMAPI: fForce = 0x%08X\n", fForce);
		s_fForceOutlookMAPI = fForce;
	}

	void ForceSystemMAPI(bool fForce)
	{
		output::DebugPrint(DBGLoadMAPI, L"ForceSystemMAPI: fForce = 0x%08X\n", fForce);
		s_fForceSystemMAPI = fForce;
	}

	HMODULE GetPrivateMAPI()
	{
		output::DebugPrint(DBGLoadMAPI, L"Enter GetPrivateMAPI\n");
		auto hinstPrivateMAPI = GetMAPIHandle();

		if (nullptr == hinstPrivateMAPI)
		{
			// First, try to attach to olmapi32.dll if it's loaded in the process
			hinstPrivateMAPI = AttachToMAPIDll(WszOlMAPI32DLL);

			// If that fails try msmapi32.dll, for Outlook 11 and below
			// Only try this in the static lib, otherwise msmapi32.dll will attach to itself.
			if (nullptr == hinstPrivateMAPI)
			{
				hinstPrivateMAPI = AttachToMAPIDll(WszMSMAPI32DLL);
			}

			// If MAPI isn't loaded in the process yet, then find the path to the DLL and
			// load it manually.
			if (nullptr == hinstPrivateMAPI)
			{
				hinstPrivateMAPI = GetDefaultMapiHandle();
			}

			if (nullptr != hinstPrivateMAPI)
			{
				SetMAPIHandle(hinstPrivateMAPI);
			}

			// Reason - if for any reason there is an instance already loaded, SetMAPIHandle()
			// will free the new one and reuse the old one
			// So we fetch the instance from the global again
			output::DebugPrint(DBGLoadMAPI, L"Exit GetPrivateMAPI: Returning GetMAPIHandle()\n");
			return GetMAPIHandle();
		}

		output::DebugPrint(DBGLoadMAPI, L"Exit GetPrivateMAPI, hinstPrivateMAPI = %p\n", hinstPrivateMAPI);
		return hinstPrivateMAPI;
	}
} // namespace mapistub
