#include "stdafx.h"
#include <windows.h>
#include <strsafe.h>
#include <msi.h>
#include <winreg.h>
#include <stdlib.h>

// Included for MFCMAPI tracing
#include "MFCOutput.h"
#include "ImportProcs.h"
#include "String.h"

/*
 * MAPI Stub Utilities
 *
 * Public Functions:
 *
 * GetPrivateMAPI()
 * Obtain a handle to the MAPI DLL. This function will load the MAPI DLL
 * if it hasn't already been loaded
 *
 * UnLoadPrivateMAPI()
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
HMODULE GetPrivateMAPI();
void UnLoadPrivateMAPI();
void ForceOutlookMAPI(bool fForce);

static wstring GetRegisteredMapiClient(wstring pwzProviderOverride, bool bDLL, bool bEx);
vector<wstring> GetInstalledOutlookMAPI();

const WCHAR WszKeyNameMailClient[] = L"Software\\Clients\\Mail";
const WCHAR WszValueNameDllPathEx[] = L"DllPathEx";
const WCHAR WszValueNameDllPath[] = L"DllPath";

const CHAR SzValueNameMSI[] = "MSIComponentID";
const CHAR SzValueNameLCID[] = "MSIApplicationLCID";

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

HMODULE GetMAPIHandle()
{
	return g_hinstMAPI;
}

void SetMAPIHandle(HMODULE hinstMAPI)
{
	DebugPrint(DBGLoadMAPI, L"Enter SetMAPIHandle: hinstMAPI = %p\n", hinstMAPI);
	HMODULE hinstNULL = nullptr;
	HMODULE hinstToFree = nullptr;

	if (hinstMAPI == nullptr)
	{
		// If we've preloaded pstprx32.dll, unload it before MAPI is unloaded to prevent dependency problems
		if (g_hModPstPrx32)
		{
			FreeLibrary(g_hModPstPrx32);
			g_hModPstPrx32 = nullptr;
		}

		hinstToFree = static_cast<HMODULE>(InterlockedExchangePointer((PVOID*)&g_hinstMAPI, static_cast<PVOID>(hinstNULL)));
	}
	else
	{
		// Preload pstprx32 to prevent crash when using autodiscover to build a new profile
		if (!g_hModPstPrx32)
		{
			g_hModPstPrx32 = LoadFromOLMAPIDir(_T("pstprx32.dll")); // STRING_OK
		}

		// Set the value only if the global is NULL
		HMODULE hinstPrev;
		// Code Analysis gives us a C28112 error when we use InterlockedCompareExchangePointer, so we instead exchange, check and exchange back
		//hinstPrev = (HMODULE)InterlockedCompareExchangePointer(reinterpret_cast<volatile PVOID*>(&g_hinstMAPI), hinstMAPI, hinstNULL);
		hinstPrev = static_cast<HMODULE>(InterlockedExchangePointer((PVOID*)&g_hinstMAPI, static_cast<PVOID>(hinstMAPI)));
		if (NULL != hinstPrev)
		{
			(void)InterlockedExchangePointer((PVOID*)&g_hinstMAPI, static_cast<PVOID>(hinstPrev));
			hinstToFree = hinstMAPI;
		}

		// If we've updated our MAPI handle, any previous addressed fetched via GetProcAddress are invalid, so we
		// have to increment a sequence number to signal that they need to be re-fetched
		InterlockedIncrement(reinterpret_cast<volatile LONG*>(&g_ulDllSequenceNum));
	}

	if (NULL != hinstToFree)
	{
		FreeLibrary(hinstToFree);
	}

	DebugPrint(DBGLoadMAPI, L"Exit SetMAPIHandle\n");
}

/*
 * RegQueryWszExpand
 * Wrapper for RegQueryValueExW which automatically expands REG_EXPAND_SZ values
 */
wstring RegQueryWszExpand(HKEY hKey, wstring lpValueName)
{
	DebugPrint(DBGLoadMAPI, L"Enter RegQueryWszExpand: hKey = %p, lpValueName = %ws\n", hKey, lpValueName.c_str());
	DWORD dwType = 0;

	wstring ret;
	WCHAR rgchValue[MAX_PATH] = { 0 };
	DWORD dwSize = sizeof rgchValue;

	auto dwErr = RegQueryValueExW(hKey, lpValueName.c_str(), nullptr, &dwType, reinterpret_cast<LPBYTE>(&rgchValue), &dwSize);

	if (dwErr == ERROR_SUCCESS)
	{
		DebugPrint(DBGLoadMAPI, L"RegQueryWszExpand: rgchValue = %ws\n", rgchValue);
		if (dwType == REG_EXPAND_SZ)
		{
			auto szPath = new WCHAR[MAX_PATH];
			// Expand the strings
			auto cch = ExpandEnvironmentStringsW(rgchValue, szPath, MAX_PATH);
			if (0 != cch && cch < MAX_PATH)
			{
				DebugPrint(DBGLoadMAPI, L"RegQueryWszExpand: rgchValue(expanded) = %ws\n", szPath);
				ret = szPath;
			}

			delete[] szPath;
		}
		else if (dwType == REG_SZ)
		{
			ret = wstring(rgchValue);
		}
	}

	DebugPrint(DBGLoadMAPI, L"Exit RegQueryWszExpand: dwErr = 0x%08X\n", dwErr);
	return ret;
}

/*
 * GetComponentPath
 * Wrapper around mapi32.dll->FGetComponentPath which maps an MSI component ID to
 * a DLL location from the default MAPI client registration values
 */
bool GetComponentPath(LPCSTR szComponent, LPSTR szQualifier, LPSTR szDllPath, DWORD cchBufferSize, bool fInstall)
{
	DebugPrint(DBGLoadMAPI, L"Enter GetComponentPath: szComponent = %hs, szQualifier = %hs, cchBufferSize = 0x%08X, fInstall = 0x%08X\n",
		szComponent, szQualifier, cchBufferSize, fInstall);
	bool fReturn = FALSE;

	typedef bool (STDAPICALLTYPE *FGetComponentPathType)(LPCSTR, LPSTR, LPSTR, DWORD, bool);

	auto hMapiStub = MyLoadLibraryW(WszMapi32);
	if (!hMapiStub)
		hMapiStub = MyLoadLibraryW(WszMapiStub);

	if (hMapiStub)
	{
		auto pFGetCompPath = reinterpret_cast<FGetComponentPathType>(GetProcAddress(hMapiStub, SzFGetComponentPath));

		if (pFGetCompPath)
		{
			fReturn = pFGetCompPath(szComponent, szQualifier, szDllPath, cchBufferSize, fInstall);
			DebugPrint(DBGLoadMAPI, L"GetComponentPath: szDllPath = %hs\n", szDllPath);
		}

		FreeLibrary(hMapiStub);
	}

	DebugPrint(DBGLoadMAPI, L"Exit GetComponentPath: fReturn = 0x%08X\n", fReturn);
	return fReturn;
}

vector<wstring> GetMAPIPaths(bool bBypassRestrictions)
{
	auto paths = vector<wstring>();
	wstring szPath;
	if (s_fForceSystemMAPI && !bBypassRestrictions)
	{
		szPath = GetMAPISystemDir();
		if (!szPath.empty()) paths.push_back(szPath);
		return paths;
	}

	if (bBypassRestrictions)
	{
		auto outlookPaths = GetInstalledOutlookMAPI();
		paths.insert(end(paths), begin(outlookPaths), end(outlookPaths));
	}

	szPath = GetRegisteredMapiClient(WszOutlookMapiClientName, false, false);
	if (!szPath.empty()) paths.push_back(szPath);
	szPath = GetRegisteredMapiClient(WszOutlookMapiClientName, true, true);
	if (!szPath.empty()) paths.push_back(szPath);
	szPath = GetRegisteredMapiClient(WszOutlookMapiClientName, true, false);
	if (!szPath.empty()) paths.push_back(szPath);
	if (bBypassRestrictions || !s_fForceOutlookMAPI)
	{
		szPath = GetMAPISystemDir();
		if (!szPath.empty()) paths.push_back(szPath);
	}

	return paths;
}

/*
 * GetMailClientFromMSIData
 * Attempt to locate the MAPI provider DLL via HKLM\Software\Clients\Mail\(provider)\MSIComponentID
 */
wstring GetMailClientFromMSIData(HKEY hkeyMapiClient)
{
	DebugPrint(DBGLoadMAPI, L"Enter GetMailClientFromMSIData\n");
	CHAR rgchMSIComponentID[MAX_PATH] = { 0 };
	CHAR rgchMSIApplicationLCID[MAX_PATH] = { 0 };
	CHAR rgchComponentPath[MAX_PATH] = { 0 };
	DWORD dwType = 0;
	wstring szPath;

	DWORD dwSizeComponentID = sizeof rgchMSIComponentID;
	DWORD dwSizeLCID = sizeof rgchMSIApplicationLCID;

	if (ERROR_SUCCESS == RegQueryValueExA(hkeyMapiClient, SzValueNameMSI, nullptr, &dwType, reinterpret_cast<LPBYTE>(&rgchMSIComponentID), &dwSizeComponentID) &&
		ERROR_SUCCESS == RegQueryValueExA(hkeyMapiClient, SzValueNameLCID, nullptr, &dwType, reinterpret_cast<LPBYTE>(&rgchMSIApplicationLCID), &dwSizeLCID))
	{
		if (GetComponentPath(rgchMSIComponentID, rgchMSIApplicationLCID, rgchComponentPath, _countof(rgchComponentPath), FALSE))
		{
			szPath = stringTowstring(rgchComponentPath);
		}
	}

	DebugPrint(DBGLoadMAPI, L"Exit GetMailClientFromMSIData: szPath = %ws\n", szPath.c_str());
	return szPath;
}

/*
 * GetMailClientFromDllPath
 * Attempt to locate the MAPI provider DLL via HKLM\Software\Clients\Mail\(provider)\DllPathEx
 */
wstring GetMailClientFromDllPath(HKEY hkeyMapiClient, bool bEx)
{
	DebugPrint(DBGLoadMAPI, L"Enter GetMailClientFromDllPath: hkeyMapiClient = %p, bEx = %d\n", hkeyMapiClient, bEx);
	wstring szPath;

	if (bEx)
	{
		szPath = RegQueryWszExpand(hkeyMapiClient, WszValueNameDllPathEx);
	}
	else
	{
		szPath = RegQueryWszExpand(hkeyMapiClient, WszValueNameDllPath);
	}

	DebugPrint(DBGLoadMAPI, L"Exit GetMailClientFromDllPath: szPath = %ws\n", szPath.c_str());

	return szPath;
}

/*
 * GetRegisteredMapiClient
 * Read the registry to discover the registered MAPI client and attempt to load its MAPI DLL.
 *
 * If wzOverrideProvider is specified, this function will load that MAPI Provider instead of the
 * currently registered provider
 */
wstring GetRegisteredMapiClient(wstring pwzProviderOverride, bool bDLL, bool bEx)
{
	DebugPrint(DBGLoadMAPI, L"Enter GetRegisteredMapiClient\n");
	auto hRes = S_OK;
	auto pwzProvider = pwzProviderOverride;
	HKEY hMailKey = nullptr;
	HKEY hkeyMapiClient = nullptr;

	// Open HKLM\Software\Clients\Mail
	WC_W32(RegOpenKeyExW(HKEY_LOCAL_MACHINE,
		WszKeyNameMailClient,
		0,
		KEY_READ,
		&hMailKey));
	if (FAILED(hRes))
	{
		hMailKey = nullptr;
	}

	// If a specific provider wasn't specified, load the name of the default MAPI provider
	wstring defaultClient;
	if (hMailKey && pwzProvider.empty())
	{
		auto rgchMailClient = new WCHAR[MAX_PATH];
		if (rgchMailClient)
		{
			// Get Outlook application path registry value
			DWORD dwSize = MAX_PATH;
			DWORD dwType = 0;
			WC_W32(RegQueryValueExW(
				hMailKey,
				NULL,
				nullptr,
				&dwType,
				reinterpret_cast<LPBYTE>(rgchMailClient),
				&dwSize));
			if (SUCCEEDED(hRes))
			{
				defaultClient = rgchMailClient;
				DebugPrint(DBGLoadMAPI, L"GetRegisteredMapiClient: HKLM\\%ws = %ws\n", WszKeyNameMailClient, defaultClient.c_str());
			}

			delete[] rgchMailClient;
		}
	}

	if (pwzProvider.empty()) pwzProvider = defaultClient;

	if (hMailKey && !pwzProvider.empty())
	{
		DebugPrint(DBGLoadMAPI, L"GetRegisteredMapiClient: pwzProvider = %ws\n", pwzProvider.c_str());
		WC_W32(RegOpenKeyExW(
			hMailKey,
			pwzProvider.c_str(),
			0,
			KEY_READ,
			&hkeyMapiClient));
		if (FAILED(hRes))
		{
			hkeyMapiClient = nullptr;
		}
	}

	wstring szPath;
	if (hkeyMapiClient)
	{
		if (bDLL)
		{
			szPath = GetMailClientFromDllPath(hkeyMapiClient, bEx);
		}
		else
		{
			szPath = GetMailClientFromMSIData(hkeyMapiClient);
		}

	}

	DebugPrint(DBGLoadMAPI, L"Exit GetRegisteredMapiClient: szPath = %ws\n", szPath.c_str());

	if (hkeyMapiClient) RegCloseKey(hkeyMapiClient);
	if (hMailKey) RegCloseKey(hMailKey);
	return szPath;
}

/*
 * GetMAPISystemDir
 * Fall back for loading System32\Mapi32.dll if all else fails
 */
wstring GetMAPISystemDir()
{
	DebugPrint(DBGLoadMAPI, L"Enter GetMAPISystemDir\n");
	WCHAR szSystemDir[MAX_PATH] = { 0 };

	if (GetSystemDirectoryW(szSystemDir, MAX_PATH))
	{
		auto szDLLPath = wstring(szSystemDir) + L"\\" + wstring(WszMapi32);
		DebugPrint(DBGLoadMAPI, L"GetMAPISystemDir: found %ws\n", szDLLPath.c_str());
		return szDLLPath;
	}

	DebugPrint(DBGLoadMAPI, L"Exit GetMAPISystemDir: found nothing\n");
	return emptystring;
}

wstring GetInstalledOutlookMAPI(int iOutlook)
{
	DebugPrint(DBGLoadMAPI, L"Enter GetInstalledOutlookMAPI(%d)\n", iOutlook);
	auto hRes = S_OK;

	if (!pfnMsiProvideQualifiedComponent || !pfnMsiGetFileVersion) return nullptr;

	UINT ret = 0;

	auto lpszTempPath = GetOutlookPath(g_pszOutlookQualifiedComponents[iOutlook], nullptr);

	if (!lpszTempPath.empty())
	{
		WCHAR szDrive[_MAX_DRIVE] = { 0 };
		WCHAR szOutlookPath[MAX_PATH] = { 0 };
		WC_D(ret, _wsplitpath_s(lpszTempPath.c_str(), szDrive, _MAX_DRIVE, szOutlookPath, MAX_PATH, NULL, NULL, NULL, NULL));

		if (SUCCEEDED(hRes))
		{
			auto szPath = wstring(szDrive) + wstring(szOutlookPath) + WszOlMAPI32DLL;

			DebugPrint(DBGLoadMAPI, L"GetInstalledOutlookMAPI: found %ws\n", szPath.c_str());
			return szPath;
		}
	}

	DebugPrint(DBGLoadMAPI, L"Exit GetInstalledOutlookMAPI: found nothing\n");
	return emptystring;
}

vector<wstring> GetInstalledOutlookMAPI()
{
	DebugPrint(DBGLoadMAPI, L"Enter GetInstalledOutlookMAPI\n");
	auto paths = vector<wstring>();
	if (!pfnMsiProvideQualifiedComponent || !pfnMsiGetFileVersion) return paths;

	for (auto iCurrentOutlook = oqcOfficeBegin; iCurrentOutlook < oqcOfficeEnd; iCurrentOutlook++)
	{
		auto szPath = GetInstalledOutlookMAPI(iCurrentOutlook);
		if (!szPath.empty()) paths.push_back(szPath);
	}

	DebugPrint(DBGLoadMAPI, L"Exit GetInstalledOutlookMAPI: found nothing\n");
	return paths;
}

WCHAR g_pszOutlookQualifiedComponents[][MAX_PATH] = {
 L"{5812C571-53F0-4467-BEFA-0A4F47A9437C}", // O16_CATEGORY_GUID_CORE_OFFICE (retail) // STRING_OK
 L"{E83B4360-C208-4325-9504-0D23003A74A5}", // O15_CATEGORY_GUID_CORE_OFFICE (retail) // STRING_OK
 L"{1E77DE88-BCAB-4C37-B9E5-073AF52DFD7A}", // O14_CATEGORY_GUID_CORE_OFFICE (retail) // STRING_OK
 L"{24AAE126-0911-478F-A019-07B875EB9996}", // O12_CATEGORY_GUID_CORE_OFFICE (retail) // STRING_OK
 L"{BC174BAD-2F53-4855-A1D5-0D575C19B1EA}", // O11_CATEGORY_GUID_CORE_OFFICE (retail) // STRING_OK
 L"{BC174BAD-2F53-4855-A1D5-1D575C19B1EA}", // O11_CATEGORY_GUID_CORE_OFFICE (debug) // STRING_OK
};

// Looks up Outlook's path given its qualified component guid
wstring GetOutlookPath(_In_ wstring szCategory, _Out_opt_ bool* lpb64)
{
	DebugPrint(DBGLoadMAPI, L"Enter GetOutlookPath: szCategory = %ws\n", szCategory.c_str());
	auto hRes = S_OK;
	DWORD dwValueBuf = 0;
	UINT ret = 0;
	wstring path;

	if (lpb64) *lpb64 = false;

	WC_D(ret, pfnMsiProvideQualifiedComponent(
		szCategory.c_str(),
		L"outlook.x64.exe", // STRING_OK
		static_cast<DWORD>(INSTALLMODE_DEFAULT),
		NULL,
		&dwValueBuf));
	if (ERROR_SUCCESS == ret)
	{
		if (lpb64) *lpb64 = true;
	}
	else
	{
		WC_D(ret, pfnMsiProvideQualifiedComponent(
			szCategory.c_str(),
			L"outlook.exe", // STRING_OK
			static_cast<DWORD>(INSTALLMODE_DEFAULT),
			NULL,
			&dwValueBuf));
	}

	if (ERROR_SUCCESS == ret)
	{
		dwValueBuf += 1;
		auto lpszTempPath = new WCHAR[dwValueBuf];

		if (lpszTempPath != nullptr)
		{
			WC_D(ret, pfnMsiProvideQualifiedComponent(
				szCategory.c_str(),
				L"outlook.x64.exe", // STRING_OK
				static_cast<DWORD>(INSTALLMODE_DEFAULT),
				lpszTempPath,
				&dwValueBuf));
			if (ERROR_SUCCESS != ret)
			{
				WC_D(ret, pfnMsiProvideQualifiedComponent(
					szCategory.c_str(),
					L"outlook.exe", // STRING_OK
					static_cast<DWORD>(INSTALLMODE_DEFAULT),
					lpszTempPath,
					&dwValueBuf));
			}

			if (ERROR_SUCCESS == ret)
			{
				path = lpszTempPath;
				DebugPrint(DBGLoadMAPI, L"Exit GetOutlookPath: Path = %ws\n", path.c_str());
			}

			delete[] lpszTempPath;
		}
	}

	if (path.empty())
	{
		DebugPrint(DBGLoadMAPI, L"Exit GetOutlookPath: nothing found\n");
	}

	return path;
}

HMODULE GetDefaultMapiHandle()
{
	DebugPrint(DBGLoadMAPI, L"Enter GetDefaultMapiHandle\n");
	HMODULE hinstMapi = nullptr;

	auto paths = GetMAPIPaths(true);
	for (auto szPath : paths)
	{
		DebugPrint(DBGLoadMAPI, L"Trying %ws\n", szPath.c_str());
		hinstMapi = MyLoadLibraryW(szPath.c_str());
		if (hinstMapi) break;
	}

	DebugPrint(DBGLoadMAPI, L"Exit GetDefaultMapiHandle: hinstMapi = %p\n", hinstMapi);
	return hinstMapi;
}

/*------------------------------------------------------------------------------
 Attach to wzMapiDll(olmapi32.dll/msmapi32.dll) if it is already loaded in the
 current process.
 ------------------------------------------------------------------------------*/
HMODULE AttachToMAPIDll(const WCHAR *wzMapiDll)
{
	DebugPrint(DBGLoadMAPI, L"Enter AttachToMAPIDll: wzMapiDll = %ws\n", wzMapiDll);
	HMODULE hinstPrivateMAPI = nullptr;
	MyGetModuleHandleExW(0UL, wzMapiDll, &hinstPrivateMAPI);
	DebugPrint(DBGLoadMAPI, L"Exit AttachToMAPIDll: hinstPrivateMAPI = %p\n", hinstPrivateMAPI);
	return hinstPrivateMAPI;
}

void UnLoadPrivateMAPI()
{
	DebugPrint(DBGLoadMAPI, L"Enter UnLoadPrivateMAPI\n");
	auto hinstPrivateMAPI = GetMAPIHandle();
	if (NULL != hinstPrivateMAPI)
	{
		SetMAPIHandle(nullptr);
	}

	DebugPrint(DBGLoadMAPI, L"Exit UnLoadPrivateMAPI\n");
}

void ForceOutlookMAPI(bool fForce)
{
	DebugPrint(DBGLoadMAPI, L"ForceOutlookMAPI: fForce = 0x%08X\n", fForce);
	s_fForceOutlookMAPI = fForce;
}

void ForceSystemMAPI(bool fForce)
{
	DebugPrint(DBGLoadMAPI, L"ForceSystemMAPI: fForce = 0x%08X\n", fForce);
	s_fForceSystemMAPI = fForce;
}

HMODULE GetPrivateMAPI()
{
	DebugPrint(DBGLoadMAPI, L"Enter GetPrivateMAPI\n");
	auto hinstPrivateMAPI = GetMAPIHandle();

	if (NULL == hinstPrivateMAPI)
	{
		// First, try to attach to olmapi32.dll if it's loaded in the process
		hinstPrivateMAPI = AttachToMAPIDll(WszOlMAPI32DLL);

		// If that fails try msmapi32.dll, for Outlook 11 and below
		// Only try this in the static lib, otherwise msmapi32.dll will attach to itself.
		if (NULL == hinstPrivateMAPI)
		{
			hinstPrivateMAPI = AttachToMAPIDll(WszMSMAPI32DLL);
		}

		// If MAPI isn't loaded in the process yet, then find the path to the DLL and
		// load it manually.
		if (NULL == hinstPrivateMAPI)
		{
			hinstPrivateMAPI = GetDefaultMapiHandle();
		}

		if (NULL != hinstPrivateMAPI)
		{
			SetMAPIHandle(hinstPrivateMAPI);
		}

		// Reason - if for any reason there is an instance already loaded, SetMAPIHandle()
		// will free the new one and reuse the old one
		// So we fetch the instance from the global again
		DebugPrint(DBGLoadMAPI, L"Exit GetPrivateMAPI: Returning GetMAPIHandle()\n");
		return GetMAPIHandle();
	}

	DebugPrint(DBGLoadMAPI, L"Exit GetPrivateMAPI, hinstPrivateMAPI = %p\n", hinstPrivateMAPI);
	return hinstPrivateMAPI;
}