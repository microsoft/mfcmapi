#include "stdafx.h"
#include <windows.h>
#include <strsafe.h>
#include <msi.h>
#include <winreg.h>
#include <stdlib.h>

// Included for MFCMAPI tracing
#include "MFCOutput.h"
#include "ImportProcs.h"
#include "MAPIFunctions.h"

/*
 *  MAPI Stub Utilities
 *
 *	Public Functions:
 *
 *		GetPrivateMAPI()
 *			Obtain a handle to the MAPI DLL.  This function will load the MAPI DLL
 *			if it hasn't already been loaded
 *
 *		UnLoadPrivateMAPI()
 *			Forces the MAPI DLL to be unloaded.  This can cause problems if the code
 *			still has outstanding allocated MAPI memory, or unmatched calls to
 *			MAPIInitialize/MAPIUninitialize
 *
 *		ForceOutlookMAPI()
 *			Instructs the stub code to always try loading the Outlook version of MAPI
 *			on the system, instead of respecting the system MAPI registration
 *			(HKLM\Software\Clients\Mail). This call must be made prior to any MAPI
 *			function calls.
 */
HMODULE GetPrivateMAPI();
void UnLoadPrivateMAPI();
void ForceOutlookMAPI();

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
//  cause a re-fetch of all stored function pointers
volatile ULONG g_ulDllSequenceNum = 1;

// Whether or not we should ignore the system MAPI registration and always try to find
//  Outlook and its MAPI DLLs
static bool s_fForceOutlookMAPI = false;

// Whether or not we should ignore the registry and load MAPI from the system directory
static bool s_fForceSystemMAPI = false;

static volatile HMODULE g_hinstMAPI = NULL;
HMODULE g_hModPstPrx32 = NULL;

__inline HMODULE GetMAPIHandle()
{
	return g_hinstMAPI;
} // GetMAPIHandle

void SetMAPIHandle(HMODULE hinstMAPI)
{
	DebugPrint(DBGLoadMAPI, _T("Enter SetMAPIHandle: hinstMAPI = %p\n"), hinstMAPI);
	HMODULE	hinstNULL = NULL;
	HMODULE	hinstToFree = NULL;

	if (hinstMAPI == NULL)
	{
		// If we've preloaded pstprx32.dll, unload it before MAPI is unloaded to prevent dependency problems
		if (g_hModPstPrx32)
		{
			::FreeLibrary(g_hModPstPrx32);
			g_hModPstPrx32 = NULL;
		}

		hinstToFree = (HMODULE)InterlockedExchangePointer((PVOID*)&g_hinstMAPI, (PVOID)hinstNULL);
	}
	else
	{
		// Preload pstprx32 to prevent crash when using autodiscover to build a new profile
		if (!g_hModPstPrx32)
		{
			g_hModPstPrx32 = LoadFromOLMAPIDir(_T("pstprx32.dll")); // STRING_OK
		}

		// Set the value only if the global is NULL
		HMODULE	hinstPrev;
		// Code Analysis gives us a C28112 error when we use InterlockedCompareExchangePointer, so we instead exchange, check and exchange back
		//hinstPrev = (HMODULE)InterlockedCompareExchangePointer(reinterpret_cast<volatile PVOID*>(&g_hinstMAPI), hinstMAPI, hinstNULL);
		hinstPrev = (HMODULE)InterlockedExchangePointer((PVOID*)&g_hinstMAPI, (PVOID)hinstMAPI);
		if (NULL != hinstPrev)
		{
			(void)InterlockedExchangePointer((PVOID*)&g_hinstMAPI, (PVOID)hinstPrev);
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
	DebugPrint(DBGLoadMAPI, _T("Exit SetMAPIHandle\n"));
} // SetMAPIHandle

/*
 *  RegQueryWszExpand
 *		Wrapper for RegQueryValueExW which automatically expands REG_EXPAND_SZ values
 */
DWORD RegQueryWszExpand(HKEY hKey, LPCWSTR lpValueName, LPWSTR lpValue, DWORD cchValueLen)
{
	DebugPrint(DBGLoadMAPI, _T("Enter RegQueryWszExpand: hKey = %p, lpValueName = %ws, cchValueLen = 0x%08X\n"),
		hKey, lpValueName, cchValueLen);
	DWORD dwErr = ERROR_SUCCESS;
	DWORD dwType = 0;

	WCHAR rgchValue[MAX_PATH] = { 0 };
	DWORD dwSize = sizeof(rgchValue);

	dwErr = RegQueryValueExW(hKey, lpValueName, 0, &dwType, (LPBYTE)&rgchValue, &dwSize);

	if (dwErr == ERROR_SUCCESS)
	{
		DebugPrint(DBGLoadMAPI, _T("RegQueryWszExpand: rgchValue = %ws\n"), rgchValue);
		if (dwType == REG_EXPAND_SZ)
		{
			// Expand the strings
			DWORD cch = ExpandEnvironmentStringsW(rgchValue, lpValue, cchValueLen);
			if ((0 == cch) || (cch > cchValueLen))
			{
				dwErr = ERROR_INSUFFICIENT_BUFFER;
				goto Exit;
			}
			DebugPrint(DBGLoadMAPI, _T("RegQueryWszExpand: rgchValue(expanded) = %ws\n"), lpValue);
		}
		else if (dwType == REG_SZ)
		{
			wcscpy_s(lpValue, cchValueLen, rgchValue);
		}
	}
Exit:
	DebugPrint(DBGLoadMAPI, _T("Exit RegQueryWszExpand: dwErr = 0x%08X\n"), dwErr);
	return dwErr;
} // RegQueryWszExpand

/*
 *  GetComponentPath
 *		Wrapper around mapi32.dll->FGetComponentPath which maps an MSI component ID to
 *		a DLL location from the default MAPI client registration values
 */
bool GetComponentPath(LPCSTR szComponent, LPSTR szQualifier, LPSTR szDllPath, DWORD cchBufferSize, bool fInstall)
{
	DebugPrint(DBGLoadMAPI, _T("Enter GetComponentPath: szComponent = %hs, szQualifier = %hs, cchBufferSize = 0x%08X, fInstall = 0x%08X\n"),
		szComponent, szQualifier, cchBufferSize, fInstall);
	HMODULE hMapiStub = NULL;
	bool fReturn = FALSE;

	typedef bool (STDAPICALLTYPE *FGetComponentPathType)(LPCSTR, LPSTR, LPSTR, DWORD, bool);

	hMapiStub = MyLoadLibraryW(WszMapi32);
	if (!hMapiStub)
		hMapiStub = MyLoadLibraryW(WszMapiStub);

	if (hMapiStub)
	{
		FGetComponentPathType pFGetCompPath = (FGetComponentPathType)GetProcAddress(hMapiStub, SzFGetComponentPath);

		if (pFGetCompPath)
		{
			fReturn = pFGetCompPath(szComponent, szQualifier, szDllPath, cchBufferSize, fInstall);
			DebugPrint(DBGLoadMAPI, _T("GetComponentPath: szDllPath = %hs\n"), szDllPath);
		}

		FreeLibrary(hMapiStub);
	}

	DebugPrint(DBGLoadMAPI, _T("Exit GetComponentPath: fReturn = 0x%08X\n"), fReturn);
	return fReturn;
} // GetComponentPath

enum mapiSource
{
	msInstalledOutlook,
	msRegisteredMSI,
	msRegisteredDLLEx,
	msRegisteredDLL,
	msSystem,
	msEnd,
};

MAPIPathIterator::MAPIPathIterator(bool bBypassRestrictions)
{
	m_bBypassRestrictions = bBypassRestrictions;
	m_szRegisteredClient = NULL;
	if (bBypassRestrictions)
	{
		CurrentSource = msInstalledOutlook;
	}
	else
	{
		if (!s_fForceSystemMAPI)
		{
			CurrentSource = msRegisteredMSI;
			if (s_fForceOutlookMAPI)
				m_szRegisteredClient = WszOutlookMapiClientName;
		}
		else
			CurrentSource = msSystem;
	}
	m_hMailKey = NULL;
	m_hkeyMapiClient = NULL;
	m_rgchMailClient = NULL;

	m_iCurrentOutlook = oqcOfficeBegin;
}

MAPIPathIterator::~MAPIPathIterator()
{
	delete[] m_rgchMailClient;
	if (m_hMailKey) RegCloseKey(m_hMailKey);
	if (m_hkeyMapiClient) RegCloseKey(m_hkeyMapiClient);
}

LPWSTR MAPIPathIterator::GetNextMAPIPath()
{
	// Mini state machine here will get the path from the current source then set the next source to search
	// Either returns the next available MAPI path or NULL if none remain
	LPWSTR szPath = NULL;
	while (msEnd != CurrentSource && !szPath)
	{
		switch (CurrentSource)
		{
		case msInstalledOutlook:
			szPath = GetNextInstalledOutlookMAPI();

			// We'll keep trying GetNextInstalledOutlookMAPI as long as it returns results
			if (!szPath)
			{
				CurrentSource = msRegisteredMSI;
			}
			break;
		case msRegisteredMSI:
			szPath = GetRegisteredMapiClient(WszOutlookMapiClientName, false, false);
			CurrentSource = msRegisteredDLLEx;
			break;
		case msRegisteredDLLEx:
			szPath = GetRegisteredMapiClient(WszOutlookMapiClientName, true, true);
			CurrentSource = msRegisteredDLL;
			break;
		case msRegisteredDLL:
			szPath = GetRegisteredMapiClient(WszOutlookMapiClientName, true, false);
			if (s_fForceOutlookMAPI && !m_bBypassRestrictions)
			{
				CurrentSource = msEnd;
			}
			else
			{
				CurrentSource = msSystem;
			}
			break;
		case msSystem:
			szPath = GetMAPISystemDir();
			CurrentSource = msEnd;
			break;
		case msEnd:
		default:
			break;
		}
	}
	return szPath;
}

/*
 *  GetMailClientFromMSIData
 *		Attempt to locate the MAPI provider DLL via HKLM\Software\Clients\Mail\(provider)\MSIComponentID
 */
LPWSTR MAPIPathIterator::GetMailClientFromMSIData(HKEY hkeyMapiClient)
{
	DebugPrint(DBGLoadMAPI, _T("Enter GetMailClientFromMSIData: hkeyMapiClient = %p\n"), hkeyMapiClient);
	CHAR rgchMSIComponentID[MAX_PATH] = { 0 };
	CHAR rgchMSIApplicationLCID[MAX_PATH] = { 0 };
	CHAR rgchComponentPath[MAX_PATH] = { 0 };
	DWORD dwType = 0;
	LPWSTR szPath = NULL;
	HRESULT hRes = S_OK;

	DWORD dwSizeComponentID = sizeof(rgchMSIComponentID);
	DWORD dwSizeLCID = sizeof(rgchMSIApplicationLCID);

	if (ERROR_SUCCESS == RegQueryValueExA(hkeyMapiClient, SzValueNameMSI, 0, &dwType, (LPBYTE)&rgchMSIComponentID, &dwSizeComponentID) &&
		ERROR_SUCCESS == RegQueryValueExA(hkeyMapiClient, SzValueNameLCID, 0, &dwType, (LPBYTE)&rgchMSIApplicationLCID, &dwSizeLCID))
	{
		if (GetComponentPath(rgchMSIComponentID, rgchMSIApplicationLCID, rgchComponentPath, _countof(rgchComponentPath), FALSE))
		{
			WC_H(AnsiToUnicode(rgchComponentPath, &szPath));
		}
	}
	DebugPrint(DBGLoadMAPI, _T("Exit GetMailClientFromMSIData: szPath = %ws\n"), szPath);
	return szPath;
} // MAPIPathIterator::GetMailClientFromMSIData

/*
 *  GetMailClientFromDllPath
 *		Attempt to locate the MAPI provider DLL via HKLM\Software\Clients\Mail\(provider)\DllPathEx
 */
LPWSTR MAPIPathIterator::GetMailClientFromDllPath(HKEY hkeyMapiClient, bool bEx)
{
	DebugPrint(DBGLoadMAPI, _T("Enter GetMailClientFromDllPath: hkeyMapiClient = %p, bEx = %d\n"), hkeyMapiClient, bEx);
	HRESULT hRes = S_OK;
	LPWSTR szPath = NULL;

	szPath = new WCHAR[MAX_PATH];

	if (szPath)
	{
		if (bEx)
		{
			WC_W32(RegQueryWszExpand(hkeyMapiClient, WszValueNameDllPathEx, szPath, MAX_PATH));
		}
		else
		{
			WC_W32(RegQueryWszExpand(hkeyMapiClient, WszValueNameDllPath, szPath, MAX_PATH));
		}
		if (FAILED(hRes))
		{
			delete[] szPath;
			szPath = NULL;
		}
	}

	DebugPrint(DBGLoadMAPI, _T("Exit GetMailClientFromDllPath: szPath = %ws\n"), szPath);
	return szPath;
} // MAPIPathIterator::GetMailClientFromDllPath

/*
 *  GetRegisteredMapiClient
 *		Read the registry to discover the registered MAPI client and attempt to load its MAPI DLL.
 *
 *		If wzOverrideProvider is specified, this function will load that MAPI Provider instead of the
 *		currently registered provider
 */
LPWSTR MAPIPathIterator::GetRegisteredMapiClient(LPCWSTR pwzProviderOverride, bool bDLL, bool bEx)
{
	DebugPrint(DBGLoadMAPI, _T("Enter GetRegisteredMapiClient\n"));
	HRESULT hRes = S_OK;
	LPWSTR szPath = NULL;
	LPCWSTR pwzProvider = pwzProviderOverride;

	if (!m_hMailKey)
	{
		// Open HKLM\Software\Clients\Mail
		WC_W32(RegOpenKeyExW(HKEY_LOCAL_MACHINE,
			WszKeyNameMailClient,
			0,
			KEY_READ,
			&m_hMailKey));
		if (FAILED(hRes))
		{
			m_hMailKey = NULL;
		}
	}

	// If a specific provider wasn't specified, load the name of the default MAPI provider
	if (m_hMailKey && !pwzProvider && !m_rgchMailClient)
	{
		m_rgchMailClient = new WCHAR[MAX_PATH];
		if (m_rgchMailClient)
		{
			// Get Outlook application path registry value
			DWORD dwSize = MAX_PATH;
			DWORD dwType = 0;
			WC_W32(RegQueryValueExW(
				m_hMailKey,
				NULL,
				0,
				&dwType,
				(LPBYTE)m_rgchMailClient,
				&dwSize));
			if (SUCCEEDED(hRes))
			{
				DebugPrint(DBGLoadMAPI, _T("GetRegisteredMapiClient: HKLM\\%ws = %ws\n"), WszKeyNameMailClient, m_rgchMailClient);
			}
			else
			{
				delete[] m_rgchMailClient;
				m_rgchMailClient = NULL;
			}
		}
	}

	if (!pwzProvider) pwzProvider = m_rgchMailClient;

	if (m_hMailKey && pwzProvider && !m_hkeyMapiClient)
	{
		DebugPrint(DBGLoadMAPI, _T("GetRegisteredMapiClient: pwzProvider = %ws%\n"), pwzProvider);
		WC_W32(RegOpenKeyExW(
			m_hMailKey,
			pwzProvider,
			0,
			KEY_READ,
			&m_hkeyMapiClient));
		if (FAILED(hRes))
		{
			m_hkeyMapiClient = NULL;
		}
	}

	if (m_hkeyMapiClient)
	{
		if (bDLL)
		{
			szPath = GetMailClientFromDllPath(m_hkeyMapiClient, bEx);
		}
		else
		{
			szPath = GetMailClientFromMSIData(m_hkeyMapiClient);
		}

	}

	DebugPrint(DBGLoadMAPI, _T("Exit GetRegisteredMapiClient: szPath = %ws\n"), szPath);
	return szPath;
} // MAPIPathIterator::GetRegisteredMapiClient

/*
 *  GetMAPISystemDir
 *		Fall back for loading System32\Mapi32.dll if all else fails
 */
LPWSTR MAPIPathIterator::GetMAPISystemDir()
{
	DebugPrint(DBGLoadMAPI, _T("Enter GetMAPISystemDir\n"));
	WCHAR szSystemDir[MAX_PATH] = { 0 };

	if (GetSystemDirectoryW(szSystemDir, MAX_PATH))
	{
		LPWSTR szDLLPath = new WCHAR[MAX_PATH];
		if (szDLLPath)
		{
			swprintf_s(szDLLPath, MAX_PATH, WszMAPISystemPath, szSystemDir, WszMapi32);
			DebugPrint(DBGLoadMAPI, _T("GetMAPISystemDir: found %ws\n"), szDLLPath);
			return szDLLPath;
		}
	}

	DebugPrint(DBGLoadMAPI, _T("Exit GetMAPISystemDir: found nothing\n"));
	return NULL;
} // MAPIPathIterator::GetMAPISystemDir

LPWSTR MAPIPathIterator::GetInstalledOutlookMAPI(int iOutlook)
{
	DebugPrint(DBGLoadMAPI, _T("Enter GetInstalledOutlookMAPI(%d)\n"), iOutlook);
	HRESULT hRes = S_OK;

	if (!pfnMsiProvideQualifiedComponent || !pfnMsiGetFileVersion) return NULL;

	UINT ret = 0;

	LPTSTR lpszTempPath = GetOutlookPath(g_pszOutlookQualifiedComponents[iOutlook], NULL);

	if (lpszTempPath)
	{
		TCHAR szDrive[_MAX_DRIVE] = { 0 };
		TCHAR szOutlookPath[MAX_PATH] = { 0 };
		WC_D(ret, _tsplitpath_s(lpszTempPath, szDrive, _MAX_DRIVE, szOutlookPath, MAX_PATH, NULL, NULL, NULL, NULL));

		if (SUCCEEDED(hRes))
		{
			LPWSTR szPath = new WCHAR[MAX_PATH];
			if (szPath)
			{
#ifdef UNICODE
				swprintf_s(szPath, MAX_PATH, WszMAPISystemDrivePath, szDrive, szOutlookPath, WszOlMAPI32DLL);
#else
				swprintf_s(szPath, MAX_PATH, szMAPISystemDrivePath, szDrive, szOutlookPath, WszOlMAPI32DLL);
#endif
			}
			delete[] lpszTempPath;
			DebugPrint(DBGLoadMAPI, _T("GetInstalledOutlookMAPI: found %ws\n"), szPath);
			return szPath;
		}
		delete[] lpszTempPath;
	}

	DebugPrint(DBGLoadMAPI, _T("Exit GetInstalledOutlookMAPI: found nothing\n"));
	return NULL;
} // MAPIPathIterator::GetInstalledOutlookMAPI

LPWSTR MAPIPathIterator::GetNextInstalledOutlookMAPI()
{
	DebugPrint(DBGLoadMAPI, _T("Enter GetNextInstalledOutlookMAPI\n"));

	if (!pfnMsiProvideQualifiedComponent || !pfnMsiGetFileVersion) return NULL;

	for (; m_iCurrentOutlook < oqcOfficeEnd; m_iCurrentOutlook++)
	{
		LPWSTR szPath = GetInstalledOutlookMAPI(m_iCurrentOutlook);
		if (szPath)
		{
			m_iCurrentOutlook++; // Make sure we don't repeat this Outlook
			return szPath;
		}
	}

	DebugPrint(DBGLoadMAPI, _T("Exit GetNextInstalledOutlookMAPI: found nothing\n"));
	return NULL;
} // MAPIPathIterator::GetNextInstalledOutlookMAPI

TCHAR g_pszOutlookQualifiedComponents[][MAX_PATH] = {
	_T("{E83B4360-C208-4325-9504-0D23003A74A5}"), // O15_CATEGORY_GUID_CORE_OFFICE (retail) // STRING_OK
	_T("{1E77DE88-BCAB-4C37-B9E5-073AF52DFD7A}"), // O14_CATEGORY_GUID_CORE_OFFICE (retail) // STRING_OK
	_T("{24AAE126-0911-478F-A019-07B875EB9996}"), // O12_CATEGORY_GUID_CORE_OFFICE (retail) // STRING_OK
	_T("{BC174BAD-2F53-4855-A1D5-0D575C19B1EA}"), // O11_CATEGORY_GUID_CORE_OFFICE (retail) // STRING_OK
	_T("{BC174BAD-2F53-4855-A1D5-1D575C19B1EA}"), // O11_CATEGORY_GUID_CORE_OFFICE (debug)  // STRING_OK
};

// Looks up Outlook's path given its qualified component guid
LPTSTR GetOutlookPath(_In_z_ LPCTSTR szCategory, _Out_opt_ bool* lpb64)
{
	DebugPrint(DBGLoadMAPI, _T("Enter GetOutlookPath: szCategory = %s\n"), szCategory);
	HRESULT hRes = S_OK;
	DWORD dwValueBuf = 0;
	UINT ret = 0;

	if (lpb64) *lpb64 = false;

	WC_D(ret, pfnMsiProvideQualifiedComponent(
		szCategory,
		_T("outlook.x64.exe"), // STRING_OK
		(DWORD)INSTALLMODE_DEFAULT,
		NULL,
		&dwValueBuf));
	if (ERROR_SUCCESS == ret)
	{
		if (lpb64) *lpb64 = true;
	}
	else
	{
		ret = ERROR_SUCCESS;
		WC_D(ret, pfnMsiProvideQualifiedComponent(
			szCategory,
			_T("outlook.exe"), // STRING_OK
			(DWORD)INSTALLMODE_DEFAULT,
			NULL,
			&dwValueBuf));
	}

	if (ERROR_SUCCESS == ret)
	{
		LPTSTR lpszTempPath = NULL;
		dwValueBuf += 1;
		lpszTempPath = new TCHAR[dwValueBuf];

		if (lpszTempPath != NULL)
		{
			WC_D(ret, pfnMsiProvideQualifiedComponent(
				szCategory,
				_T("outlook.x64.exe"), // STRING_OK
				(DWORD)INSTALLMODE_DEFAULT,
				lpszTempPath,
				&dwValueBuf));
			if (ERROR_SUCCESS != ret)
			{
				ret = ERROR_SUCCESS;
				WC_D(ret, pfnMsiProvideQualifiedComponent(
					szCategory,
					_T("outlook.exe"), // STRING_OK
					(DWORD)INSTALLMODE_DEFAULT,
					lpszTempPath,
					&dwValueBuf));
			}

			if (ERROR_SUCCESS == ret)
			{
				DebugPrint(DBGLoadMAPI, _T("Exit GetOutlookPath: Path = %s\n"), lpszTempPath);
				return lpszTempPath;
			}

			delete[] lpszTempPath;
		}
	}

	DebugPrint(DBGLoadMAPI, _T("Exit GetOutlookPath: nothing found\n"));
	return NULL;
} // LPWSTR GetOutlookPath

HMODULE GetDefaultMapiHandle()
{
	DebugPrint(DBGLoadMAPI, _T("Enter GetDefaultMapiHandle\n"));
	HMODULE hinstMapi = NULL;

	LPWSTR szPath = NULL;
	MAPIPathIterator* mpi = new MAPIPathIterator(false);

	if (mpi)
	{
		while (!hinstMapi)
		{
			szPath = mpi->GetNextMAPIPath();
			if (!szPath) break;

			DebugPrint(DBGLoadMAPI, _T("Trying %ws\n"), szPath);
			hinstMapi = MyLoadLibraryW(szPath);
			delete[] szPath;
		}
	}

	delete mpi;
	DebugPrint(DBGLoadMAPI, _T("Exit GetDefaultMapiHandle: hinstMapi = %p\n"), hinstMapi);
	return hinstMapi;
} // GetDefaultMapiHandle

/*------------------------------------------------------------------------------
	Attach to wzMapiDll(olmapi32.dll/msmapi32.dll) if it is already loaded in the
	current process.
	------------------------------------------------------------------------------*/
HMODULE AttachToMAPIDll(const WCHAR *wzMapiDll)
{
	DebugPrint(DBGLoadMAPI, _T("Enter AttachToMAPIDll: wzMapiDll = %ws\n"), wzMapiDll);
	HMODULE	hinstPrivateMAPI = NULL;
	MyGetModuleHandleExW(0UL, wzMapiDll, &hinstPrivateMAPI);
	DebugPrint(DBGLoadMAPI, _T("Exit AttachToMAPIDll: hinstPrivateMAPI = %p\n"), hinstPrivateMAPI);
	return hinstPrivateMAPI;
} // AttachToMAPIDll

void UnLoadPrivateMAPI()
{
	DebugPrint(DBGLoadMAPI, _T("Enter UnLoadPrivateMAPI\n"));
	HMODULE hinstPrivateMAPI = NULL;

	hinstPrivateMAPI = GetMAPIHandle();
	if (NULL != hinstPrivateMAPI)
	{
		SetMAPIHandle(NULL);
	}
	DebugPrint(DBGLoadMAPI, _T("Exit UnLoadPrivateMAPI\n"));
} // UnLoadPrivateMAPI

void ForceOutlookMAPI(bool fForce)
{
	DebugPrint(DBGLoadMAPI, _T("ForceOutlookMAPI: fForce = 0x%08X\n"), fForce);
	s_fForceOutlookMAPI = fForce;
} // ForceOutlookMAPI

void ForceSystemMAPI(bool fForce)
{
	DebugPrint(DBGLoadMAPI, _T("ForceSystemMAPI: fForce = 0x%08X\n"), fForce);
	s_fForceSystemMAPI = fForce;
} // ForceSystemMAPI

HMODULE GetPrivateMAPI()
{
	DebugPrint(DBGLoadMAPI, _T("Enter GetPrivateMAPI\n"));
	HMODULE hinstPrivateMAPI = GetMAPIHandle();

	if (NULL == hinstPrivateMAPI)
	{
		// First, try to attach to olmapi32.dll if it's loaded in the process
		hinstPrivateMAPI = AttachToMAPIDll(WszOlMAPI32DLL);

		// If that fails try msmapi32.dll, for Outlook 11 and below
		//  Only try this in the static lib, otherwise msmapi32.dll will attach to itself.
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
		DebugPrint(DBGLoadMAPI, _T("Exit GetPrivateMAPI: Returning GetMAPIHandle()\n"));
		return GetMAPIHandle();
	}

	DebugPrint(DBGLoadMAPI, _T("Exit GetPrivateMAPI, hinstPrivateMAPI = %p\n"), hinstPrivateMAPI);
	return hinstPrivateMAPI;
} // GetPrivateMAPI