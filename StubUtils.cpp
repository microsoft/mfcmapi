#include "stdafx.h"
#include <windows.h>
#include <strsafe.h>
#include <msi.h>
#include <winreg.h>
#include <stdlib.h>

// Included for MFCMAPI tracing
#include "MFCOutput.h"
#include "ImportProcs.h"

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

// Exists to allow some logging
_Check_return_ HMODULE MyLoadLibraryW(_In_z_ LPCWSTR lpszLibFileName)
{
	HMODULE hMod = NULL;
	HRESULT hRes = S_OK;
	DebugPrint(DBGLoadLibrary,_T("MyLoadLibrary - loading \"%ws\"\n"),lpszLibFileName);
	WC_D(hMod,LoadLibraryW(lpszLibFileName));
	if (hMod)
	{
		DebugPrint(DBGLoadLibrary,_T("MyLoadLibrary - \"%ws\" loaded at %p\n"),lpszLibFileName,hMod);
	}
	else
	{
		DebugPrint(DBGLoadLibrary,_T("MyLoadLibrary - \"%ws\" failed to load\n"),lpszLibFileName);
	}
	return hMod;
} // MyLoadLibrary

static volatile HMODULE g_hinstMAPI = NULL;

__inline HMODULE GetMAPIHandle()
{
	return g_hinstMAPI;
} // GetMAPIHandle

void SetMAPIHandle(HMODULE hinstMAPI)
{
	DebugPrint(DBGLoadMAPI,_T("Enter SetMAPIHandle: hinstMAPI = %p\n"),hinstMAPI);
	HMODULE	hinstNULL = NULL;
	HMODULE	hinstToFree = NULL;

	if (hinstMAPI == NULL)
	{
		hinstToFree = (HMODULE)InterlockedExchangePointer((PVOID*) &g_hinstMAPI, (PVOID) hinstNULL);
	}
	else
	{
		// Set the value only if the global is NULL
		HMODULE	hinstPrev;
		hinstPrev = (HMODULE)InterlockedCompareExchangePointer(reinterpret_cast<volatile PVOID*>(&g_hinstMAPI), hinstMAPI, hinstNULL);
		if (NULL != hinstPrev)
		{
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
	DebugPrint(DBGLoadMAPI,_T("Exit SetMAPIHandle\n"));
} // SetMAPIHandle

/*
 *  RegQueryWszExpand
 *		Wrapper for RegQueryValueExW which automatically expands REG_EXPAND_SZ values
 */
DWORD RegQueryWszExpand(HKEY hKey, LPCWSTR lpValueName, LPWSTR lpValue, DWORD cchValueLen)
{
	DebugPrint(DBGLoadMAPI,_T("Enter RegQueryWszExpand: hKey = %p, lpValueName = %ws, cchValueLen = 0x%08X\n"),
		hKey,lpValueName,cchValueLen);
	DWORD dwErr = ERROR_SUCCESS;
	DWORD dwType = 0;

	WCHAR rgchValue[MAX_PATH] = {0};
	DWORD dwSize = sizeof(rgchValue);

	dwErr = RegQueryValueExW(hKey, lpValueName, 0, &dwType, (LPBYTE) &rgchValue, &dwSize);

	if (dwErr == ERROR_SUCCESS)
	{
		DebugPrint(DBGLoadMAPI,_T("RegQueryWszExpand: rgchValue = %ws\n"),rgchValue);
		if (dwType == REG_EXPAND_SZ)
		{
			// Expand the strings
			DWORD cch = ExpandEnvironmentStringsW(rgchValue, lpValue, cchValueLen);
			if ((0 == cch) || (cch > cchValueLen))
			{
				dwErr = ERROR_INSUFFICIENT_BUFFER;
				goto Exit;
			}
			DebugPrint(DBGLoadMAPI,_T("RegQueryWszExpand: rgchValue(expanded) = %ws\n"), lpValue);
		}
		else if (dwType == REG_SZ)
		{
			wcscpy_s(lpValue, cchValueLen, rgchValue);
		}
	}
Exit:
	DebugPrint(DBGLoadMAPI,_T("Exit RegQueryWszExpand: dwErr = 0x%08X\n"),dwErr);
	return dwErr;
} // RegQueryWszExpand

/*
 *  GetComponentPath
 *		Wrapper around mapi32.dll->FGetComponentPath which maps an MSI component ID to 
 *		a DLL location from the default MAPI client registration values
 */
bool GetComponentPath(LPCSTR szComponent, LPSTR szQualifier, LPSTR szDllPath, DWORD cchBufferSize, bool fInstall)
{
	DebugPrint(DBGLoadMAPI,_T("Enter GetComponentPath: szComponent = %hs, szQualifier = %hs, cchBufferSize = 0x%08X, fInstall = 0x%08X\n"),
		szComponent,szQualifier,cchBufferSize,fInstall);
	HMODULE hMapiStub = NULL;
	bool fReturn = FALSE;

	typedef bool (STDAPICALLTYPE *FGetComponentPathType)(LPCSTR, LPSTR, LPSTR, DWORD, bool);

	hMapiStub = MyLoadLibraryW(WszMapi32);
	if (!hMapiStub)
		hMapiStub = MyLoadLibraryW(WszMapiStub);

	if (hMapiStub)
	{
		FGetComponentPathType pFGetCompPath = (FGetComponentPathType)GetProcAddress(hMapiStub, SzFGetComponentPath);

		fReturn = pFGetCompPath(szComponent, szQualifier, szDllPath, cchBufferSize, fInstall);
		DebugPrint(DBGLoadMAPI,_T("GetComponentPath: szDllPath = %hs\n"),szDllPath);

		FreeLibrary(hMapiStub);
	}

	DebugPrint(DBGLoadMAPI,_T("Exit GetComponentPath: fReturn = 0x%08X\n"),fReturn);
	return fReturn;
} // GetComponentPath

/*
 *  LoadMailClientFromMSIData
 *		Attempt to locate the MAPI provider DLL via HKLM\Software\Clients\Mail\(provider)\MSIComponentID
 */
HMODULE LoadMailClientFromMSIData(HKEY hkeyMapiClient)
{
	DebugPrint(DBGLoadMAPI,_T("Enter LoadMailClientFromMSIData: hkeyMapiClient = %p\n"),hkeyMapiClient);
	HMODULE		hinstMapi = NULL;
	CHAR		rgchMSIComponentID[MAX_PATH];
	CHAR		rgchMSIApplicationLCID[MAX_PATH];
	CHAR		rgchComponentPath[MAX_PATH];
	DWORD		dwType = 0;

	DWORD dwSizeComponentID = sizeof(rgchMSIComponentID);
	DWORD dwSizeLCID = sizeof(rgchMSIApplicationLCID);

	if (ERROR_SUCCESS == RegQueryValueExA(	hkeyMapiClient, SzValueNameMSI,0,
											&dwType, (LPBYTE) &rgchMSIComponentID, &dwSizeComponentID)
		&& ERROR_SUCCESS == RegQueryValueExA(	hkeyMapiClient, SzValueNameLCID,0,
												&dwType, (LPBYTE) &rgchMSIApplicationLCID, &dwSizeLCID))
	{
		if (GetComponentPath(rgchMSIComponentID, rgchMSIApplicationLCID, 
			rgchComponentPath, _countof(rgchComponentPath), FALSE))
		{
			DebugPrint(DBGLoadMAPI,_T("LoadMailClientFromMSIData: Loading %hs\n"),rgchComponentPath);
			hinstMapi = LoadLibraryA(rgchComponentPath);
		}
	}
	DebugPrint(DBGLoadMAPI,_T("Exit LoadMailClientFromMSIData: hinstMapi = %p\n"),hinstMapi);
	return hinstMapi;
} // LoadMailClientFromMSIData

/*
 *  LoadMailClientFromDllPath
 *		Attempt to locate the MAPI provider DLL via HKLM\Software\Clients\Mail\(provider)\DllPathEx
 */
HMODULE LoadMailClientFromDllPath(HKEY hkeyMapiClient)
{
	DebugPrint(DBGLoadMAPI,_T("Enter LoadMailClientFromDllPath: hkeyMapiClient = %p\n"),hkeyMapiClient);
	HMODULE hinstMapi = NULL;
	WCHAR rgchDllPath[MAX_PATH] = {0};

	DWORD dwSizeDllPath = _countof(rgchDllPath);

	if (ERROR_SUCCESS == RegQueryWszExpand(hkeyMapiClient, WszValueNameDllPathEx, rgchDllPath, dwSizeDllPath))
	{
		DebugPrint(DBGLoadMAPI,_T("LoadMailClientFromDllPath: DllPathEx = %ws\n"),rgchDllPath);
		hinstMapi = MyLoadLibraryW(rgchDllPath);
	}

	if (!hinstMapi)
	{
		dwSizeDllPath = _countof(rgchDllPath);
		if (ERROR_SUCCESS == RegQueryWszExpand(hkeyMapiClient, WszValueNameDllPath, rgchDllPath, dwSizeDllPath))
		{
			DebugPrint(DBGLoadMAPI,_T("LoadMailClientFromDllPath: DllPath = %ws\n"),rgchDllPath);
			hinstMapi = MyLoadLibraryW(rgchDllPath);
		}
	}
	DebugPrint(DBGLoadMAPI,_T("Exit LoadMailClientFromDllPath: hinstMapi = %p\n"),hinstMapi);
	return hinstMapi;
} // LoadMailClientFromDllPath

/*
 *  LoadRegisteredMapiClient
 *		Read the registry to discover the registered MAPI client and attempt to load its MAPI DLL.
 *		
 *		If wzOverrideProvider is specified, this function will load that MAPI Provider instead of the 
 *		currently registered provider
 */
HMODULE LoadRegisteredMapiClient(LPCWSTR pwzProviderOverride)
{
	DebugPrint(DBGLoadMAPI,_T("Enter LoadRegisteredMapiClient\n"));
	HMODULE		hinstMapi = NULL;
	DWORD		dwType;
	HKEY 		hkey = NULL, hkeyMapiClient = NULL;
	WCHAR		rgchMailClient[MAX_PATH];
	LPCWSTR		pwzProvider = pwzProviderOverride;

	// Open HKLM\Software\Clients\Mail
	if (ERROR_SUCCESS == RegOpenKeyExW( HKEY_LOCAL_MACHINE,
										WszKeyNameMailClient,
										0,
										KEY_READ,
										&hkey))
	{
		// If a specific provider wasn't specified, load the name of the default MAPI provider
		if (!pwzProvider)
		{
			// Get Outlook application path registry value
			DWORD dwSize = sizeof(rgchMailClient);
			if SUCCEEDED(RegQueryValueExW(	hkey,
				NULL,
				0,
				&dwType,
				(LPBYTE) &rgchMailClient,
				&dwSize))
			if (dwType != REG_SZ)
				goto Error;

			DebugPrint(DBGLoadMAPI,_T("LoadRegisteredMapiClient: HKLM\\%ws = %ws\n"),WszKeyNameMailClient,rgchMailClient);
			pwzProvider = rgchMailClient;
		}

		if (pwzProvider)
		{
			DebugPrint(DBGLoadMAPI,_T("LoadRegisteredMapiClient: pwzProvider = %ws%\n"), pwzProvider);
			if SUCCEEDED(RegOpenKeyExW( hkey,
				pwzProvider,
				0,
				KEY_READ,
				&hkeyMapiClient))
			{
				hinstMapi = LoadMailClientFromMSIData(hkeyMapiClient);

				if (!hinstMapi)
					hinstMapi = LoadMailClientFromDllPath(hkeyMapiClient);
			}
		}
	}

Error:
	DebugPrint(DBGLoadMAPI,_T("Exit LoadRegisteredMapiClient: hinstMapi = %p\n"),hinstMapi);
	return hinstMapi;
} // LoadRegisteredMapiClient

/*
 *  LoadMAPIFromSystemDir
 *		Fall back for loading System32\Mapi32.dll if all else fails
 */
HMODULE LoadMAPIFromSystemDir()
{
	DebugPrint(DBGLoadMAPI,_T("Enter LoadMAPIFromSystemDir\n"));
	WCHAR szSystemDir[MAX_PATH] = {0};

	if (GetSystemDirectoryW(szSystemDir, MAX_PATH))
	{
		WCHAR szDLLPath[MAX_PATH] = {0};
		swprintf_s(szDLLPath, _countof(szDLLPath), WszMAPISystemPath, szSystemDir, WszMapi32);
		DebugPrint(DBGLoadMAPI,_T("LoadMAPIFromSystemDir: loading %ws\n"),szDLLPath);
		return MyLoadLibraryW(szDLLPath);
	}

	DebugPrint(DBGLoadMAPI,_T("Exit LoadMAPIFromSystemDir: loading nothing\n"));
	return NULL;
} // LoadMAPIFromSystemDir

HMODULE GetDefaultMapiHandle()
{
	DebugPrint(DBGLoadMAPI,_T("Enter GetDefaultMapiHandle\n"));
	HMODULE	hinstMapi = NULL;

	// Try to respect the machine's default MAPI client settings.  If the active MAPI provider
	//  is Outlook, don't load and instead run the logic below
	if (!s_fForceSystemMAPI)
	{
		if (s_fForceOutlookMAPI)
			hinstMapi = LoadRegisteredMapiClient(WszOutlookMapiClientName);
		else
			hinstMapi = LoadRegisteredMapiClient(NULL);
	}

	// If MAPI still isn't loaded, load the stub from the system directory
	if (!hinstMapi && !s_fForceOutlookMAPI)
	{
		hinstMapi = LoadMAPIFromSystemDir();
	}

	DebugPrint(DBGLoadMAPI,_T("Exit LoadDefaultNonOutlookMapiClient: hinstMapi = %p\n"),hinstMapi);
	return hinstMapi;
} // GetDefaultMapiHandle

/*------------------------------------------------------------------------------
    Attach to wzMapiDll(olmapi32.dll/msmapi32.dll) if it is already loaded in the
	current process.
------------------------------------------------------------------------------*/
HMODULE AttachToMAPIDll(const WCHAR *wzMapiDll)
{
	DebugPrint(DBGLoadMAPI,_T("Enter AttachToMAPIDll: wzMapiDll = %ws\n"),wzMapiDll);
	HMODULE	hinstPrivateMAPI = NULL;
	MyGetModuleHandleExW(0UL, wzMapiDll, &hinstPrivateMAPI);
	DebugPrint(DBGLoadMAPI,_T("Exit AttachToMAPIDll: hinstPrivateMAPI = %p\n"),hinstPrivateMAPI);
	return hinstPrivateMAPI;
} // AttachToMAPIDll

void UnLoadPrivateMAPI()
{
	DebugPrint(DBGLoadMAPI,_T("Enter UnLoadPrivateMAPI\n"));
	HMODULE hinstPrivateMAPI = NULL;

	hinstPrivateMAPI = GetMAPIHandle();
	if (NULL != hinstPrivateMAPI)
	{
		SetMAPIHandle(NULL);
	}
	DebugPrint(DBGLoadMAPI,_T("Exit UnLoadPrivateMAPI\n"));
} // UnLoadPrivateMAPI

void ForceOutlookMAPI(bool fForce)
{
	DebugPrint(DBGLoadMAPI,_T("ForceOutlookMAPI: fForce = 0x%08X\n"),fForce);
	s_fForceOutlookMAPI = fForce;
} // ForceOutlookMAPI

void ForceSystemMAPI(bool fForce)
{
	DebugPrint(DBGLoadMAPI,_T("ForceSystemMAPI: fForce = 0x%08X\n"),fForce);
	s_fForceSystemMAPI = fForce;
} // ForceSystemMAPI

HMODULE GetPrivateMAPI()
{
	DebugPrint(DBGLoadMAPI,_T("Enter GetPrivateMAPI\n"));
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
		DebugPrint(DBGLoadMAPI,_T("Exit GetPrivateMAPI: Returning GetMAPIHandle()\n"));
		return GetMAPIHandle();
	}

	DebugPrint(DBGLoadMAPI,_T("Exit GetPrivateMAPI, hinstPrivateMAPI = %p\n"),hinstPrivateMAPI);
	return hinstPrivateMAPI;
} // GetPrivateMAPI