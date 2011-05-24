#include "stdafx.h"
#include <windows.h>
#include <strsafe.h>
#include <msi.h>
#include <winreg.h>
#include <stdlib.h>


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
const WCHAR WszKeyNameMSI[] = L"Software\\Clients\\Mail\\Microsoft Outlook";
const WCHAR WszValueNameLCID[] = L"MSIApplicationLCID";
const WCHAR WszValueNameOfficeLCID[] = L"MSIOfficeLCID";
const WCHAR WszKeyNameSoftware[] = L"Software";
const WCHAR WszKeyNamePolicy[] = L"Software\\Policy";
const WCHAR WszValueNameDllPathEx[] = L"DllPathEx";
const WCHAR WszValueNameDllPath[] = L"DllPath";

const CHAR SzValueNameMSI[] = "MSIComponentID";
const CHAR SzValueNameLCID[] = "MSIApplicationLCID";

const WCHAR WszOutlookMapiClientName[] = L"Microsoft Outlook";

const WCHAR WszValueNameMSI[] = L"MSIComponentID";
const WCHAR WszCurVersion[] = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion";
const WCHAR WszCommonFiles[] = L"CommonFilesDir";
const WCHAR WszMSMAPI32DLL[] = L"MSMAPI32.DLL";
const WCHAR WszMAPIPathFormat[] = L"System\\MSMAPI\\%d\\%s";
const WCHAR WszFullQualifier[] = L"%lu\\NT";
const WCHAR WszShortQualifier[] = L"%lu";
const WCHAR WszMAPISystemPath[] = L"%s\\%s";

static const WCHAR WszOutlookAppPath[] = L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\OUTLOOK.EXE";

static const WCHAR WszOutlookExe[] = L"outlook.exe";

static const WCHAR WszPrivateMAPI[] = L"olmapi32.dll";
static const WCHAR WszPrivateMAPI_11[] = L"msmapi32.dll";
static const WCHAR WszMapi32[] = L"mapi32.dll";
static const WCHAR WszMapiStub[] = L"mapistub.dll";

static const CHAR SzFGetComponentPath[] = "FGetComponentPath";

// Sequence number which is incremented every time we set our MAPI handle which will
//  cause a re-fetch of all stored function pointers
volatile ULONG g_ulDllSequenceNum = 1;

// Whether or not we should ignore the system MAPI registration and always try to find
//  Outlook and its MAPI DLLs
static bool s_fForceOutlookMAPI = false;


// Constants for use with GetInstalledLCID API
#define	lastUILanguage		0
#define	userLanguage		1
#define	systemLanguage		2
#define	defaultLanguage		3
#define	MAX_LANGLOADORDER	4

// MSI aware FindFile API flags
#define FINDFILEMSI_FIND_QUALIFIED      0x00000001
// Skip the INSTALLMODE_DEFAULT => INSTALLMODE_EXISTING + INSTALLMODE_DEFAULT
#define FINDFILEMSI_SKIP_OPTIMIZATION_FOR_EXISTING 0x80000000

// Basic MSI capable LoadLibrary implementation
#define LOADLIBMSI_LOAD_QUALIFIED       0x80000000
#define LOADLIBMSI_LOAD_ALWAYS          0x40000000
#define LOADLIBMSI_IGNORE_PRIOR         0x20000000
#define LOADLIBMSI_FLAG_MASK            0xFF000000


/*
 *  GetOutlookLCID
 *
 *  Purpose:
 *  Retrieves the preferred LCID for UI from the appropriate locations in
 *  the registry.
 *
 */
LCID GetOutlookLCID()
{
	HRESULT hr = S_OK;
	DWORD dwSize;
	DWORD dwSize2;
	DWORD dwType;
	HKEY hkMAPI1 = NULL;
	HKEY hkMAPI2 = NULL;
	HKEY hkSoftware = NULL;
	HKEY hkPolicy = NULL;
	HKEY hkSoftwareLcid = NULL;
	HKEY hkPolicyLcid = NULL;
	WCHAR rgwchMSILCID[(MAX_PATH * 2) + 10];
	LCID  lcid = 0;

	if (RegOpenKeyExW(HKEY_CURRENT_USER, WszKeyNamePolicy, 0, KEY_READ,
					  &hkPolicy) != ERROR_SUCCESS)
	{
		hkPolicy = NULL;
	}
	if (RegOpenKeyExW(HKEY_CURRENT_USER, WszKeyNameSoftware, 0, KEY_READ,
					  &hkSoftware) != ERROR_SUCCESS)
	{
		hkSoftware= NULL;
	}
	if ((hkSoftware == NULL) && (hkPolicy == NULL))
	{
		lcid =0;
		goto Done;
	}

	// Check the Current user tree
	if (RegOpenKeyExW(HKEY_CURRENT_USER, WszKeyNameMSI, 0, KEY_READ,
					  &hkMAPI1) == ERROR_SUCCESS)
	{
		dwSize = sizeof(rgwchMSILCID);
		hr = RegQueryValueExW(hkMAPI1, WszValueNameLCID, 0, &dwType,
							  (LPBYTE)rgwchMSILCID, &dwSize);
		if ((ERROR_SUCCESS == hr) && (REG_MULTI_SZ == dwType))
		{
			// Read the key
			LPWSTR szKey = NULL;

			// Check policy hive first
			if (RegOpenKeyExW(hkPolicy, rgwchMSILCID, 0, KEY_READ,
							  &hkPolicyLcid) != ERROR_SUCCESS)
			{
				hkPolicyLcid = NULL;
			}

			if (RegOpenKeyExW(hkSoftware, rgwchMSILCID, 0, KEY_READ,
							  &hkSoftwareLcid) != ERROR_SUCCESS)
			{
				hkSoftwareLcid = NULL;
			}
			dwSize = sizeof(lcid);
			dwSize2 = sizeof(lcid);
			szKey = &rgwchMSILCID[wcslen(rgwchMSILCID)+1];
			if ( (hkPolicyLcid &&
					(RegQueryValueExW(hkPolicyLcid, szKey, 0, &dwType,  (LPBYTE)&lcid, &dwSize) == ERROR_SUCCESS) &&
					(REG_DWORD == dwType) ) ||
					(hkSoftwareLcid &&
					 (RegQueryValueExW(hkSoftwareLcid, szKey, 0, &dwType,
									   (LPBYTE)&lcid, &dwSize2) == ERROR_SUCCESS) &&
					 (REG_DWORD == dwType) ))
			{
				goto Done;
			}
			if (NULL != hkPolicyLcid)
			{
				RegCloseKey(hkPolicyLcid);
				hkPolicyLcid = NULL;
			}
			if (NULL != hkSoftwareLcid)
			{
				RegCloseKey(hkSoftwareLcid);
				hkSoftwareLcid = NULL;
			}
		}

		// Now try the office key
		dwSize = sizeof(rgwchMSILCID);
		hr = RegQueryValueExW(hkMAPI1, WszValueNameOfficeLCID, 0, &dwType,
							  (LPBYTE)rgwchMSILCID, &dwSize);
		if ((ERROR_SUCCESS == hr) && (REG_MULTI_SZ == dwType))
		{
			// Read the key
			LPWSTR szKey = NULL;

			// Check policy hive first
			if (RegOpenKeyExW(hkPolicy, rgwchMSILCID, 0, KEY_READ,
							  &hkPolicyLcid) != ERROR_SUCCESS)
			{
				hkPolicyLcid = NULL;
			}

			if (RegOpenKeyExW(hkSoftware, rgwchMSILCID, 0, KEY_READ,
							  &hkSoftwareLcid) != ERROR_SUCCESS)
			{
				hkSoftwareLcid = NULL;
			}
			dwSize = sizeof(lcid);
			dwSize2 = sizeof(lcid);
			szKey = &rgwchMSILCID[wcslen(rgwchMSILCID)+1];
			if ( (hkPolicyLcid &&
					(RegQueryValueExW(hkPolicyLcid, szKey, 0, &dwType,  (LPBYTE)&lcid, &dwSize) == ERROR_SUCCESS) &&
					(REG_DWORD == dwType) ) ||
					(hkSoftwareLcid &&
					 (RegQueryValueExW(hkSoftwareLcid, szKey, 0, &dwType,
									   (LPBYTE)&lcid, &dwSize2) == ERROR_SUCCESS) &&
					 (REG_DWORD == dwType) ))
			{
				goto Done;
			}
			if (NULL != hkPolicyLcid)
			{
				RegCloseKey(hkPolicyLcid);
				hkPolicyLcid = NULL;
			}
			if (NULL != hkSoftwareLcid)
			{
				RegCloseKey(hkSoftwareLcid);
				hkSoftwareLcid = NULL;
			}
		}
		lcid = 0;
	}

	// Check the local machine tree
	if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, WszKeyNameMSI, 0, KEY_READ,
					  &hkMAPI1) == ERROR_SUCCESS)
	{
		dwSize = sizeof(rgwchMSILCID);
		hr = RegQueryValueExW(hkMAPI1, WszValueNameLCID, 0, &dwType,
							  (LPBYTE)rgwchMSILCID, &dwSize);
		if ((ERROR_SUCCESS == hr) && (REG_MULTI_SZ == dwType))
		{
			// Read the key
			LPWSTR szKey = NULL;

			// Check policy hive first
			if (RegOpenKeyExW(hkPolicy, rgwchMSILCID, 0, KEY_READ,
							  &hkPolicyLcid) != ERROR_SUCCESS)
			{
				hkPolicyLcid = NULL;
			}

			if (RegOpenKeyExW(hkSoftware, rgwchMSILCID, 0, KEY_READ,
							  &hkSoftwareLcid) != ERROR_SUCCESS)
			{
				hkSoftwareLcid = NULL;
			}
			dwSize = sizeof(lcid);
			dwSize2 = sizeof(lcid);
			szKey = &rgwchMSILCID[wcslen(rgwchMSILCID)+1];
			if ( (hkSoftwareLcid &&
					(RegQueryValueExW(hkSoftwareLcid, szKey, 0, &dwType,  (LPBYTE)&lcid, &dwSize) == ERROR_SUCCESS) &&
					(REG_DWORD == dwType) ) ||
					(hkPolicyLcid &&
					 (RegQueryValueExW(hkPolicyLcid, szKey, 0, &dwType,
									   (LPBYTE)&lcid, &dwSize2) == ERROR_SUCCESS) &&
					 (REG_DWORD == dwType) ))
			{
				goto Done;
			}
			if (NULL != hkPolicyLcid)
			{
				RegCloseKey(hkPolicyLcid);
				hkPolicyLcid = NULL;
			}
			if (NULL != hkSoftwareLcid)
			{
				RegCloseKey(hkSoftwareLcid);
				hkSoftwareLcid = NULL;
			}
		}
		lcid = 0;
	}

	if (RegOpenKeyExW(HKEY_LOCAL_MACHINE, WszKeyNameMSI, 0, KEY_READ,
					  &hkMAPI1) == ERROR_SUCCESS)
	{
		dwSize = sizeof(rgwchMSILCID);
		hr = RegQueryValueExW(hkMAPI1, WszValueNameOfficeLCID, 0, &dwType,
							  (LPBYTE)rgwchMSILCID, &dwSize);
		if ((ERROR_SUCCESS == hr) && (REG_MULTI_SZ == dwType))
		{
			// Read the key
			LPWSTR szKey = NULL;

			// Check policy hive first
			if (RegOpenKeyExW(hkPolicy, rgwchMSILCID, 0, KEY_READ,
							  &hkPolicyLcid) != ERROR_SUCCESS)
			{
				hkPolicyLcid = NULL;
			}

			if (RegOpenKeyExW(hkSoftware, rgwchMSILCID, 0, KEY_READ,
							  &hkSoftwareLcid) != ERROR_SUCCESS)
			{
				hkSoftwareLcid = NULL;
			}
			dwSize = sizeof(lcid);
			dwSize2 = sizeof(lcid);
			szKey = &rgwchMSILCID[wcslen(rgwchMSILCID)+1];
			if ( (hkSoftwareLcid &&
					(RegQueryValueExW(hkSoftwareLcid, szKey, 0, &dwType,  (LPBYTE)&lcid, &dwSize) == ERROR_SUCCESS) &&
					(REG_DWORD == dwType) ) ||
					(hkPolicyLcid &&
					 (RegQueryValueExW(hkPolicyLcid, szKey, 0, &dwType,
									   (LPBYTE)&lcid, &dwSize2) == ERROR_SUCCESS) &&
					 (REG_DWORD == dwType) ))
			{
				goto Done;
			}
			if (NULL != hkPolicyLcid)
			{
				RegCloseKey(hkPolicyLcid);
				hkPolicyLcid = NULL;
			}
			if (NULL != hkSoftwareLcid)
			{
				RegCloseKey(hkSoftwareLcid);
				hkSoftwareLcid = NULL;
			}
		}

		// Now try the office key
		dwSize = sizeof(rgwchMSILCID);
		hr = RegQueryValueExW(hkMAPI1, WszValueNameOfficeLCID, 0, &dwType,
							  (LPBYTE)rgwchMSILCID, &dwSize);
		if ((ERROR_SUCCESS == hr) && (REG_MULTI_SZ == dwType))
		{
			// Read the key
			LPWSTR szKey = NULL;

			// Check policy hive first
			if (RegOpenKeyExW(hkPolicy, rgwchMSILCID, 0, KEY_READ,
							  &hkPolicyLcid) != ERROR_SUCCESS)
			{
				hkPolicyLcid = NULL;
			}

			if (RegOpenKeyExW(hkSoftware, rgwchMSILCID, 0, KEY_READ,
							  &hkSoftwareLcid) != ERROR_SUCCESS)
			{
				hkSoftwareLcid = NULL;
			}
			dwSize = sizeof(lcid);
			dwSize2 = sizeof(lcid);
			szKey = &rgwchMSILCID[wcslen(rgwchMSILCID)+1];
			if ( (hkSoftwareLcid &&
					(RegQueryValueExW(hkSoftwareLcid, szKey, 0, &dwType,  (LPBYTE)&lcid, &dwSize) == ERROR_SUCCESS) &&
					(REG_DWORD == dwType) ) ||
					(hkPolicyLcid &&
					 (RegQueryValueExW(hkPolicyLcid, szKey, 0, &dwType,
									   (LPBYTE)&lcid, &dwSize2) == ERROR_SUCCESS) &&
					 (REG_DWORD == dwType) ))
			{
				goto Done;
			}
			if (NULL != hkPolicyLcid)
			{
				RegCloseKey(hkPolicyLcid);
				hkPolicyLcid = NULL;
			}
			if (NULL != hkSoftwareLcid)
			{
				RegCloseKey(hkSoftwareLcid);
				hkSoftwareLcid = NULL;
			}
		}
		lcid = 0;
	}

Done:
	if (NULL != hkMAPI1)
		RegCloseKey(hkMAPI1);

	if (NULL != hkMAPI2)
		RegCloseKey(hkMAPI2);
	if (NULL != hkPolicy)
		RegCloseKey(hkPolicy);
	if (NULL != hkSoftware)
		RegCloseKey(hkSoftware);
	if (NULL != hkPolicyLcid)
		RegCloseKey(hkPolicyLcid);
	if (NULL != hkSoftwareLcid)
		RegCloseKey(hkSoftwareLcid);

	return lcid;
}


/*
 *  GetInstalledLCID
 *
 *  Purpose:
 *  Retrieves the preferred LCID for UI from the appropriate locations in
 *  the registry.
 *
 *  Arguments:
 *      int          szLangType      Type of LCID to retrieve
 *
 */
LCID GetInstalledLCID(int langType)
{
	LCID  lcid = 0;

	switch (langType)
	{
	case lastUILanguage:
		lcid = GetOutlookLCID();
		break;
		// If we are asked for non-registry based LCIDs, provide them
	case userLanguage:
		lcid = GetUserDefaultLCID();
		break;
	case systemLanguage:
		lcid = GetSystemDefaultLCID();
		break;
	case defaultLanguage:
		lcid = 1033;
		break;
	}

	return lcid;
}



/*
 *  LoadLibraryRegW
 *
 *  Purpose:
 *  Retrieves a path from the registry, adds the provided DLL name (replacing
 *  a file name if already present) and attempts to load the library.
 *
 *  Arguments:
 *      LPCSTR      szDLL        Additional DLL path information.  May be NULL.
 *                                 Will replace any filename (if present)
 *      DWORD       dwFlags      Flags
 *      HKEY        hKey         Key to check
 *      LPCWSTR     wszRegRoot   Sub-key to check
 *      LPCWSTR     wszRegValue  Registry value to actually use
 *
 *  Returns:
 *      HMODULE
 *
 */
inline HMODULE LoadLibraryRegW(LPCWSTR wszDLL, DWORD dwFlags, HKEY hKeyRoot,
								 LPCWSTR wszRegRoot, LPCWSTR wszRegValue)
{
	DWORD           cch;
	DWORD           cchDLL;
	DWORD           dwFileAtt;
	DWORD           dwErr;
	DWORD           dwType;
	HMODULE         hInstRet = NULL;
	HKEY            hKey = NULL;
	LPWSTR          pwszPath;
	WCHAR*          pwchFileName;
	WCHAR           rgwchExpandPath[MAX_PATH];
	WCHAR           rgwchFullPath[MAX_PATH];
	WCHAR           rgwchRegPath[MAX_PATH];
	WCHAR           rgwchTestPath[MAX_PATH];


	// Try to open the registry
	dwErr = RegOpenKeyExW(hKeyRoot, wszRegRoot, 0, KEY_QUERY_VALUE, &hKey);
	if (ERROR_SUCCESS != dwErr)
	{
		goto Exit;
	}

	// Query the Value
	cch = sizeof(rgwchRegPath);
	dwErr = RegQueryValueExW(hKey, wszRegValue, 0, &dwType, (LPBYTE)&rgwchRegPath,
							 &cch);
	if (ERROR_SUCCESS != dwErr)
	{
		goto Exit;
	}
	else if ((REG_SZ != dwType) && (REG_EXPAND_SZ != dwType))
	{
		dwErr = ERROR_INVALID_PARAMETER;
		goto Exit;
	}

	// Set this as the default path - remember that cch from the registry
	//  calls includes the NULL terminator

	pwszPath = rgwchRegPath;

	// Expand any environment strings which may be present

	if (REG_EXPAND_SZ == dwType)
	{
		// Expand the strings

		cch = ExpandEnvironmentStringsW(pwszPath, rgwchExpandPath,
										_countof(rgwchExpandPath));
		if ((0 == cch) || (cch > _countof(rgwchExpandPath)))
		{
			dwErr = ERROR_INSUFFICIENT_BUFFER;
			goto Exit;
		}

		// Reset the pointers.  ExpandEnvironmentStrings returns a bad
		//  length.

		pwszPath = rgwchExpandPath;
	}
	cch = lstrlenW(pwszPath);

	// Check the path the we have to make sure there really is a file there

	dwFileAtt = GetFileAttributesW(pwszPath);
	if (0xFFFFFFFF == dwFileAtt)
	{
		dwErr = GetLastError();
		goto Exit;
	}

	// See if we need to add or replace any information

	if (NULL != wszDLL)
	{
		// If the current path is to a directory then we just need to
		//  add the data at the end

		cchDLL = lstrlenW(wszDLL);
		if (dwFileAtt & FILE_ATTRIBUTE_DIRECTORY)
		{
			// Add a '\' if necessary

			if ((L'\\' != pwszPath[cch-1]) && (L'\\' != wszDLL[0]))
			{
				pwszPath[cch] = L'\\';
				cch++;
			}
			else if ((L'\\' == pwszPath[cch-1]) && (L'\\' == wszDLL[0]))
			{
				cch--;
			}

			// Make sure that the name will fit

			if (MAX_PATH < (cch + cchDLL))
			{
				dwErr = ERROR_INSUFFICIENT_BUFFER;
				goto Exit;
			}

			// Build up the full path in the buffer

			wcscpy_s((pwszPath + cch), MAX_PATH - cch, wszDLL);
		}
		else
		{
			// Locate the filename

			cch = GetFullPathNameW(pwszPath, _countof(rgwchFullPath),
								   rgwchFullPath, &pwchFileName);
			if (0 == cch)
			{
				dwErr = GetLastError();
				goto Exit;
			}
			else if (_countof(rgwchFullPath) < cch)
			{
				dwErr = ERROR_INSUFFICIENT_BUFFER;
				goto Exit;
			}
			pwszPath = rgwchFullPath;

			// Make sure the everything will fit

			cch = (DWORD)(pwchFileName - pwszPath);

			if (MAX_PATH < (cch + cchDLL))
			{
				dwErr = ERROR_INSUFFICIENT_BUFFER;
				goto Exit;
			}

			// Build up the full path in the buffer

			if (L'\\' == *wszDLL)
			{
				wszDLL++;
			}
			wcscpy_s(pwchFileName, MAX_PATH - cch, wszDLL);
		}
	}

	// If we are supposed to, see if the DLL is already loaded in memory

	if (!(dwFlags & LOADLIBMSI_IGNORE_PRIOR))
	{
		// Get the file name from the full path

		cch = GetFullPathNameW(pwszPath, _countof(rgwchTestPath),
							   rgwchTestPath, &pwchFileName);
		pwszPath = rgwchTestPath;

		// Check to see if it is already ordered

		hInstRet = (HMODULE)GetModuleHandleW(pwchFileName);
		if (NULL != hInstRet)
		{
			if (GetModuleFileNameW(hInstRet, rgwchRegPath,
								   _countof(rgwchRegPath)) == 0)
			{
				hInstRet = NULL;
			}
			else
			{
				pwszPath = rgwchRegPath;
			}
		}
	}

	// Try to load the DLL

	hInstRet = LoadLibraryW(pwszPath);

Exit:
	// If we did not get the DLL and we have been told to always load it,
	//  try to see if it is on the path

	if ((NULL == hInstRet) && (NULL != wszDLL) &&
			(dwFlags & LOADLIBMSI_LOAD_ALWAYS))
	{
		// Try to load the DLL based just on the basic DLL information

		hInstRet = LoadLibraryW(wszDLL);
		dwErr = ERROR_SUCCESS;
	}

	if (ERROR_SUCCESS != dwErr)
	{
		SetLastError(dwErr);
	}
	if (hKey)
	{
		RegCloseKey(hKey);
	}
	return hInstRet;
}


/*
 *  FindFileMSIW
 *
 *  Purpose:
 *  Uses MSI to lookup and locate a file based on the published component that
 *  it is in.  Can install the file if the correct MSI flags are passed.
 *
 *  Arguments:
 *      LPCSTR          szFilePath       Path to the file (assumed to be
 *                                        relative to the path returned
 *                                        by the component)- May be NULL
 *      DWORD           dwFlags          Flags
 *      LPCWSTR         wszProduct       Product name for MSI
 *      LPCWSTR         wszComponent     Component name for MSI
 *      LPCWSTR         wszFeatQual      Feature name for standard MSI call,
 *                                        qualifier for qualified call
 *      DWORD           dwInstallMode    MSI installation mode
 *      LPWSTR          wszDest          Resulting path
 *      DWORD*          pcchDest         IN: Size of buffer
 *                                       OUT: Length of path
 *
 *  Returns:
 *      DWORD   Error if any
 *
 */
inline DWORD FindFileMSIW(LPCWSTR wszFilePath, DWORD dwFlags, LPCWSTR wszProduct,
						  LPCWSTR wszComponent, LPCWSTR wszFeatQual, DWORD dwInstallMode,
						  LPWSTR wszDest, DWORD* pcchDest)
{
	DWORD           cchCompPath;
	DWORD           cchDestPath;
	DWORD           dwFileAtt;
	DWORD           dwErr;
	WCHAR*          pwchDestPath;
	WCHAR*          pwchFile;
	WCHAR           rgwchTempPath[MAX_PATH * 2];
	WCHAR           rgwchCompPath[MAX_PATH * 2];
	DWORD           dwInstallModeFirstTry = dwInstallMode;
	bool            fRetry = FALSE;

	//
	// Turn ON the INSTALLMODE_EXISTING optimization if
	// the INSTALLMODE_EXISTING optimization is NOT Overridden,
	//
	if ( (!( dwFlags & FINDFILEMSI_SKIP_OPTIMIZATION_FOR_EXISTING)) &&
			( ((DWORD)INSTALLMODE_DEFAULT) == dwInstallMode) )
	{
		fRetry = TRUE;
		dwInstallModeFirstTry = (DWORD)INSTALLMODE_EXISTING;
	}



	// Use the appropriate format of MSI to locate the component path

	cchCompPath = _countof(rgwchCompPath);
	if (dwFlags & FINDFILEMSI_FIND_QUALIFIED)
	{
		// Call MSI to provide the qualified component

		//
		// If we are given an INSTALLMODE_DEFAULT, then first try INSTALLMODE_EXISTING,
		// and if that fails then try INSTALLMODE_DEFAULT.
		//

		dwErr = MsiProvideQualifiedComponentW(wszComponent, wszFeatQual,
											  dwInstallModeFirstTry, rgwchCompPath,
											  &cchCompPath);

		if ( fRetry && (ERROR_FILE_NOT_FOUND == dwErr) )
		{
			dwErr = MsiProvideQualifiedComponentW(wszComponent, wszFeatQual,
												  dwInstallMode, rgwchCompPath,
												  &cchCompPath);
		}

		if (ERROR_SUCCESS != dwErr)
		{
			goto Exit;
		}
	}
	else
	{
		// Call MSI to provide the feature

		//
		// If we are given an INSTALLMODE_DEFAULT, then first try INSTALLMODE_EXISTING,
		// and if that fails then try INSTALLMODE_DEFAULT.
		//

		dwErr = MsiProvideComponentW(wszProduct, wszFeatQual, wszComponent,
									 dwInstallModeFirstTry, rgwchCompPath,
									 &cchCompPath);

		if ( fRetry && (ERROR_FILE_NOT_FOUND == dwErr) )
		{
			dwErr = MsiProvideComponentW(wszProduct, wszFeatQual, wszComponent,
										 dwInstallMode, rgwchCompPath,
										 &cchCompPath);
		}


		if (ERROR_SUCCESS != dwErr)
		{
			goto Exit;
		}
	}

	// Make sure that we have a good path

	dwFileAtt = GetFileAttributesW(rgwchCompPath);
	if (0xFFFFFFFF == dwFileAtt)
	{
		// We should have successfully installed the file so this failure
		//  should be considered catastrophic

		dwErr = ERROR_FILE_NOT_FOUND;
		goto Exit;
	}

	// If we were handed a DLL name, make sure that we have a path to
	//  it

	if (NULL != wszFilePath)
	{
		DWORD           cchFileName;
		DWORD           cchTempBuff;

		// MSI can return general paths or paths to the principle file, if
		//  we have the file then we want to make sure we replace it with
		//  the correct file name

		if (!(FILE_ATTRIBUTE_DIRECTORY & dwFileAtt))
		{

			// We know we do not have a directory so break the path into
			//  path and file

			cchCompPath = GetFullPathNameW(rgwchCompPath, _countof(rgwchTempPath),
										   rgwchTempPath, &pwchFile);
			if ((0 == cchCompPath) || (_countof(rgwchTempPath) < cchCompPath))
			{
				dwErr = ERROR_FILENAME_EXCED_RANGE;
				goto Exit;
			}
			*pwchFile = L'\0';

			// Record where we pick the path up from

			pwchDestPath = rgwchTempPath;
			cchDestPath = (DWORD)(pwchFile - rgwchTempPath);
		}
		else
		{
			pwchDestPath = rgwchCompPath;
			cchDestPath = cchCompPath;
		}

		// Make sure that the path name is properly terminated

		if (L'\\' != pwchDestPath[cchDestPath - 1])
		{
			pwchDestPath[cchDestPath] = L'\\';
			cchDestPath++;
		}

		// Build up a path name

		cchTempBuff = (DWORD) min(_countof(rgwchCompPath), _countof(rgwchTempPath));
		cchFileName = lstrlenW(wszFilePath);
		if (cchTempBuff <= (cchDestPath + cchFileName))
		{
			dwErr = ERROR_FILENAME_EXCED_RANGE;
			goto Exit;
		}
		wcscpy_s(&(pwchDestPath[cchDestPath]), cchTempBuff - cchDestPath, wszFilePath);
	}
	else
	{
		// Make sure we have a file and not a directory

		if (FILE_ATTRIBUTE_DIRECTORY & dwFileAtt)
		{
			dwErr = ERROR_FILE_NOT_FOUND;
			goto Exit;
		}

		// Set-up the pointers for the final cleanup

		pwchDestPath = rgwchCompPath;
		cchDestPath = cchCompPath;
	}

	// Locate the file name from the path

	cchDestPath = GetFullPathNameW(pwchDestPath, *pcchDest, wszDest, &pwchFile);
	if (0 == cchDestPath)
	{
		dwErr = GetLastError();
		goto Exit;
	}
	else if (*pcchDest < cchDestPath)
	{
		dwErr = ERROR_FILENAME_EXCED_RANGE;
		*pcchDest = cchDestPath;
		goto Exit;
	}
	dwErr = ERROR_SUCCESS;
	*pcchDest = cchDestPath;

Exit:
	return dwErr;
}


/*
 *  LoadLibraryMSIExW
 *
 *  Purpose:
 *  Loads a DLL into the current process space after using the MS Windows
 *  Installer to install the DLL if necessary.
 *
 *  Arguments:
 *      LPCWSTR         wszDLLPath      DLL to load (may be a relative path)
 *      HANDLE          hFile           Must be NULL
 *      DWORD           dwFlags         Standard LoadLibraryEx flags
 *      LPCWSTR         wszProduct      Product name for MSI
 *      LPCWSTR         wszComponent    Component name for MSI
 *      LPCWSTR         wszFeatQual     Feature name for standard MSI call,
 *                                        qualifier if doing a qualified
 *                                        install.
 *      DWORD           dwInstallMode   MSI installation mode
 *
 */

inline HMODULE LoadLibraryMSIExW(LPCWSTR wszDLLPath, HANDLE /* hFile */,
								   DWORD dwFlags, LPCWSTR wszProduct,
								   LPCWSTR wszComponent, LPCWSTR wszFeatQual,
								   DWORD dwInstallMode)
{
	DWORD           cchCompPath;
	DWORD           dwErr = ERROR_SUCCESS;
	HMODULE         hInstDLL = NULL;
	WCHAR           rgwchCompPath[MAX_PATH];

	// Get a path to the file

	cchCompPath = _countof(rgwchCompPath);
	dwErr = FindFileMSIW(wszDLLPath,
						 ((dwFlags & LOADLIBMSI_LOAD_QUALIFIED) ?
						  FINDFILEMSI_FIND_QUALIFIED : 0),
						 wszProduct, wszComponent, wszFeatQual, dwInstallMode,
						 rgwchCompPath, &cchCompPath);
	if (ERROR_SUCCESS != dwErr)
	{
		goto Exit;
	}

	// If we are supposed to, see if the DLL is already loaded in memory

	if (!(dwFlags & LOADLIBMSI_IGNORE_PRIOR))
	{
		DWORD           cchFullPath;
		WCHAR*          pwchFile;
		WCHAR           rgwchFullPath[MAX_PATH];

		// Get the file name from the full path

		cchFullPath = GetFullPathNameW(rgwchCompPath, _countof(rgwchFullPath),
									   rgwchFullPath, &pwchFile);

		// Check to see if it is already ordered

		hInstDLL = (HMODULE)GetModuleHandleW(pwchFile);
		if (NULL != hInstDLL)
		{
			if (GetModuleFileNameW(hInstDLL, rgwchCompPath,
								   _countof(rgwchCompPath)) == 0)
			{
				hInstDLL = NULL;
			}
		}
	}

	// Load the DLL

	hInstDLL = LoadLibraryExW(rgwchCompPath, NULL,
							  (dwFlags & ~LOADLIBMSI_FLAG_MASK));

Exit:
	// See if we are supposed to try and always load

	if ((NULL == hInstDLL) && (dwFlags & LOADLIBMSI_LOAD_ALWAYS))
	{
		hInstDLL = LoadLibraryExW(wszDLLPath, NULL,
								  (dwFlags & ~LOADLIBMSI_FLAG_MASK));
		dwErr = ERROR_SUCCESS;
	}

	// Record any error that we have

	if (ERROR_SUCCESS != dwErr)
	{
		SetLastError(dwErr);
	}
	return hInstDLL;
}

/*
 *  LoadLibraryLangMSIW
 *
 *  Purpose:
 *  Loads the specified DLL using language base qualifiers to make sure the
 *  correct UI language is loaded.
 *
 *  Arguments:
 *      LPCWSTR         wszDLLPath      DLL to load (may be a relative path)
 *      DWORD           dwFlags         LoadLibraryMSI flags
 *      LPCWSTR         wszComponent    Component name for MSI
 *      LPCWSTR         wszQualFormat   'printf' style string for formatting
 *                                        the qualifier with the LCID
 *      DWORD           dwInstallMode   MSI installation mode
 *
 */
HMODULE LoadLibraryLangMSIW(LPCWSTR wszDLLPath, DWORD dwFlags,
							  LPCWSTR wszComponent, LPCWSTR wszQualFormat,
							  DWORD dwInstallMode, LCID* plcid)
{
	HMODULE         hInstDLL = NULL;
	int             iLCID;
	LCID            lcid = 0;
	WCHAR           rgwchQualifier[32];

	// Try to use MSI to locate the correct DLL to use

	*plcid = 0;
	for (iLCID = 0; iLCID < MAX_LANGLOADORDER; iLCID++)
	{
		// Retrieve the current UI LCID and make sure it is valid

		lcid = GetInstalledLCID(iLCID);
		if (0 == lcid)
		{
			continue;
		}

		// Build up the qualifier string
		swprintf_s(rgwchQualifier, _countof(rgwchQualifier), wszQualFormat, lcid);
		hInstDLL = LoadLibraryMSIExW(wszDLLPath, NULL,
									 (dwFlags | LOADLIBMSI_LOAD_QUALIFIED),
									 NULL, wszComponent, rgwchQualifier,
									 dwInstallMode);
		if (NULL != hInstDLL)
		{
			break;
		}
	}
	if (hInstDLL != NULL)
	{
		*plcid = lcid;
	}
	return hInstDLL;
}



static volatile HMODULE		g_hinstMAPI = NULL;


__inline HMODULE GetMAPIHandle()
{
	return g_hinstMAPI;
}

void SetMAPIHandle(HMODULE hinstMAPI)
{
	HMODULE	hinstNULL = NULL;
	HMODULE	hinstToFree = NULL;

	if (hinstMAPI == NULL)
	{
		hinstToFree = (HMODULE)InterlockedExchangePointer((PVOID*)&g_hinstMAPI, hinstNULL);
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
}

__inline bool __stdcall FFileExistW( LPCWSTR wzFileName )
{
	DWORD attr;

	if (NULL == wzFileName)
		return false;

	attr = GetFileAttributesW( wzFileName );

	return ( ( attr != INVALID_FILE_ATTRIBUTES ) && ( ( attr & FILE_ATTRIBUTE_DIRECTORY ) == 0 ) );
}


bool FCanUseDefaultOldMapiPath(HMODULE * pInstMAPI)
{
	HRESULT             hr;
	DWORD               cbSize;
	DWORD               dwType;
	HKEY                hkMAPI = NULL;
	int                 idxLCID;
	LCID                lcid;
	WCHAR               rgwchMSIComp[(sizeof(GUID) * 2) + 10];
	WCHAR               rgwchQualFormat[MAX_PATH];
	HMODULE             hMAPI = 0;

	// Attempt to avoid MSI and pick up PrivateMAPI directly
	for (idxLCID = 0; idxLCID < MAX_LANGLOADORDER; idxLCID++)
	{
		lcid = GetInstalledLCID(idxLCID);
		if (0 != lcid)
		{
			swprintf_s(rgwchQualFormat, _countof(rgwchQualFormat), WszMAPIPathFormat, lcid,
					   WszMSMAPI32DLL);
			hMAPI = LoadLibraryRegW(rgwchQualFormat,
									LOADLIBMSI_IGNORE_PRIOR,
									HKEY_LOCAL_MACHINE, WszCurVersion,
									WszCommonFiles);

			if (NULL != hMAPI)
			{
				break;
			}
		}
	}

	// If we cannot pick up MAPI directly, consider using the MSI
	//  component registered in the Outlook mail clients section
	//  to try and locate MAPI or OMI

	if (NULL == hMAPI)
	{
		if ((RegOpenKeyExW(HKEY_CURRENT_USER, WszKeyNameMSI, 0, KEY_READ,
						   &hkMAPI) == ERROR_SUCCESS) ||
				(RegOpenKeyExW(HKEY_LOCAL_MACHINE, WszKeyNameMSI, 0, KEY_READ,
							   &hkMAPI) == ERROR_SUCCESS))
		{
			cbSize = sizeof(rgwchMSIComp);
			hr = RegQueryValueExW(hkMAPI, WszValueNameMSI, 0, &dwType,
								  (LPBYTE)rgwchMSIComp, &cbSize);
			if ((ERROR_SUCCESS == hr) && (REG_SZ == dwType))
			{
				// Try to use MSI to locate and load the correct DLL
				wcscpy_s(rgwchQualFormat, _countof(rgwchQualFormat), WszFullQualifier);
				hMAPI = LoadLibraryLangMSIW(NULL, 0, rgwchMSIComp,
											rgwchQualFormat,
											(DWORD)INSTALLMODE_EXISTING,
											&lcid);
				if (NULL == hMAPI)
				{
					// Try to get the component without the qualification
					wcscpy_s(rgwchQualFormat, _countof(rgwchQualFormat), WszShortQualifier);
					hMAPI = LoadLibraryLangMSIW(NULL, 0, rgwchMSIComp,
												rgwchQualFormat,
												(DWORD)INSTALLMODE_EXISTING,
												&lcid);
				}
			}
			RegCloseKey(hkMAPI);
		}
	}

	if (hMAPI)
	{
		*pInstMAPI = hMAPI;
		return TRUE;
	}
	else
	{
		return FALSE;
	}
}

/*
 *  RegQueryWszExpand
 *		Wrapper for RegQueryValueExW which automatically expands REG_EXPAND_SZ values
 */
DWORD RegQueryWszExpand(HKEY hKey, LPCWSTR lpValueName, LPWSTR lpValue, DWORD cchValueLen)
{
	DWORD dwErr = ERROR_SUCCESS;
	DWORD dwType = 0;

	WCHAR rgchValue[MAX_PATH];
	DWORD dwSize = sizeof(rgchValue);

	dwErr = RegQueryValueExW(hKey, lpValueName, 0, &dwType, (LPBYTE) &rgchValue, &dwSize);

	if (dwErr == ERROR_SUCCESS)
	{
		if (dwType == REG_EXPAND_SZ)
		{
			// Expand the strings
			DWORD cch = ExpandEnvironmentStringsW(rgchValue, lpValue, cchValueLen);
			if ((0 == cch) || (cch > cchValueLen))
			{
				dwErr = ERROR_INSUFFICIENT_BUFFER;
				goto Exit;
			}
		}
		else if (dwType == REG_SZ)
		{
			wcscpy_s(lpValue, cchValueLen, rgchValue);
		}
	}
Exit:
	return dwErr;
}

/*
 *  GetComponentPath
 *		Wrapper around mapi32.dll->FGetComponentPath which maps an MSI component ID to 
 *		a DLL location from the default MAPI client registration values
 *
 */
bool GetComponentPath(LPCSTR szComponent, LPSTR szQualifier, LPSTR szDllPath, DWORD cchBufferSize, bool fInstall)
{
	HMODULE hMapiStub = NULL;
	bool fReturn = FALSE;

	typedef bool (STDAPICALLTYPE *FGetComponentPathType)(LPCSTR, LPSTR, LPSTR, DWORD, bool);

	hMapiStub = LoadLibraryW(WszMapi32);
	if (!hMapiStub)
		hMapiStub = LoadLibraryW(WszMapiStub);

	if (hMapiStub)
	{
		FGetComponentPathType pFGetCompPath = (FGetComponentPathType)GetProcAddress(hMapiStub, SzFGetComponentPath);

		fReturn = pFGetCompPath(szComponent, szQualifier, szDllPath, cchBufferSize, fInstall);

		FreeLibrary(hMapiStub);
	}

	return fReturn;
}


/*
 *  LoadMailClientFromMSIData
 *		Attempt to locate the MAPI provider DLL via HKLM\Software\Clients\Mail\(provider)\MSIComponentID
 */
HMODULE LoadMailClientFromMSIData(HKEY hkeyMapiClient)
{
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
			hinstMapi = LoadLibraryA(rgchComponentPath);
		}
	}
	return hinstMapi;
}

/*
 *  LoadMailClientFromDllPath
 *		Attempt to locate the MAPI provider DLL via HKLM\Software\Clients\Mail\(provider)\DllPathEx
 */
HMODULE LoadMailClientFromDllPath(HKEY hkeyMapiClient)
{
	HMODULE		hinstMapi = NULL;
	WCHAR		rgchDllPath[MAX_PATH];

	DWORD dwSizeDllPath = _countof(rgchDllPath);

	if (ERROR_SUCCESS == RegQueryWszExpand(	hkeyMapiClient, WszValueNameDllPathEx, rgchDllPath, dwSizeDllPath))
	{
		hinstMapi = LoadLibraryW(rgchDllPath);
	}

	if (!hinstMapi)
	{
		dwSizeDllPath = _countof(rgchDllPath);
		if (ERROR_SUCCESS == RegQueryWszExpand(	hkeyMapiClient, WszValueNameDllPath, rgchDllPath, dwSizeDllPath))
		{
			hinstMapi = LoadLibraryW(rgchDllPath);
		}
	}
	return hinstMapi;
}

/*
 *  LoadDefaultNonOutlookMapiClient
 *		Read the registry to discover the registered MAPI client and if it is not Outlook
 *		attempt to load it's MAPI DLL.
 */
HMODULE LoadDefaultNonOutlookMapiClient()
{
	HMODULE		hinstMapi = NULL;
	DWORD		dwType;
	HKEY 		hkey = NULL, hkeyMapiClient = NULL;
	WCHAR		rgchMailClient[MAX_PATH];

	// Open HKLM\Software\Clients\Mail
	if (ERROR_SUCCESS == RegOpenKeyExW( HKEY_LOCAL_MACHINE,
										WszKeyNameMailClient,
										0,
										KEY_READ,
										&hkey))
	{
		// Get Outlook application path registry value
		DWORD dwSize = sizeof(rgchMailClient);
		if SUCCEEDED(RegQueryValueExW(	hkey,
										NULL,
										0,
										&dwType,
										(LPBYTE) &rgchMailClient,
										&dwSize))
		{
			if (dwType != REG_SZ)
				goto Error;

			if (0 == wcscmp(rgchMailClient, WszOutlookMapiClientName))
				goto Error;

			if SUCCEEDED(RegOpenKeyExW( hkey,
				rgchMailClient,
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
	return hinstMapi;
}

/*
 *  LoadMAPIFromSystemDir
 *		Fall back for loading System32\Mapi32.dll if all else fails
 */
HMODULE LoadMAPIFromSystemDir()
{
	WCHAR szSystemDir[MAX_PATH] = {0};

	if (GetSystemDirectoryW(szSystemDir, MAX_PATH))
	{
		WCHAR szDLLPath[MAX_PATH] = {0};
		swprintf_s(szDLLPath, _countof(szDLLPath), WszMAPISystemPath, szSystemDir, WszMapi32);
		return LoadLibraryW(szDLLPath);
	}

	return NULL;
}


HMODULE GetDefaultMapiHandle()
{
	size_t	cchRemain;
	DWORD	dwSize;
	DWORD	dwType;
	bool	fDefault = FALSE;
	HKEY 	hkey = NULL;
	WCHAR	rgchOutlookDir[MAX_PATH];
	LPWSTR	wzRemain;
	HMODULE	hinstMapi = NULL;
	WCHAR	wzPath[MAX_PATH];
	DWORD	cchMax = _countof(wzPath);

	// Try to respect the machine's default MAPI client settings.  If the active MAPI provider
	//  is Outlook, don't load and instead run the logic below
	if (!s_fForceOutlookMAPI)
		hinstMapi = LoadDefaultNonOutlookMapiClient();

	// If we've successfully loaded the default MAPI provider, then we are done 
	if (hinstMapi)
		goto Done;

	// Open HKLM\Software\Microsoft\Windows\CurrentVersion
	if SUCCEEDED(RegOpenKeyExW( HKEY_LOCAL_MACHINE,
								WszOutlookAppPath,
								0,
								KEY_READ,
								&hkey))
	{
		// Get Outlook application path registry value
		dwSize = sizeof(rgchOutlookDir);
		if SUCCEEDED(RegQueryValueExW(	hkey,
										NULL,
										0,
										&dwType,
										(LPBYTE) &rgchOutlookDir,
										&dwSize))
		{
			if (dwType != REG_SZ)
				goto Error;

			// remove outlook.exe from end of path
			_wcslwr_s(rgchOutlookDir, _countof(rgchOutlookDir));
			WCHAR * pCutHere = wcsstr(rgchOutlookDir, WszOutlookExe);
			if (pCutHere)
				pCutHere[0] = L'\0';
		}
		// copy the path to the output buffer and append the MAPI dll name
		if SUCCEEDED(StringCchCopyExW(wzPath, cchMax, rgchOutlookDir, &wzRemain, &cchRemain, 0UL))
		{
			UINT uiOldErrMode;

			uiOldErrMode = SetErrorMode(0);
			// Set the DLL directory so that private MAPI can find all of its load time linked DLLs.
			if (FALSE == SetDllDirectoryW(wzPath))
			{
				// If we fail, then we can't use this path for loading dlls
				fDefault = FALSE;
				goto Error;
			}
			if (SUCCEEDED(wcscpy_s(wzRemain, cchRemain, WszPrivateMAPI)))
			{
				fDefault = FFileExistW(wzPath); // return the existence of the file
			}
			// If we failed to load olmapi32, try msmapi32
			if (!fDefault && FCanUseDefaultOldMapiPath(&hinstMapi))
			{
				fDefault = FFileExistW(wzPath); // return the existence of the file
			}

			if (fDefault)
			{
				hinstMapi = LoadLibraryW(wzPath);
			}
			// Restore the DLL path
			SetDllDirectoryW(NULL);
		}
	}

	// If MAPI still isn't loaded, load the stub from the system directory
	if (!hinstMapi && !s_fForceOutlookMAPI)
	{
		hinstMapi = LoadMAPIFromSystemDir();
	}

Done:
Error:
	if (hkey)
	{
		RegCloseKey(hkey);
	}

	return hinstMapi;
}


/*------------------------------------------------------------------------------
    Attach to wzMapiDll(olmapi32.dll/msmapi32.dll) if it is already loaded in the
	current process.
------------------------------------------------------------------------------*/
HMODULE AttachToMAPIDll(const WCHAR *wzMapiDll)
{
	HMODULE	hinstPrivateMAPI = NULL;
	GetModuleHandleExW(0UL, wzMapiDll, &hinstPrivateMAPI);
	return hinstPrivateMAPI;
}


void UnLoadPrivateMAPI()
{
	HMODULE hinstPrivateMAPI = NULL;

	hinstPrivateMAPI = GetMAPIHandle();
	if (NULL != hinstPrivateMAPI)
	{
		SetMAPIHandle(NULL);
	}
}

void ForceOutlookMAPI(bool fForce)
{
	s_fForceOutlookMAPI = fForce;
}


HMODULE GetPrivateMAPI()
{
	HMODULE hinstPrivateMAPI = GetMAPIHandle();

	if (NULL == hinstPrivateMAPI)
	{
		// First, try to attach to olmapi32.dll if it's loaded in the process
		hinstPrivateMAPI = AttachToMAPIDll(WszPrivateMAPI);

		// If that fails try msmapi32.dll, for Outlook 11 and below
		//  Only try this in the static lib, otherwise msmapi32.dll will attach to itself.
		if (NULL == hinstPrivateMAPI)
		{
			hinstPrivateMAPI = AttachToMAPIDll(WszPrivateMAPI_11);
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
		return GetMAPIHandle();
	}

	return hinstPrivateMAPI;
}
