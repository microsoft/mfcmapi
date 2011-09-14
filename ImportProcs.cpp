#include "stdafx.h"
#include "ImportProcs.h"
#include "MAPIFunctions.h"

HMODULE	hModAclui = NULL;
HMODULE hModRichEd20 = NULL;
HMODULE hModOle32 = NULL;
HMODULE hModUxTheme = NULL;
HMODULE hModInetComm = NULL;
HMODULE hModMSI = NULL;
HMODULE hModKernel32 = NULL;

typedef HTHEME (STDMETHODCALLTYPE CLOSETHEMEDATA)
(
 HTHEME hTheme);
typedef CLOSETHEMEDATA* LPCLOSETHEMEDATA;

typedef bool (WINAPI HEAPSETINFORMATION) (
    HANDLE HeapHandle,
    HEAP_INFORMATION_CLASS HeapInformationClass,
    PVOID HeapInformation,
    SIZE_T HeapInformationLength);
typedef HEAPSETINFORMATION* LPHEAPSETINFORMATION;

typedef bool (WINAPI GETMODULEHANDLEEXW) (
    DWORD    dwFlags,
    LPCWSTR lpModuleName,
    HMODULE* phModule);
typedef GETMODULEHANDLEEXW* LPGETMODULEHANDLEEXW;

typedef HRESULT (STDMETHODCALLTYPE MIMEOLEGETCODEPAGECHARSET)
(
 CODEPAGEID cpiCodePage,
 CHARSETTYPE ctCsetType,
 LPHCHARSET phCharset
 );
typedef MIMEOLEGETCODEPAGECHARSET* LPMIMEOLEGETCODEPAGECHARSET;

typedef UINT (WINAPI MSIPROVIDECOMPONENTW)
(
 LPCWSTR szProduct,
 LPCWSTR szFeature,
 LPCWSTR szComponent,
 DWORD   dwInstallMode,
 LPWSTR  lpPathBuf,
 LPDWORD pcchPathBuf
 );
typedef MSIPROVIDECOMPONENTW FAR * LPMSIPROVIDECOMPONENTW;

typedef UINT (WINAPI MSIPROVIDEQUALIFIEDCOMPONENTW)
(
 LPCWSTR szCategory,
 LPCWSTR szQualifier,
 DWORD   dwInstallMode,
 LPWSTR  lpPathBuf,
 LPDWORD pcchPathBuf
 );
typedef MSIPROVIDEQUALIFIEDCOMPONENTW FAR * LPMSIPROVIDEQUALIFIEDCOMPONENTW;

LPEDITSECURITY				pfnEditSecurity = NULL;

LPSTGCREATESTORAGEEX		pfnStgCreateStorageEx = NULL;

LPOPENTHEMEDATA				pfnOpenThemeData = NULL;
LPCLOSETHEMEDATA			pfnCloseThemeData = NULL;
LPGETTHEMEMARGINS			pfnGetThemeMargins = NULL;

// From inetcomm.dll
LPMIMEOLEGETCODEPAGECHARSET pfnMimeOleGetCodePageCharset = NULL;

// From MSI.dll
LPMSIPROVIDEQUALIFIEDCOMPONENT pfnMsiProvideQualifiedComponent = NULL;
LPMSIGETFILEVERSION pfnMsiGetFileVersion = NULL;
LPMSIPROVIDECOMPONENTW pfnMsiProvideComponentW = NULL;
LPMSIPROVIDEQUALIFIEDCOMPONENTW pfnMsiProvideQualifiedComponentW = NULL;

// From kernel32.dll
LPHEAPSETINFORMATION pfnHeapSetInformation = NULL;
LPGETMODULEHANDLEEXW pfnGetModuleHandleExW = NULL;

// Exists to allow some logging
_Check_return_ HMODULE MyLoadLibrary(_In_z_ LPCTSTR lpszLibFileName)
{
	HMODULE hMod = NULL;
	HRESULT hRes = S_OK;
	DebugPrint(DBGLoadLibrary,_T("MyLoadLibrary - loading \"%s\"\n"),lpszLibFileName);
	WC_D(hMod,LoadLibrary(lpszLibFileName));
	if (hMod)
	{
		DebugPrint(DBGLoadLibrary,_T("MyLoadLibrary - \"%s\" loaded at %p\n"),lpszLibFileName,hMod);
	}
	else
	{
		DebugPrint(DBGLoadLibrary,_T("MyLoadLibrary - \"%s\" failed to load\n"),lpszLibFileName);
	}
	return hMod;
} // MyLoadLibrary

// Loads szModule at the handle given by lphModule, then looks for szEntryPoint.
// Will not load a module or entry point twice
void LoadProc(LPTSTR szModule, HMODULE* lphModule, LPSTR szEntryPoint, FARPROC* lpfn)
{
	if (!szEntryPoint || !lpfn || !lphModule) return;
	if (*lpfn) return;
	if (!*lphModule && szModule)
	{
		*lphModule = LoadFromSystemDir(szModule);
	}
	if (!*lphModule) return;

	HRESULT hRes = S_OK;
	WC_D(*lpfn, GetProcAddress(
		*lphModule,
		szEntryPoint));
} // LoadProc

// We do this to avoid certain problems Outlook 11's MAPI has when the system RichEd is loaded.
// Note that we don't worry about ever unloading this
void LoadRichEd()
{
	if (hModRichEd20) return;
	HRESULT hRes = S_OK;
	HKEY hCurrentVersion = NULL;

	WC_W32(RegOpenKeyEx(
		HKEY_LOCAL_MACHINE,
		_T("Software\\Microsoft\\Windows\\CurrentVersion"), // STRING_OK
		NULL,
		KEY_READ,
		&hCurrentVersion));
	hRes = S_OK;

	if (hCurrentVersion)
	{
		DWORD dwKeyType = NULL;
		LPTSTR szCF = NULL;

		WC_H(HrGetRegistryValue(
			hCurrentVersion,
			_T("CommonFilesDir"), // STRING_OK
			&dwKeyType,
			(LPVOID*) &szCF));
		hRes = S_OK;

		if (szCF)
		{
			size_t cchCF = NULL;
			size_t cchOffice11 = NULL;

			LPTSTR szOffice11RichEd = _T("\\Microsoft Shared\\office11\\riched20.dll"); // STRING_OK

			EC_H(StringCchLength(szCF,STRSAFE_MAX_CCH,&cchCF));
			EC_H(StringCchLength(szOffice11RichEd,STRSAFE_MAX_CCH,&cchOffice11));

			LPTSTR szRichEdFullPath = NULL;
			szRichEdFullPath = new TCHAR[cchOffice11+cchCF+1];

			if (szRichEdFullPath)
			{
				EC_H(StringCchPrintf(szRichEdFullPath,cchOffice11+cchCF+1,_T("%s%s"),szCF,szOffice11RichEd)); // STRING_OK
				hModRichEd20 = MyLoadLibrary(szRichEdFullPath);
				delete[] szRichEdFullPath;
			}

			if (!hModRichEd20)
			{
				LPTSTR szOffice11RichEdDebug = _T("\\Microsoft Shared Debug\\office11\\riched20.dll"); // STRING_OK
				EC_H(StringCchLength(szOffice11RichEdDebug,STRSAFE_MAX_CCH,&cchOffice11));

				LPTSTR szRichEdDebugFullPath = NULL;
				szRichEdDebugFullPath = new TCHAR[cchOffice11+cchCF+1];

				if (szRichEdDebugFullPath)
				{
					EC_H(StringCchPrintf(szRichEdDebugFullPath,cchOffice11+cchCF+1,_T("%s%s"),szCF,szOffice11RichEdDebug)); // STRING_OK
					hModRichEd20 = MyLoadLibrary(szRichEdDebugFullPath);
					delete[] szRichEdDebugFullPath;
				}
			}

			delete[] szCF;
		}

		EC_W32(RegCloseKey(hCurrentVersion));
	}
} // LoadRichEd

_Check_return_ HMODULE LoadFromSystemDir(_In_z_ LPTSTR szDLLName)
{
	if (!szDLLName) return NULL;

	HRESULT	hRes = S_OK;
	HMODULE	hModRet = NULL;
	TCHAR	szDLLPath[MAX_PATH] = {0};
	UINT	uiRet = NULL;

	static TCHAR	szSystemDir[MAX_PATH] = {0};
	static bool		bSystemDirLoaded = false;

	DebugPrint(DBGLoadLibrary,_T("LoadFromSystemDir - loading \"%s\"\n"),szDLLName);

	if (!bSystemDirLoaded)
	{
		WC_D(uiRet,GetSystemDirectory(szSystemDir, MAX_PATH));
		bSystemDirLoaded = true;
	}

	WC_H(StringCchPrintf(szDLLPath,_countof(szDLLPath),_T("%s\\%s"),szSystemDir,szDLLName)); // STRING_OK
	DebugPrint(DBGLoadLibrary,_T("LoadFromSystemDir - loading from \"%s\"\n"),szDLLPath);
	hModRet = MyLoadLibrary(szDLLPath);

	return hModRet;
} // LoadFromSystemDir

void ImportProcs()
{
	LoadProc(_T("aclui.dll"),&hModAclui,"EditSecurity", (FARPROC*) &pfnEditSecurity); // STRING_OK;
	LoadProc(_T("ole32.dll"), &hModOle32, "StgCreateStorageEx", (FARPROC*) &pfnStgCreateStorageEx); // STRING_OK;
	LoadProc(_T("uxtheme.dll"), &hModUxTheme, "OpenThemeData", (FARPROC*) &pfnOpenThemeData); // STRING_OK;
	LoadProc(_T("uxtheme.dll"), &hModUxTheme, "CloseThemeData", (FARPROC*) &pfnCloseThemeData); // STRING_OK;
	LoadProc(_T("uxtheme.dll"), &hModUxTheme, "GetThemeMargins", (FARPROC*) &pfnGetThemeMargins); // STRING_OK;
#ifdef _UNICODE
	LoadProc(_T("msi.dll"), &hModMSI, "MsiProvideQualifiedComponentW", (FARPROC*) &pfnMsiProvideQualifiedComponent); // STRING_OK;
	LoadProc(_T("msi.dll"), &hModMSI, "MsiGetFileVersionW", (FARPROC*) &pfnMsiGetFileVersion); // STRING_OK;
#else
	LoadProc(_T("msi.dll"), &hModMSI, "MsiProvideQualifiedComponentA", (FARPROC*) &pfnMsiProvideQualifiedComponent); // STRING_OK;
	LoadProc(_T("msi.dll"), &hModMSI, "MsiGetFileVersionA", (FARPROC*) &pfnMsiGetFileVersion); // STRING_OK;
#endif
} // ImportProcs

// Opens the mail key for the specified MAPI client, such as 'Microsoft Outlook' or 'ExchangeMAPI'
// Pass NULL to open the mail key for the default MAPI client
_Check_return_ HKEY GetMailKey(_In_opt_z_ LPCTSTR szClient)
{
	DebugPrint(DBGLoadLibrary,_T("Enter GetMailKey(%s)\n"),szClient?szClient:_T("Default"));
	HRESULT hRes = S_OK;
	HKEY hMailKey = NULL;
	bool bClientIsDefault = false;

	// If szClient is NULL, we need to read the name of the default MAPI client
	if (!szClient)
	{
		HKEY hDefaultMailKey = NULL;
		WC_W32(RegOpenKeyEx(
			HKEY_LOCAL_MACHINE,
			_T("Software\\Clients\\Mail"), // STRING_OK
			NULL,
			KEY_READ,
			&hDefaultMailKey));
		if (hDefaultMailKey)
		{
			DWORD dwKeyType = NULL;
			WC_H(HrGetRegistryValue(
				hDefaultMailKey,
				_T(""), // get the default value
				&dwKeyType,
				(LPVOID*) &szClient));
			DebugPrint(DBGLoadLibrary,_T("Default MAPI = %s\n"),szClient?szClient:_T("Default"));
			bClientIsDefault = true;
			EC_W32(RegCloseKey(hDefaultMailKey));
		}
	}

	if (szClient)
	{
		TCHAR szMailKey[256];
		EC_H(StringCchPrintf(
			szMailKey,
			_countof(szMailKey),
			_T("Software\\Clients\\Mail\\%s"), // STRING_OK
			szClient));

		if (SUCCEEDED(hRes))
		{
			WC_W32(RegOpenKeyEx(
				HKEY_LOCAL_MACHINE,
				szMailKey,
				NULL,
				KEY_READ,
				&hMailKey));
		}
	}
	if (bClientIsDefault) delete[] szClient;

	return hMailKey;
} // GetMailKey

// Gets MSI IDs for the specified MAPI client, such as 'Microsoft Outlook' or 'ExchangeMAPI'
// Pass NULL to get the IDs for the default MAPI client
// Allocates with new, delete with delete[]
void GetMapiMsiIds(_In_opt_z_ LPCTSTR szClient, _Deref_out_opt_z_ LPSTR* lpszComponentID, _Deref_out_opt_z_ LPSTR* lpszAppLCID, _Deref_out_opt_z_ LPSTR* lpszOfficeLCID)
{
	DebugPrint(DBGLoadLibrary,_T("GetMapiMsiIds(%s)\n"),szClient);
	HRESULT hRes = S_OK;

	if (lpszComponentID) *lpszComponentID = 0;
	if (lpszAppLCID) *lpszAppLCID = 0;
	if (lpszOfficeLCID) *lpszOfficeLCID = 0;

	HKEY hKey = GetMailKey(szClient);

	if (hKey)
	{
		DWORD dwKeyType = NULL;

		if (lpszComponentID)
		{
			WC_H(HrGetRegistryValueA(
				hKey,
				"MSIComponentID", // STRING_OK
				&dwKeyType,
				(LPVOID*) lpszComponentID));
			DebugPrint(DBGLoadLibrary,_T("MSIComponentID = %hs\n"),*lpszComponentID?*lpszComponentID:"<not found>");
			hRes = S_OK;
		}

		if (lpszAppLCID)
		{
			WC_H(HrGetRegistryValueA(
				hKey,
				"MSIApplicationLCID", // STRING_OK
				&dwKeyType,
				(LPVOID*) lpszAppLCID));
			DebugPrint(DBGLoadLibrary,_T("MSIApplicationLCID = %hs\n"),*lpszAppLCID?*lpszAppLCID:"<not found>");
			hRes = S_OK;
		}

		if (lpszOfficeLCID)
		{
			WC_H(HrGetRegistryValueA(
				hKey,
				"MSIOfficeLCID", // STRING_OK
				&dwKeyType,
				(LPVOID*) lpszOfficeLCID));
			DebugPrint(DBGLoadLibrary,_T("MSIOfficeLCID = %hs\n"),*lpszOfficeLCID?*lpszOfficeLCID:"<not found>");
			hRes = S_OK;
		}

		EC_W32(RegCloseKey(hKey));
	}
} // GetMapiMsiIds

void GetMAPIPath(_In_opt_z_ LPCTSTR szClient, _Inout_z_count_(cchMAPIPath) LPTSTR szMAPIPath, ULONG cchMAPIPath)
{
	HRESULT hRes = S_OK;

	szMAPIPath[0] = '\0'; // Terminate String at pos 0 (safer if we fail below)

	// Find some strings:
	LPSTR szComponentID = NULL;
	LPSTR szAppLCID = NULL;
	LPSTR szOfficeLCID = NULL;

	GetMapiMsiIds(szClient,&szComponentID,&szAppLCID,&szOfficeLCID);

	if (szComponentID)
	{
#ifdef UNICODE
		CHAR lpszPath[MAX_PATH] = {0};
		ULONG cchPath = _countof(lpszPath);
#else
		LPSTR lpszPath = szMAPIPath;
		ULONG cchPath = cchMAPIPath;
#endif

		if (szAppLCID)
		{
			WC_B(GetComponentPath(
				szComponentID,
				szAppLCID,
				lpszPath,
				cchPath,
				false));
		}
		if ((FAILED(hRes) || lpszPath[0] == _T('\0')) && szOfficeLCID)
		{
			hRes = S_OK;
			WC_B(GetComponentPath(
				szComponentID,
				szOfficeLCID,
				lpszPath,
				cchPath,
				false));
		}
		if (FAILED(hRes) || lpszPath[0] == _T('\0'))
		{
			hRes = S_OK;
			WC_B(GetComponentPath(
				szComponentID,
				NULL,
				lpszPath,
				cchPath,
				false));
		}
#ifdef UNICODE
		int iRet = 0;
		// Convert to Unicode.
		EC_D(iRet,MultiByteToWideChar(
			CP_ACP,
			0,
			lpszPath,
			_countof(lpszPath),
			szMAPIPath,
			cchMAPIPath));
#endif
	}

	delete[] szComponentID;
	delete[] szOfficeLCID;
	delete[] szAppLCID;
} // GetMAPIPath

// Declaration missing from MAPI headers
_Check_return_ STDAPI OpenStreamOnFileW(_In_ LPALLOCATEBUFFER lpAllocateBuffer,
										_In_ LPFREEBUFFER lpFreeBuffer,
										ULONG ulFlags,
										_In_z_ LPCWSTR lpszFileName,
										_In_opt_z_ LPCWSTR lpszPrefix,
										_Out_ LPSTREAM FAR * lppStream);

// Since I never use lpszPrefix, I don't convert it
// To make certain of that, I pass NULL for it
// If I ever do need this param, I'll have to fix this
_Check_return_ STDMETHODIMP MyOpenStreamOnFile(_In_ LPALLOCATEBUFFER lpAllocateBuffer,
											   _In_ LPFREEBUFFER lpFreeBuffer,
											   ULONG ulFlags,
											   _In_z_ LPCWSTR lpszFileName,
											   _In_opt_z_ LPCWSTR /*lpszPrefix*/,
											   _Out_ LPSTREAM FAR * lppStream)
{
	HRESULT hRes = S_OK;

	hRes = OpenStreamOnFileW(

		lpAllocateBuffer,
		lpFreeBuffer,
		ulFlags,
		lpszFileName,
		NULL,
		lppStream);
	if (MAPI_E_CALL_FAILED == hRes)
	{
		// Convert new file name to Ansi
		LPSTR lpAnsiCharStr = NULL;
		EC_H(UnicodeToAnsi(
			lpszFileName,
			&lpAnsiCharStr));
		if (SUCCEEDED(hRes))
		{
			hRes = OpenStreamOnFile(
				lpAllocateBuffer,
				lpFreeBuffer,
				ulFlags,
				(LPTSTR) lpAnsiCharStr,
				NULL,
				lppStream);
		}
		delete[] lpAnsiCharStr;
	}
	return hRes;
} // MyOpenStreamOnFile

_Check_return_ HRESULT HrDupPropset(
					 int cprop,
					 _In_count_(cprop) LPSPropValue rgprop,
					 _In_ LPVOID lpObject,
					 _In_ LPSPropValue*	prgprop)
{
	ULONG cb = NULL;
	HRESULT hRes = S_OK;

	//	Find out how much memory we need
	EC_H(ScCountProps(cprop, rgprop, &cb));

	if (SUCCEEDED(hRes) && cb)
	{
		//	Obtain memory
		if (lpObject != NULL)
		{
			EC_H(MAPIAllocateMore(cb, lpObject, (LPVOID*)prgprop));
		}
		else
		{
			EC_H(MAPIAllocateBuffer(cb, (LPVOID*)prgprop));
		}

		if (SUCCEEDED(hRes) && prgprop)
		{
			//	Copy the properties
			EC_H(ScCopyProps(cprop, rgprop, *prgprop, &cb));
		}
	}

	return hRes;
} // HrDupPropset

typedef ACTIONS* LPACTIONS;

// swiped from EDK rules sample
_Check_return_ STDAPI HrCopyActions(
					 _In_ LPACTIONS lpActsSrc, // source action ptr
					 _In_ LPVOID lpObject, // ptr to existing MAPI buffer
					 _In_ LPACTIONS*	lppActsDst) // ptr to destination ACTIONS buffer
{
	if (!lpActsSrc || !lppActsDst) return MAPI_E_INVALID_PARAMETER;
	if (lpActsSrc->cActions <= 0 || lpActsSrc->lpAction == NULL) return MAPI_E_INVALID_PARAMETER;

	bool fNullObject = (lpObject == NULL);
	HRESULT hRes = S_OK;
	ULONG i = 0;
	LPACTION lpActDst = NULL;
	LPACTION lpActSrc = NULL;
	LPACTIONS lpActsDst = NULL;

	*lppActsDst = NULL;

	if (lpObject != NULL)
	{
		WC_H(MAPIAllocateMore(sizeof(ACTIONS), lpObject, (LPVOID*)lppActsDst));
	}
	else
	{
		WC_H(MAPIAllocateBuffer(sizeof(ACTIONS), (LPVOID*)lppActsDst));
		lpObject = *lppActsDst;
	}

	if (FAILED(hRes)) return hRes;
	// no short circuit returns after here

	lpActsDst = *lppActsDst;
	*lpActsDst = *lpActsSrc;
	lpActsDst->lpAction = NULL;

	WC_H(MAPIAllocateMore(sizeof(ACTION) * lpActsDst->cActions,
		lpObject,
		(LPVOID*)&(lpActsDst->lpAction)));
	if (SUCCEEDED(hRes) && lpActsDst->lpAction)
	{
		// Initialize acttype values for all members of the array to a value
		// that will not cause deallocation errors should the copy fail.
		for (i = 0; i < lpActsDst->cActions; i++)
			lpActsDst->lpAction[i].acttype = OP_BOUNCE;

		// Now actually copy all the members of the array.
		for (i = 0; i < lpActsDst->cActions; i++)
		{
			lpActDst = &(lpActsDst->lpAction[i]);
			lpActSrc = &(lpActsSrc->lpAction[i]);

			*lpActDst = *lpActSrc;

			switch (lpActSrc->acttype)
			{
			case OP_MOVE:			// actMoveCopy
			case OP_COPY:
				if (lpActDst->actMoveCopy.cbStoreEntryId &&
					lpActDst->actMoveCopy.lpStoreEntryId)
				{
					WC_H(MAPIAllocateMore(lpActDst->actMoveCopy.cbStoreEntryId,
						lpObject,
						(LPVOID*)
						&(lpActDst->actMoveCopy.lpStoreEntryId)));
					if (FAILED(hRes)) break;

					memcpy(lpActDst->actMoveCopy.lpStoreEntryId,
						lpActSrc->actMoveCopy.lpStoreEntryId,
						lpActSrc->actMoveCopy.cbStoreEntryId);
				}


				if (lpActDst->actMoveCopy.cbFldEntryId &&
					lpActDst->actMoveCopy.lpFldEntryId)
				{
					WC_H(MAPIAllocateMore(lpActDst->actMoveCopy.cbFldEntryId,
						lpObject,
						(LPVOID*)
						&(lpActDst->actMoveCopy.lpFldEntryId)));
					if (FAILED(hRes)) break;

					memcpy(lpActDst->actMoveCopy.lpFldEntryId,
						lpActSrc->actMoveCopy.lpFldEntryId,
						lpActSrc->actMoveCopy.cbFldEntryId);
				}

				break;

			case OP_REPLY:			// actReply
			case OP_OOF_REPLY:
				if (lpActDst->actReply.cbEntryId &&
					lpActDst->actReply.lpEntryId)
				{
					WC_H(MAPIAllocateMore(lpActDst->actReply.cbEntryId,
						lpObject,
						(LPVOID*)
						&(lpActDst->actReply.lpEntryId)));
					if (FAILED(hRes)) break;

					memcpy(lpActDst->actReply.lpEntryId,
						lpActSrc->actReply.lpEntryId,
						lpActSrc->actReply.cbEntryId);
				}
				break;

			case OP_DEFER_ACTION:	// actDeferAction
				if (lpActSrc->actDeferAction.pbData &&
					lpActSrc->actDeferAction.cbData)
				{
					WC_H(MAPIAllocateMore(lpActDst->actDeferAction.cbData,
						lpObject,
						(LPVOID*)
						&(lpActDst->actDeferAction.pbData)));
					if (FAILED(hRes)) break;

					memcpy(lpActDst->actDeferAction.pbData,
						lpActSrc->actDeferAction.pbData,
						lpActDst->actDeferAction.cbData);
				}
				break;

			case OP_FORWARD:		// lpadrlist
			case OP_DELEGATE:
				lpActDst->lpadrlist = NULL;

				if (lpActSrc->lpadrlist && lpActSrc->lpadrlist->cEntries)
				{
					ULONG j = 0;

					WC_H(MAPIAllocateMore(CbADRLIST(lpActSrc->lpadrlist),
						lpObject,
						(LPVOID*)&(lpActDst->lpadrlist)));
					if (FAILED(hRes)) break;

					lpActDst->lpadrlist->cEntries = lpActSrc->lpadrlist->cEntries;

					// Initialize the new ADRENTRYs and validate cValues.
					for (j = 0; j < lpActSrc->lpadrlist->cEntries; j++)
					{
						lpActDst->lpadrlist->aEntries[j] = lpActSrc->lpadrlist->aEntries[j];
						lpActDst->lpadrlist->aEntries[j].rgPropVals = NULL;

						if (lpActDst->lpadrlist->aEntries[j].cValues == 0)
						{
							hRes = MAPI_E_INVALID_PARAMETER;
							break;
						}
					}

					// Copy the rgPropVals.
					for (j = 0; j < lpActSrc->lpadrlist->cEntries; j++)
					{
						WC_H(HrDupPropset(
							lpActDst->lpadrlist->aEntries[j].cValues,
							lpActSrc->lpadrlist->aEntries[j].rgPropVals,
							lpObject,
							&lpActDst->lpadrlist->aEntries[j].rgPropVals));
						if (FAILED(hRes)) break;
					}
				}

				break;

			case OP_TAG:			// propTag
				WC_H(MyPropCopyMore(
					&lpActDst->propTag,
					&lpActSrc->propTag,
					MAPIAllocateMore,
					lpObject));
				if (FAILED(hRes)) break;

				break;

			case OP_BOUNCE:			// scBounceCode
			case OP_DELETE:			// union not used
			case OP_MARK_AS_READ:
				break; // Nothing to do!

			default:				// error!
				{
					hRes = MAPI_E_INVALID_PARAMETER;
					break;
				}
			}
		}
	}

	if (FAILED(hRes))
	{
		if (fNullObject)
			MAPIFreeBuffer(*lppActsDst);
	}

	return hRes;
} // HrCopyActions

_Check_return_ HRESULT HrCopyRestrictionArray(
							   _In_ LPSRestriction lpResSrc, // source restriction
							   _In_ LPVOID lpObject, // ptr to existing MAPI buffer
							   ULONG cRes, // # elements in array
							   _In_count_(cRes) LPSRestriction lpResDest // destination restriction
							   )
{
	if (!lpResSrc || !lpResDest || !lpObject) return MAPI_E_INVALID_PARAMETER;
	HRESULT hRes = S_OK;
	ULONG i = 0;

	for (i = 0; i < cRes; i++)
	{
		// Copy all the members over
		lpResDest[i] = lpResSrc[i];

		// Now fix up the pointers
		switch (lpResSrc[i].rt)
		{
			// Structures for these two types are identical
		case RES_AND:
		case RES_OR:
			if (lpResSrc[i].res.resAnd.cRes && lpResSrc[i].res.resAnd.lpRes)
			{
				if (lpResSrc[i].res.resAnd.cRes > ULONG_MAX/sizeof(SRestriction))
				{
					hRes = MAPI_E_CALL_FAILED;
					break;
				}
				WC_H(MAPIAllocateMore(sizeof(SRestriction) * lpResSrc[i].res.resAnd.cRes,
					lpObject,
					(LPVOID*)
					&lpResDest[i].res.resAnd.lpRes));
				if (FAILED(hRes)) break;

				WC_H(HrCopyRestrictionArray(
					lpResSrc[i].res.resAnd.lpRes,
					lpObject,
					lpResSrc[i].res.resAnd.cRes,
					lpResDest[i].res.resAnd.lpRes));
				if (FAILED(hRes)) break;
			}
			break;

			// Structures for these two types are identical
		case RES_NOT:
		case RES_COUNT:
			if (lpResSrc[i].res.resNot.lpRes)
			{
				WC_H(MAPIAllocateMore(sizeof(SRestriction),
					lpObject,
					(LPVOID*)
					&lpResDest[i].res.resNot.lpRes));
				if (FAILED(hRes)) break;

				WC_H(HrCopyRestrictionArray(
					lpResSrc[i].res.resNot.lpRes,
					lpObject,
					1,
					lpResDest[i].res.resNot.lpRes));
				if (FAILED(hRes)) break;
			}
			break;

			// Structures for these two types are identical
		case RES_CONTENT:
		case RES_PROPERTY:
			if (lpResSrc[i].res.resContent.lpProp)
			{
				WC_H(HrDupPropset(
					1,
					lpResSrc[i].res.resContent.lpProp,
					lpObject,
					&lpResDest[i].res.resContent.lpProp));
				if (FAILED(hRes)) break;
			}
			break;

		case RES_COMPAREPROPS:
		case RES_BITMASK:
		case RES_SIZE:
		case RES_EXIST:
			break; // Nothing to do.

		case RES_SUBRESTRICTION:
			if (lpResSrc[i].res.resSub.lpRes)
			{
				WC_H(MAPIAllocateMore(sizeof(SRestriction),
					lpObject,
					(LPVOID*)
					&lpResDest[i].res.resSub.lpRes));
				if (FAILED(hRes)) break;

				WC_H(HrCopyRestrictionArray(
					lpResSrc[i].res.resSub.lpRes,
					lpObject,
					1,
					lpResDest[i].res.resSub.lpRes));
				if (FAILED(hRes)) break;
			}
			break;

			// Structures for these two types are identical
		case RES_COMMENT:
		case RES_ANNOTATION:
			if (lpResSrc[i].res.resComment.lpRes)
			{
				WC_H(MAPIAllocateMore(sizeof(SRestriction),
					lpObject,
					(LPVOID*)
					&lpResDest[i].res.resComment.lpRes));
				if (FAILED(hRes)) break;

				WC_H(HrCopyRestrictionArray(
					lpResSrc[i].res.resComment.lpRes,
					lpObject,
					1,
					lpResDest[i].res.resComment.lpRes));
				if (FAILED(hRes)) break;
			}

			if (lpResSrc[i].res.resComment.cValues && lpResSrc[i].res.resComment.lpProp)
			{
				WC_H(HrDupPropset(
					lpResSrc[i].res.resComment.cValues,
					lpResSrc[i].res.resComment.lpProp,
					lpObject,
					&lpResDest[i].res.resComment.lpProp));
				if (FAILED(hRes)) break;
			}
			break;

		default:
			hRes = MAPI_E_INVALID_PARAMETER;
			break;
		}
	}

	return hRes;
} // HrCopyRestrictionArray

_Check_return_ STDAPI HrCopyRestriction(
						 _In_ LPSRestriction lpResSrc, // source restriction ptr
						 _In_opt_ LPVOID lpObject, // ptr to existing MAPI buffer
						 _In_ LPSRestriction* lppResDest // dest restriction buffer ptr
						 )
{
	if (!lppResDest) return MAPI_E_INVALID_PARAMETER;
	*lppResDest = NULL;
	if (!lpResSrc) return S_OK;

	bool fNullObject =(lpObject == NULL);
	HRESULT	hRes = S_OK;

	if (lpObject != NULL)
	{
		WC_H(MAPIAllocateMore(sizeof(SRestriction),
			lpObject,
			(LPVOID*)lppResDest));
	}
	else
	{
		WC_H(MAPIAllocateBuffer(sizeof(SRestriction),
			(LPVOID*)lppResDest));
		lpObject = *lppResDest;
	}
	if (FAILED(hRes)) return hRes;
	// no short circuit returns after here

	WC_H(HrCopyRestrictionArray(
		lpResSrc,
		lpObject,
		1,
		*lppResDest));

	if (FAILED(hRes))
	{
		if (fNullObject)
			MAPIFreeBuffer(*lppResDest);
	}

	return hRes;
} // HrCopyRestriction

// This augmented PropCopyMore is implicitly tied to the built-in MAPIAllocateMore and MAPIAllocateBuffer through
// the calls to HrCopyRestriction and HrCopyActions. Rewriting those functions to accept function pointers is
// expensive for no benefit here. So if you borrow this code, be careful if you plan on using other allocators.
_Check_return_ STDAPI_(SCODE) MyPropCopyMore(_In_ LPSPropValue lpSPropValueDest,
											 _In_ LPSPropValue lpSPropValueSrc,
											 _In_ ALLOCATEMORE * lpfAllocMore,
											 _In_ LPVOID lpvObject)
{
	HRESULT hRes = S_OK;
	switch ( PROP_TYPE(lpSPropValueSrc->ulPropTag) )
	{
	case PT_SRESTRICTION:
	case PT_ACTIONS:
		{
			// It's an action or restriction - we know how to copy those:
			memcpy( (BYTE *) lpSPropValueDest,
				(BYTE *) lpSPropValueSrc,
				sizeof(SPropValue));
			if (PT_SRESTRICTION == PROP_TYPE(lpSPropValueSrc->ulPropTag))
			{
				LPSRestriction lpNewRes = NULL;
				WC_H(HrCopyRestriction(
					(LPSRestriction) lpSPropValueSrc->Value.lpszA,
					lpvObject,
					&lpNewRes));
				lpSPropValueDest->Value.lpszA = (LPSTR) lpNewRes;
			}
			else
			{
				ACTIONS* lpNewAct = NULL;
				WC_H(HrCopyActions(
					(ACTIONS*) lpSPropValueSrc->Value.lpszA,
					lpvObject,
					&lpNewAct));
				lpSPropValueDest->Value.lpszA = (LPSTR) lpNewAct;
			}
			break;
		}
	default:
		hRes = PropCopyMore(lpSPropValueDest,lpSPropValueSrc,lpfAllocMore,lpvObject);
	}
	return hRes;
} // MyPropCopyMore

void WINAPI MyHeapSetInformation(_In_opt_ HANDLE HeapHandle,
								 _In_ HEAP_INFORMATION_CLASS HeapInformationClass,
								 _In_opt_count_(HeapInformationLength) PVOID HeapInformation,
								 _In_ SIZE_T HeapInformationLength)
{
	if (!pfnHeapSetInformation)
	{
		LoadProc(_T("kernel32.dll"), &hModKernel32, "HeapSetInformation", (FARPROC*) &pfnHeapSetInformation); // STRING_OK;
	}

	if (pfnHeapSetInformation) pfnHeapSetInformation(HeapHandle,HeapInformationClass,HeapInformation,HeapInformationLength);
} // MyHeapSetInformation

HRESULT WINAPI MyMimeOleGetCodePageCharset(
	CODEPAGEID cpiCodePage,
	CHARSETTYPE ctCsetType,
	LPHCHARSET phCharset)
{
	if (!pfnMimeOleGetCodePageCharset)
	{
		LoadProc(_T("inetcomm.dll"), &hModInetComm, "MimeOleGetCodePageCharset", (FARPROC*) &pfnMimeOleGetCodePageCharset); // STRING_OK;
	}

	if (pfnMimeOleGetCodePageCharset) return pfnMimeOleGetCodePageCharset(cpiCodePage,ctCsetType,phCharset);
	return MAPI_E_CALL_FAILED;
} // MyMimeOleGetCodePageCharset

STDAPI_(UINT) MsiProvideComponentW(
								 LPCWSTR szProduct,
								 LPCWSTR szFeature,
								 LPCWSTR szComponent,
								 DWORD dwInstallMode,
								 LPWSTR lpPathBuf,
								 LPDWORD pcchPathBuf)
{
	if (!pfnMsiProvideComponentW)
	{
		LoadProc(_T("msi.dll"), &hModMSI, "MimeOleGetCodePageCharset", (FARPROC*) &pfnMsiProvideComponentW); // STRING_OK;
	}

	if (pfnMsiProvideComponentW) return pfnMsiProvideComponentW(szProduct,szFeature,szComponent,dwInstallMode,lpPathBuf,pcchPathBuf);
	return ERROR_NOT_SUPPORTED;
} // MsiProvideComponentW

STDAPI_(UINT) MsiProvideQualifiedComponentW(
	LPCWSTR szCategory,
	LPCWSTR szQualifier,
	DWORD dwInstallMode,
	LPWSTR lpPathBuf,
	LPDWORD pcchPathBuf)
{
	if (!pfnMsiProvideQualifiedComponentW)
	{
		LoadProc(_T("msi.dll"), &hModMSI, "MsiProvideQualifiedComponentW", (FARPROC*) &pfnMsiProvideQualifiedComponentW); // STRING_OK;
	}

	if (pfnMsiProvideQualifiedComponentW) return pfnMsiProvideQualifiedComponentW(szCategory,szQualifier,dwInstallMode,lpPathBuf,pcchPathBuf);
	return ERROR_NOT_SUPPORTED;
} // MsiProvideQualifiedComponentW

BOOL WINAPI MyGetModuleHandleExW(
	DWORD dwFlags,
	LPCWSTR lpModuleName,
	HMODULE* phModule)
{
	if (!pfnGetModuleHandleExW)
	{
		LoadProc(_T("kernel32.dll"), &hModMSI, "GetModuleHandleExW", (FARPROC*) &pfnGetModuleHandleExW); // STRING_OK;
	}

	if (pfnGetModuleHandleExW) return pfnGetModuleHandleExW(dwFlags,lpModuleName,phModule);
	*phModule = GetModuleHandleW(lpModuleName);
	return (*phModule != NULL);
} // MyGetModuleHandleExW