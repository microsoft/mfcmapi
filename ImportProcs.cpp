#include "stdafx.h"
#include "ImportProcs.h"
#include "MAPIFunctions.h"
#include <mapival.h>

_Check_return_ BOOL GetComponentPath(
					  _In_z_ LPSTR szComponent,
					  _In_opt_z_ LPSTR szQualifier,
					  _Inout_z_count_(cchDLLPath) TCHAR* szDllPath,
					  DWORD cchDLLPath);

_Check_return_ HKEY GetMailKey(_In_opt_z_ LPCTSTR szClient);
void GetMapiMsiIds(_In_opt_z_ LPCTSTR szClient, _Deref_out_opt_z_ LPSTR* lpszComponentID, _Deref_out_opt_z_ LPSTR* lpszAppLCID, _Deref_out_opt_z_ LPSTR* lpszOfficeLCID);

// Note: when the user loads MAPI manually, hModMSMAPI and hModMAPI will be the same
HMODULE	hModMSMAPI = NULL; // Address of Outlook's MAPI
HMODULE	hModMAPI = NULL; // Address of MAPI32 in System32
HMODULE	hModMAPIStub = NULL; // Address of the MAPI stub library

HMODULE	hModAclui = NULL;
HMODULE hModRichEd20 = NULL;
HMODULE hModOle32 = NULL;
HMODULE hModUxTheme = NULL;
HMODULE hModInetComm = NULL;
HMODULE hModMSI = NULL;
HMODULE hModKernel32 = NULL;

// For GetComponentPath
typedef BOOL (STDAPICALLTYPE FGETCOMPONENTPATH)(
	LPSTR szComponent,
	LPSTR szQualifier,
	LPSTR szDllPath,
	DWORD cchBufferSize,
	BOOL fInstall);
typedef FGETCOMPONENTPATH FAR * LPFGETCOMPONENTPATH;

// For all the MAPI funcs we use
typedef SCODE (STDMETHODCALLTYPE HRGETONEPROP)(
	LPMAPIPROP lpMapiProp,
	ULONG ulPropTag,
	LPSPropValue FAR *lppProp);
typedef HRGETONEPROP *LPHRGETONEPROP;

typedef void (STDMETHODCALLTYPE FREEPROWS)(
	LPSRowSet lpRows);
typedef FREEPROWS	*LPFREEPROWS;

typedef	LPSPropValue (STDMETHODCALLTYPE PPROPFINDPROP)(
	LPSPropValue lpPropArray, ULONG cValues,
	ULONG ulPropTag);
typedef PPROPFINDPROP *LPPPROPFINDPROP;

typedef SCODE (STDMETHODCALLTYPE SCDUPPROPSET)(
	int cValues, LPSPropValue lpPropArray,
	LPALLOCATEBUFFER lpAllocateBuffer, LPSPropValue FAR *lppPropArray);
typedef SCDUPPROPSET *LPSCDUPPROPSET;

typedef SCODE (STDMETHODCALLTYPE SCCOUNTPROPS)(
	int cValues, LPSPropValue lpPropArray, ULONG FAR *lpcb);
typedef SCCOUNTPROPS *LPSCCOUNTPROPS;

typedef SCODE (STDMETHODCALLTYPE SCCOPYPROPS)(
	int cValues, LPSPropValue lpPropArray, LPVOID lpvDst,
	ULONG FAR *lpcb);
typedef SCCOPYPROPS *LPSCCOPYPROPS;

typedef SCODE (STDMETHODCALLTYPE OPENIMSGONISTG)(
	LPMSGSESS		lpMsgSess,
	LPALLOCATEBUFFER lpAllocateBuffer,
	LPALLOCATEMORE	lpAllocateMore,
	LPFREEBUFFER	lpFreeBuffer,
	LPMALLOC		lpMalloc,
	LPVOID			lpMapiSup,
	LPSTORAGE		lpStg,
	MSGCALLRELEASE FAR *lpfMsgCallRelease,
	ULONG			ulCallerData,
	ULONG			ulFlags,
	LPMESSAGE		FAR *lppMsg );
typedef OPENIMSGONISTG *LPOPENIMSGONISTG;

typedef	LPMALLOC (STDMETHODCALLTYPE MAPIGETDEFAULTMALLOC)(
	void);
typedef MAPIGETDEFAULTMALLOC *LPMAPIGETDEFAULTMALLOC;

typedef	void (STDMETHODCALLTYPE CLOSEIMSGSESSION)(
	LPMSGSESS lpMsgSess);
typedef CLOSEIMSGSESSION *LPCLOSEIMSGSESSION;

typedef	SCODE (STDMETHODCALLTYPE OPENIMSGSESSION)(
	LPMALLOC		lpMalloc,
	ULONG			ulFlags,
	LPMSGSESS FAR	*lppMsgSess);
typedef OPENIMSGSESSION *LPOPENIMSGSESSION;

typedef SCODE (STDMETHODCALLTYPE HRQUERYALLROWS)(
	LPMAPITABLE lpTable,
	LPSPropTagArray lpPropTags,
	LPSRestriction lpRestriction,
	LPSSortOrderSet lpSortOrderSet,
	LONG crowsMax,
	LPSRowSet FAR *lppRows);
typedef HRQUERYALLROWS *LPHRQUERYALLROWS;

typedef SCODE (STDMETHODCALLTYPE MAPIOPENFORMMGR)(
	LPMAPISESSION pSession, LPMAPIFORMMGR FAR * ppmgr);
typedef MAPIOPENFORMMGR *LPMAPIOPENFORMMGR;

typedef HRESULT (STDMETHODCALLTYPE RTFSYNC) (
	LPMESSAGE lpMessage, ULONG ulFlags, BOOL FAR * lpfMessageUpdated);
typedef RTFSYNC *LPRTFSYNC;

typedef SCODE (STDMETHODCALLTYPE HRSETONEPROP)(
	LPMAPIPROP lpMapiProp,
	LPSPropValue lpProp);
typedef HRSETONEPROP *LPHRSETONEPROP;

typedef void (STDMETHODCALLTYPE FREEPADRLIST)(
	LPADRLIST lpAdrList);
typedef FREEPADRLIST *LPFREEPADRLIST;

typedef SCODE (STDMETHODCALLTYPE PROPCOPYMORE)(
	LPSPropValue		lpSPropValueDest,
	LPSPropValue		lpSPropValueSrc,
	ALLOCATEMORE *	lpfAllocMore,
	LPVOID			lpvObject );
typedef PROPCOPYMORE *LPPROPCOPYMORE;

typedef HRESULT (STDMETHODCALLTYPE WRAPCOMPRESSEDRTFSTREAM) (
	LPSTREAM lpCompressedRTFStream, ULONG ulFlags, LPSTREAM FAR * lpUncompressedRTFStream);
typedef WRAPCOMPRESSEDRTFSTREAM *LPWRAPCOMPRESSEDRTFSTREAM;

typedef SCODE (STDMETHODCALLTYPE HRVALIDATEIPMSUBTREE)(
	LPMDB lpMDB, ULONG ulFlags,
	ULONG FAR *lpcValues, LPSPropValue FAR *lppValues,
	LPMAPIERROR FAR *lpperr);
typedef HRVALIDATEIPMSUBTREE *LPHRVALIDATEIPMSUBTREE;

typedef HRESULT (STDAPICALLTYPE MAPIOPENLOCALFORMCONTAINER)(LPMAPIFORMCONTAINER FAR * ppfcnt);
typedef MAPIOPENLOCALFORMCONTAINER *LPMAPIOPENLOCALFORMCONTAINER;

typedef SCODE (STDAPICALLTYPE HRDISPATCHNOTIFICATIONS)(ULONG ulFlags);
typedef HRDISPATCHNOTIFICATIONS *LPHRDISPATCHNOTIFICATIONS;

typedef HRESULT (STDMETHODCALLTYPE WRAPSTOREENTRYID)(
	ULONG ulFlags,
	__in LPTSTR lpszDLLName,
	ULONG cbOrigEntry,
	LPENTRYID lpOrigEntry,
	ULONG *lpcbWrappedEntry,
	LPENTRYID *lppWrappedEntry);
typedef WRAPSTOREENTRYID *LPWRAPSTOREENTRYID;

typedef SCODE (STDMETHODCALLTYPE CREATEIPROP)(
	LPCIID					lpInterface,
	ALLOCATEBUFFER FAR *	lpAllocateBuffer,
	ALLOCATEMORE FAR *		lpAllocateMore,
	FREEBUFFER FAR *		lpFreeBuffer,
	LPVOID					lpvReserved,
	LPPROPDATA FAR *		lppPropData);
typedef CREATEIPROP *LPCREATEIPROP;

typedef SCODE (STDMETHODCALLTYPE CREATETABLE)(
	LPCIID					lpInterface,
	ALLOCATEBUFFER FAR *	lpAllocateBuffer,
	ALLOCATEMORE FAR *		lpAllocateMore,
	FREEBUFFER FAR *		lpFreeBuffer,
	LPVOID					lpvReserved,
	ULONG					ulTableType,
	ULONG					ulPropTagIndexColumn,
	LPSPropTagArray		lpSPropTagArrayColumns,
	LPTABLEDATA FAR *		lppTableData);
typedef CREATETABLE *LPCREATETABLE;

typedef HRESULT (STDMETHODCALLTYPE HRVALIDATEPARAMETERS)(
	METHODS eMethod,
	LPVOID FAR *ppFirstArg);
typedef HRVALIDATEPARAMETERS *LPHRVALIDATEPARAMETERS;

typedef HRESULT (STDMETHODCALLTYPE BUILDDISPLAYTABLE)(
	LPALLOCATEBUFFER	lpAllocateBuffer,
	LPALLOCATEMORE		lpAllocateMore,
	LPFREEBUFFER		lpFreeBuffer,
	LPMALLOC			lpMalloc,
	HINSTANCE			hInstance,
	UINT				cPages,
	LPDTPAGE			lpPage,
	ULONG				ulFlags,
	LPMAPITABLE *		lppTable,
	LPTABLEDATA	*		lppTblData);
typedef BUILDDISPLAYTABLE *LPBUILDDISPLAYTABLE;

typedef int (STDMETHODCALLTYPE MNLS_LSTRLENW)(
	LPCWSTR	lpString);
typedef MNLS_LSTRLENW *LPMNLS_LSTRLENW;

typedef HRESULT (STDMETHODCALLTYPE FAR * LPOPENSTREAMONFILEW) (
	LPALLOCATEBUFFER	lpAllocateBuffer,
	LPFREEBUFFER		lpFreeBuffer,
	ULONG				ulFlags,
	__in LPCWSTR		lpszFileName,
	__in_opt LPCWSTR	lpszPrefix,
	LPSTREAM FAR *		lppStream);

typedef HTHEME (STDMETHODCALLTYPE CLOSETHEMEDATA)
(
 HTHEME hTheme);
typedef CLOSETHEMEDATA FAR * LPCLOSETHEMEDATA;

typedef BOOL (WINAPI HEAPSETINFORMATION) (
    HANDLE HeapHandle,
    HEAP_INFORMATION_CLASS HeapInformationClass,
    PVOID HeapInformation,
    SIZE_T HeapInformationLength);
typedef HEAPSETINFORMATION FAR * LPHEAPSETINFORMATION;

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

// From kernel32.dll
LPHEAPSETINFORMATION pfnHeapSetInformation = NULL;

// All of these get loaded from a MAPI DLL:
LPFGETCOMPONENTPATH			pfnFGetComponentPath = NULL;
LPLAUNCHWIZARDENTRY			pfnLaunchWizard = NULL;
LPMAPIALLOCATEBUFFER		pfnMAPIAllocateBuffer = NULL;
LPMAPIALLOCATEMORE			pfnMAPIAllocateMore = NULL;
LPMAPIFREEBUFFER			pfnMAPIFreeBuffer = NULL;
LPHRGETONEPROP				pfnHrGetOneProp = NULL;
LPMAPIINITIALIZE			pfnMAPIInitialize = NULL;
LPMAPIUNINITIALIZE			pfnMAPIUninitialize = NULL;
LPFREEPROWS					pfnLPFreeProws = NULL;
LPPPROPFINDPROP				pfnPpropFindProp = NULL;
LPSCDUPPROPSET				pfnScDupPropset = NULL;
LPSCCOUNTPROPS				pfnScCountProps = NULL;
LPSCCOPYPROPS				pfnScCopyProps = NULL;
LPOPENIMSGONISTG			pfnOpenIMsgOnIStg = NULL;
LPMAPIGETDEFAULTMALLOC		pfnMAPIGetDefaultMalloc = NULL;
LPOPENTNEFSTREAMEX			pfnOpenTnefStreamEx = NULL;
LPOPENSTREAMONFILE			pfnOpenStreamOnFile = NULL;
LPOPENSTREAMONFILEW			pfnOpenStreamOnFileW = NULL;
LPCLOSEIMSGSESSION			pfnCloseIMsgSession = NULL;
LPOPENIMSGSESSION			pfnOpenIMsgSession = NULL;
LPHRQUERYALLROWS			pfnHrQueryAllRows = NULL;
LPMAPIOPENFORMMGR			pfnMAPIOpenFormMgr = NULL;
LPRTFSYNC					pfnRTFSync = NULL;
LPHRSETONEPROP				pfnHrSetOneProp = NULL;
LPFREEPADRLIST				pfnFreePadrlist = NULL;
LPPROPCOPYMORE				pfnPropCopyMore = NULL;
LPWRAPCOMPRESSEDRTFSTREAM	pfnWrapCompressedRTFStream = NULL;
LPWRAPSTOREENTRYID			pfnWrapStoreEntryID = NULL;
LPCREATEIPROP				pfnCreateIProp = NULL;
LPCREATETABLE				pfnCreateTable = NULL;
LPHRVALIDATEPARAMETERS		pfnHrValidateParameters = NULL;
LPBUILDDISPLAYTABLE			pfnBuildDisplayTable = NULL;
LPMNLS_LSTRLENW				pfnMNLS_lstrlenW = NULL;
LPMAPILOGONEX				pfnMAPILogonEx = NULL;
LPMAPIADMINPROFILES			pfnMAPIAdminProfiles = NULL;
LPHRVALIDATEIPMSUBTREE		pfnHrValidateIPMSubtree = NULL;
LPWRAPCOMPRESSEDRTFSTREAMEX	pfnWrapEx = NULL;
LPMAPIOPENLOCALFORMCONTAINER pfnMAPIOpenLocalFormContainer = NULL;
LPHRDISPATCHNOTIFICATIONS	pfnHrDispatchNotifications = NULL;

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

void LoadGetComponentPath()
{
	if (pfnFGetComponentPath) return;

	// Try to get FGetComponentPath from mapi32.dll
	LoadProc(_T("mapi32.dll"), &hModMAPI, "FGetComponentPath", (FARPROC*) &pfnFGetComponentPath); // STRING_OK
	// If we don't find it there, try mapistub.dll
	LoadProc(_T("mapistub.dll"), &hModMAPIStub, "FGetComponentPath", (FARPROC*) &pfnFGetComponentPath); // STRING_OK

	DebugPrint(DBGLoadLibrary,_T("FGetComponentPath loaded at %p\n"),pfnFGetComponentPath);
} // LoadGetComponentPath

_Check_return_ BOOL GetComponentPath(
					  _In_z_ LPSTR szComponent,
					  _In_opt_z_ LPSTR szQualifier,
					  _Inout_z_count_(cchDLLPath) TCHAR* szDllPath,
					  DWORD cchDLLPath)
{
	HRESULT hRes = S_OK;
	LoadGetComponentPath();

	if (!pfnFGetComponentPath) return false;

	DebugPrint(DBGLoadLibrary,_T("GetComponentPath - Called with szComponent = \"%hs\", szQualifier = \"%hs\"\n"),szComponent,szQualifier);
#ifdef UNICODE
	int iRet = 0;
	CHAR szAsciiPath[MAX_PATH] = {0};

	EC_B(pfnFGetComponentPath(
		szComponent,
		szQualifier,
		szAsciiPath,
		_countof(szAsciiPath),
		TRUE));

	// Convert to Unicode.
	EC_D(iRet,MultiByteToWideChar(
		CP_ACP,
		0,
		szAsciiPath,
		_countof(szAsciiPath),
		szDllPath,
		cchDLLPath));
#else
	EC_B(pfnFGetComponentPath(
		szComponent,
		szQualifier,
		szDllPath,
		cchDLLPath,
		TRUE));
#endif
	DebugPrint(DBGLoadLibrary,_T("GetComponentPath - hRes = 0x%08X\n"),hRes);
	if (SUCCEEDED(hRes)) DebugPrint(DBGLoadLibrary,_T("GetComponentPath - szDllPath = \"%s\"\n"),szDllPath);
	return SUCCEEDED(hRes);
} // GetComponentPath

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
	static BOOL		bSystemDirLoaded = false;

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
	LoadProc(_T("inetcomm.dll"), &hModInetComm, "MimeOleGetCodePageCharset", (FARPROC*) &pfnMimeOleGetCodePageCharset); // STRING_OK;
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
	BOOL bClientIsDefault = false;

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
		if (szAppLCID)
		{
			WC_B(GetComponentPath(
				szComponentID,
				szAppLCID,
				szMAPIPath,
				cchMAPIPath));
		}
		if ((FAILED(hRes) || szMAPIPath[0] == _T('\0')) && szOfficeLCID)
		{
			hRes = S_OK;
			WC_B(GetComponentPath(
				szComponentID,
				szOfficeLCID,
				szMAPIPath,
				cchMAPIPath));
		}
		if (FAILED(hRes) || szMAPIPath[0] == _T('\0'))
		{
			hRes = S_OK;
			WC_B(GetComponentPath(
				szComponentID,
				NULL,
				szMAPIPath,
				cchMAPIPath));
		}
	}

	delete[] szComponentID;
	delete[] szOfficeLCID;
	delete[] szAppLCID;
} // GetMAPIPath

void AutoLoadMAPI()
{
	DebugPrint(DBGLoadLibrary,_T("AutoLoadMAPI - loading MAPI exports\n"));

	// Attempt load default MAPI first if we have one
	// This will handle Outlook MAPI
	if (!hModMSMAPI)
	{
		TCHAR szMSMAPI32path[MAX_PATH] = {0};
		GetMAPIPath(NULL,szMSMAPI32path,_countof(szMSMAPI32path));
		if (szMSMAPI32path[0] != NULL)
		{
			hModMSMAPI = MyLoadLibrary(szMSMAPI32path);
		}
	}
	if (hModMSMAPI) LoadMAPIFuncs(hModMSMAPI);

	// In case that fails - load the stub library from system32
	// This will handle Exchange MAPI and rare case where we don't load Outlook MAPI
	if (!hModMAPI) hModMAPI = LoadFromSystemDir(_T("mapi32.dll")); // STRING_OK
	// If hModMAPI is the same as hModMSMAPI, then we don't need to try it again
	if (hModMAPI && hModMAPI != hModMSMAPI) LoadMAPIFuncs(hModMAPI);
} // AutoLoadMAPI

void UnloadMAPI()
{
	DebugPrint(DBGLoadLibrary,_T("UnloadMAPI - unloading MAPI exports\n"));
	// Blank out all the functions we've loaded from MAPI DLLs:
	pfnFGetComponentPath = NULL;
	pfnLaunchWizard = NULL;
	pfnMAPIAllocateBuffer = NULL;
	pfnMAPIAllocateMore = NULL;
	pfnMAPIFreeBuffer = NULL;
	pfnHrGetOneProp = NULL;
	pfnMAPIInitialize = NULL;
	pfnMAPIUninitialize = NULL;
	pfnLPFreeProws = NULL;
	pfnPpropFindProp = NULL;
	pfnScDupPropset = NULL;
	pfnScCountProps = NULL;
	pfnScCopyProps = NULL;
	pfnOpenIMsgOnIStg = NULL;
	pfnMAPIGetDefaultMalloc = NULL;
	pfnOpenTnefStreamEx = NULL;
	pfnOpenStreamOnFile = NULL;
	pfnOpenStreamOnFileW = NULL;
	pfnCloseIMsgSession = NULL;
	pfnOpenIMsgSession = NULL;
	pfnHrQueryAllRows = NULL;
	pfnMAPIOpenFormMgr = NULL;
	pfnRTFSync = NULL;
	pfnHrSetOneProp = NULL;
	pfnFreePadrlist = NULL;
	pfnPropCopyMore = NULL;
	pfnWrapCompressedRTFStream = NULL;
	pfnWrapStoreEntryID = NULL;
	pfnCreateIProp = NULL;
	pfnCreateTable = NULL;
	pfnHrValidateParameters = NULL;
	pfnBuildDisplayTable = NULL;
	pfnMNLS_lstrlenW = NULL;
	pfnMAPILogonEx = NULL;
	pfnMAPIAdminProfiles = NULL;
	pfnHrValidateIPMSubtree = NULL;
	pfnWrapEx = NULL;
	pfnMAPIOpenLocalFormContainer = NULL;
	pfnHrDispatchNotifications = NULL;

	HRESULT hRes = S_OK;
	if (hModMSMAPI) WC_B(FreeLibrary(hModMSMAPI));
	hModMSMAPI = NULL;
	hRes = S_OK;
	if (hModMAPIStub) WC_B(FreeLibrary(hModMAPIStub));
	hModMAPIStub = NULL;
	hRes = S_OK;
	if (hModMAPI) WC_B(FreeLibrary(hModMAPI));
	hModMAPI = NULL;
	hRes = S_OK;
} // UnloadMAPI

struct MAPI_FUNC_ENTRY
{
	FARPROC* fFunc;
	LPSTR szEntryPoint;
};
#define MAKE_MAPI_FUNC_ENTRY(_fFunc, _szEntryPoint) {(FARPROC*)&(_fFunc), (_szEntryPoint)},
MAPI_FUNC_ENTRY g_MAPIFuncs[] = {
	MAKE_MAPI_FUNC_ENTRY(pfnLaunchWizard, LAUNCHWIZARDENTRYNAME)
	MAKE_MAPI_FUNC_ENTRY(pfnMAPIAllocateBuffer, "MAPIAllocateBuffer") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnMAPIAllocateMore, "MAPIAllocateMore") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnMAPIFreeBuffer, "MAPIFreeBuffer") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnHrGetOneProp, "HrGetOneProp@12") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnHrGetOneProp, "HrGetOneProp") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnMAPIInitialize, "MAPIInitialize") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnMAPIUninitialize, "MAPIUninitialize") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnLPFreeProws, "FreeProws@4") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnLPFreeProws, "FreeProws") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnPpropFindProp, "PpropFindProp@12") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnPpropFindProp, "PpropFindProp") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnScDupPropset, "ScDupPropset@16") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnScDupPropset, "ScDupPropset") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnScCountProps, "ScCountProps@12") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnScCountProps, "ScCountProps") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnScCopyProps, "ScCopyProps@16") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnScCopyProps, "ScCopyProps") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnOpenIMsgOnIStg, "OpenIMsgOnIStg@44") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnOpenIMsgOnIStg, "OpenIMsgOnIStg") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnMAPIGetDefaultMalloc, "MAPIGetDefaultMalloc@0") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnMAPIGetDefaultMalloc, "MAPIGetDefaultMalloc") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnOpenTnefStreamEx, "OpenTnefStreamEx") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnOpenStreamOnFile, "OpenStreamOnFile") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnOpenStreamOnFileW, "OpenStreamOnFileW") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnCloseIMsgSession, "CloseIMsgSession@4") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnCloseIMsgSession, "CloseIMsgSession") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnOpenIMsgSession, "OpenIMsgSession@12") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnOpenIMsgSession, "OpenIMsgSession") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnHrQueryAllRows, "HrQueryAllRows@24") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnHrQueryAllRows, "HrQueryAllRows") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnMAPIOpenFormMgr, "MAPIOpenFormMgr") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnRTFSync, "RTFSync") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnHrSetOneProp, "HrSetOneProp@8") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnHrSetOneProp, "HrSetOneProp") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnFreePadrlist, "FreePadrlist@4") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnFreePadrlist, "FreePadrlist") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnPropCopyMore, "PropCopyMore@16") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnPropCopyMore, "PropCopyMore") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnWrapCompressedRTFStream, "WrapCompressedRTFStream") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnWrapStoreEntryID, "WrapStoreEntryID@24") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnWrapStoreEntryID, "WrapStoreEntryID") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnCreateIProp, "CreateIProp@24") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnCreateIProp, "CreateIProp") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnCreateTable, "CreateTable@36") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnCreateTable, "CreateTable") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnHrValidateParameters, "HrValidateParameters@8") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnHrValidateParameters, "HrValidateParameters") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnBuildDisplayTable, "BuildDisplayTable@40") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnBuildDisplayTable, "BuildDisplayTable") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnMNLS_lstrlenW, "MNLS_lstrlenW") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnMAPILogonEx, "MAPILogonEx") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnMAPIAdminProfiles, "MAPIAdminProfiles") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnHrValidateIPMSubtree, "HrValidateIPMSubtree@20") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnHrValidateIPMSubtree, "HrValidateIPMSubtree") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnWrapEx, "WrapCompressedRTFStreamEx") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnMAPIOpenLocalFormContainer, "MAPIOpenLocalFormContainer") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnHrDispatchNotifications, "HrDispatchNotifications@4") // STRING_OK
	MAKE_MAPI_FUNC_ENTRY(pfnHrDispatchNotifications, "HrDispatchNotifications") // STRING_OK
};
ULONG g_cMAPIFuncs = _countof(g_MAPIFuncs);

void LoadMAPIFuncs(HMODULE hMod)
{
	DebugPrint(DBGLoadLibrary,_T("LoadMAPIFuncs - loading from %p\n"),hMod);
	HRESULT hRes = S_OK;
	if (!hMod) return;

	ULONG i = 0;
	for (i = 0 ; i < g_cMAPIFuncs ; i++)
	{
		if (!*g_MAPIFuncs[i].fFunc)
		{
			hRes = S_OK;
			WC_D((*g_MAPIFuncs[i].fFunc), GetProcAddress(hMod, g_MAPIFuncs[i].szEntryPoint));
			if (!*g_MAPIFuncs[i].fFunc) DebugPrint(DBGLoadLibrary,_T("Failed to load \"%hs\" from %p\n"), g_MAPIFuncs[i].szEntryPoint, hMod);
		}
	}
} // LoadMAPIFuncs

#define CHECKLOAD(pfn) \
{ \
	if (!(pfn)) AutoLoadMAPI(); \
	if (!(pfn)) \
	{ \
		DebugPrint(DBGLoadLibrary,_T("Function %hs not loaded\n"),#pfn); \
	} \
} // CHECKLOAD

STDMETHODIMP_(SCODE) MAPIAllocateBuffer(ULONG cbSize,
										LPVOID FAR * lppBuffer)
{
	CHECKLOAD(pfnMAPIAllocateBuffer);
	if (pfnMAPIAllocateBuffer) return pfnMAPIAllocateBuffer(cbSize,lppBuffer);
	return MAPI_E_CALL_FAILED;
} // MAPIAllocateBuffer

STDMETHODIMP_(SCODE) MAPIAllocateMore(ULONG cbSize,
									  LPVOID lpObject,
									  LPVOID FAR * lppBuffer)
{
	CHECKLOAD(pfnMAPIAllocateMore);
	if (pfnMAPIAllocateMore) return pfnMAPIAllocateMore(cbSize,lpObject,lppBuffer);
	return MAPI_E_CALL_FAILED;
} // MAPIAllocateMore

STDAPI_(ULONG) MAPIFreeBuffer(LPVOID lpBuffer)
{
	// If we don't have this, the memory's already leaked - don't compound it with a crash
	if (!lpBuffer) return NULL;
	CHECKLOAD(pfnMAPIFreeBuffer);
	if (pfnMAPIFreeBuffer) return pfnMAPIFreeBuffer(lpBuffer);
	return NULL;
} // MAPIFreeBuffer

STDAPI HrGetOneProp(LPMAPIPROP lpMapiProp,
					ULONG ulPropTag,
					LPSPropValue FAR * lppProp)
{
	CHECKLOAD(pfnHrGetOneProp);
	if (pfnHrGetOneProp) return pfnHrGetOneProp(lpMapiProp,ulPropTag,lppProp);
	return MAPI_E_CALL_FAILED;
} // HrGetOneProp

STDMETHODIMP MAPIInitialize(LPVOID lpMapiInit)
{
	CHECKLOAD(pfnMAPIInitialize);
	if (pfnMAPIInitialize) return pfnMAPIInitialize(lpMapiInit);
	return MAPI_E_CALL_FAILED;
} // MAPIInitialize

STDAPI_(void) MAPIUninitialize(void)
{
	CHECKLOAD(pfnMAPIUninitialize);
	if (pfnMAPIUninitialize) pfnMAPIUninitialize();
} // MAPIUninitialize

STDAPI_(void) FreeProws(LPSRowSet lpRows)
{
	CHECKLOAD(pfnLPFreeProws);
	if (pfnLPFreeProws) pfnLPFreeProws(lpRows);
} // FreeProws

STDAPI_(LPSPropValue) PpropFindProp(LPSPropValue lpPropArray,
									ULONG cValues,
									ULONG ulPropTag)
{
	CHECKLOAD(pfnPpropFindProp);
	if (pfnPpropFindProp) return pfnPpropFindProp(lpPropArray,cValues,ulPropTag);
	return NULL;
} // PpropFindProp

STDAPI_(SCODE) ScDupPropset(int cValues,
							LPSPropValue lpPropArray,
							LPALLOCATEBUFFER lpAllocateBuffer,
							LPSPropValue FAR *lppPropArray)
{
	CHECKLOAD(pfnScDupPropset);
	if (pfnScDupPropset) return pfnScDupPropset(cValues,lpPropArray,lpAllocateBuffer,lppPropArray);
	return MAPI_E_CALL_FAILED;
} // ScDupPropset

STDAPI_(SCODE) ScCountProps(int cValues,
							LPSPropValue lpPropArray,
							ULONG FAR *lpcb)
{
	CHECKLOAD(pfnScCountProps);
	if (pfnScCountProps) return pfnScCountProps(cValues,lpPropArray,lpcb);
	return MAPI_E_CALL_FAILED;
} // ScCountProps

STDAPI_(SCODE) ScCopyProps(int cValues,
						   LPSPropValue lpPropArray,
						   LPVOID lpvDst,
						   ULONG FAR *lpcb)
{
	CHECKLOAD(pfnScCopyProps);
	if (pfnScCopyProps) return pfnScCopyProps(cValues,lpPropArray,lpvDst,lpcb);
	return MAPI_E_CALL_FAILED;
} // ScCopyProps

STDAPI_(SCODE) OpenIMsgOnIStg(LPMSGSESS lpMsgSess,
							  LPALLOCATEBUFFER lpAllocateBuffer,
							  LPALLOCATEMORE lpAllocateMore,
							  LPFREEBUFFER lpFreeBuffer,
							  LPMALLOC lpMalloc,
							  LPVOID lpMapiSup,
							  LPSTORAGE lpStg,
							  MSGCALLRELEASE FAR * lpfMsgCallRelease,
							  ULONG ulCallerData,
							  ULONG ulFlags,
							  LPMESSAGE FAR * lppMsg)
{
	CHECKLOAD(pfnOpenIMsgOnIStg);
	if (pfnOpenIMsgOnIStg) return pfnOpenIMsgOnIStg(
		lpMsgSess,
		lpAllocateBuffer,
		lpAllocateMore,
		lpFreeBuffer,
		lpMalloc,
		lpMapiSup,
		lpStg,
		lpfMsgCallRelease,
		ulCallerData,
		ulFlags,
		lppMsg);
	return MAPI_E_CALL_FAILED;
} // OpenIMsgOnIStg

STDAPI_(LPMALLOC) MAPIGetDefaultMalloc(VOID)
{
	CHECKLOAD(pfnMAPIGetDefaultMalloc);
	if (pfnMAPIGetDefaultMalloc) return pfnMAPIGetDefaultMalloc();
	return NULL;
} // MAPIGetDefaultMalloc

STDMETHODIMP OpenTnefStreamEx(LPVOID lpvSupport,
							  LPSTREAM lpStream,
							  LPTSTR lpszStreamName,
							  ULONG ulFlags,
							  LPMESSAGE lpMessage,
							  WORD wKeyVal,
							  LPADRBOOK lpAdressBook,
							  LPITNEF FAR * lppTNEF)
{
	CHECKLOAD(pfnOpenTnefStreamEx);
	if (pfnOpenTnefStreamEx) return pfnOpenTnefStreamEx(
		lpvSupport,
		lpStream,
		lpszStreamName,
		ulFlags,
		lpMessage,
		wKeyVal,
		lpAdressBook,
		lppTNEF);
	return MAPI_E_CALL_FAILED;
} // OpenTnefStreamEx

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
	CHECKLOAD(pfnOpenStreamOnFile);
	CHECKLOAD(pfnOpenStreamOnFileW);
	HRESULT hRes = S_OK;
	if (!pfnOpenStreamOnFile && !pfnOpenStreamOnFileW) return MAPI_E_CALL_FAILED;
	if (pfnOpenStreamOnFileW)
	{
		hRes = pfnOpenStreamOnFileW(
			lpAllocateBuffer,
			lpFreeBuffer,
			ulFlags,
			lpszFileName,
			NULL,
			lppStream);
	}
	else
	{
		// Convert new file name to Ansi
		LPSTR lpAnsiCharStr = NULL;
		EC_H(UnicodeToAnsi(
			lpszFileName,
			&lpAnsiCharStr));
		if (SUCCEEDED(hRes))
		{
			hRes = pfnOpenStreamOnFile(
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

STDAPI_(void) CloseIMsgSession(LPMSGSESS lpMsgSess)
{
	CHECKLOAD(pfnCloseIMsgSession);
	if (pfnCloseIMsgSession) pfnCloseIMsgSession(
		lpMsgSess);
} // CloseIMsgSession

STDAPI_(SCODE) OpenIMsgSession(LPMALLOC lpMalloc,
							   ULONG ulFlags,
							   LPMSGSESS FAR * lppMsgSess)
{
	CHECKLOAD(pfnOpenIMsgSession);
	if (pfnOpenIMsgSession) return pfnOpenIMsgSession(
		lpMalloc,
		ulFlags,
		lppMsgSess);
	return MAPI_E_CALL_FAILED;
} // OpenIMsgSession

STDAPI HrQueryAllRows(LPMAPITABLE lpTable,
					  LPSPropTagArray lpPropTags,
					  LPSRestriction lpRestriction,
					  LPSSortOrderSet lpSortOrderSet,
					  LONG crowsMax,
					  LPSRowSet FAR *lppRows)
{
	CHECKLOAD(pfnHrQueryAllRows);
	if (pfnHrQueryAllRows) return pfnHrQueryAllRows(
		lpTable,
		lpPropTags,
		lpRestriction,
		lpSortOrderSet,
		crowsMax,
		lppRows);
	return MAPI_E_CALL_FAILED;
} // HrQueryAllRows

STDAPI MAPIOpenFormMgr(LPMAPISESSION pSession, LPMAPIFORMMGR FAR * ppmgr)
{
	CHECKLOAD(pfnMAPIOpenFormMgr);
	if (pfnMAPIOpenFormMgr) return pfnMAPIOpenFormMgr(
		pSession,
		ppmgr);
	return MAPI_E_CALL_FAILED;
} // MAPIOpenFormMgr

STDAPI_(HRESULT) RTFSync(LPMESSAGE lpMessage, ULONG ulFlags, BOOL FAR * lpfMessageUpdated)
{
	CHECKLOAD(pfnRTFSync);
	if (pfnRTFSync) return pfnRTFSync(lpMessage,ulFlags,lpfMessageUpdated);
	return MAPI_E_CALL_FAILED;
} // RTFSync

STDAPI HrSetOneProp(LPMAPIPROP lpMapiProp,
					LPSPropValue lpProp)
{
	CHECKLOAD(pfnHrSetOneProp);
	if (pfnHrSetOneProp) return pfnHrSetOneProp(lpMapiProp,lpProp);
	return MAPI_E_CALL_FAILED;
} // HrSetOneProp

STDAPI_(void) FreePadrlist(LPADRLIST lpAdrlist)
{
	CHECKLOAD(pfnFreePadrlist);
	if (pfnFreePadrlist) pfnFreePadrlist(lpAdrlist);
} // FreePadrlist

_Check_return_ HRESULT HrDupPropset(
					 int cprop,
					 _In_count_(cprop) LPSPropValue rgprop,
					 _In_ LPVOID lpObject,
					 _In_ LPSPropValue FAR *	prgprop)
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
			EC_H(MAPIAllocateMore(cb, lpObject, (LPVOID FAR *)prgprop));
		}
		else
		{
			EC_H(MAPIAllocateBuffer(cb, (LPVOID FAR *)prgprop));
		}

		if (SUCCEEDED(hRes) && prgprop)
		{
			//	Copy the properties
			EC_H(ScCopyProps(cprop, rgprop, *prgprop, &cb));
		}
	}

	return hRes;
} // HrDupPropset

typedef ACTIONS FAR *	LPACTIONS;

// swiped from EDK rules sample
_Check_return_ STDAPI HrCopyActions(
					 _In_ LPACTIONS lpActsSrc, // source action ptr
					 _In_ LPVOID lpObject, // ptr to existing MAPI buffer
					 _In_ LPACTIONS FAR *	lppActsDst) // ptr to destination ACTIONS buffer
{
	if (!lpActsSrc || !lppActsDst) return MAPI_E_INVALID_PARAMETER;
	if (lpActsSrc->cActions <= 0 || lpActsSrc->lpAction == NULL) return MAPI_E_INVALID_PARAMETER;

	BOOL fNullObject = (lpObject == NULL);
	HRESULT hRes = S_OK;
	ULONG i = 0;
	LPACTION lpActDst = NULL;
	LPACTION lpActSrc = NULL;
	LPACTIONS lpActsDst = NULL;

	*lppActsDst = NULL;

	if (lpObject != NULL)
	{
		WC_H(MAPIAllocateMore(sizeof(ACTIONS), lpObject, (LPVOID FAR *)lppActsDst));
	}
	else
	{
		WC_H(MAPIAllocateBuffer(sizeof(ACTIONS), (LPVOID FAR *)lppActsDst));
		lpObject = *lppActsDst;
	}

	if (FAILED(hRes)) return hRes;
	// no short circuit returns after here

	lpActsDst = *lppActsDst;
	*lpActsDst = *lpActsSrc;
	lpActsDst->lpAction = NULL;

	WC_H(MAPIAllocateMore(sizeof(ACTION) * lpActsDst->cActions,
		lpObject,
		(LPVOID FAR *)&(lpActsDst->lpAction)));
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
						(LPVOID FAR *)
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
						(LPVOID FAR *)
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
						(LPVOID FAR *)
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
						(LPVOID FAR *)
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
						(LPVOID FAR *)&(lpActDst->lpadrlist)));
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
				WC_H(PropCopyMore(
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
					(LPVOID FAR *)
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
					(LPVOID FAR *)
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
					(LPVOID FAR *)
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
					(LPVOID FAR *)
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
						 _In_ LPSRestriction FAR * lppResDest // dest restriction buffer ptr
						 )
{
	if (!lppResDest) return MAPI_E_INVALID_PARAMETER;
	*lppResDest = NULL;
	if (!lpResSrc) return S_OK;

	BOOL fNullObject =(lpObject == NULL);
	HRESULT	hRes = S_OK;

	if (lpObject != NULL)
	{
		WC_H(MAPIAllocateMore(sizeof(SRestriction),
			lpObject,
			(LPVOID FAR *)lppResDest));
	}
	else
	{
		WC_H(MAPIAllocateBuffer(sizeof(SRestriction),
			(LPVOID FAR *)lppResDest));
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
STDAPI_(SCODE) PropCopyMore(LPSPropValue lpSPropValueDest,
							LPSPropValue lpSPropValueSrc,
							ALLOCATEMORE * lpfAllocMore,
							LPVOID lpvObject)
{
	CHECKLOAD(pfnPropCopyMore);
	if (!pfnPropCopyMore) return MAPI_E_CALL_FAILED;
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
		hRes = pfnPropCopyMore(lpSPropValueDest,lpSPropValueSrc,lpfAllocMore,lpvObject);
	}
	return hRes;
} // PropCopyMore

STDAPI_(HRESULT) WrapCompressedRTFStream(LPSTREAM lpCompressedRTFStream,
										 ULONG ulFlags,
										 LPSTREAM FAR * lpUncompressedRTFStream)
{
	CHECKLOAD(pfnWrapCompressedRTFStream);
	if (pfnWrapCompressedRTFStream) return pfnWrapCompressedRTFStream(
		lpCompressedRTFStream,
		ulFlags,
		lpUncompressedRTFStream);
	return MAPI_E_CALL_FAILED;
} // WrapCompressedRTFStream

STDMETHODIMP MAPILogonEx(ULONG_PTR ulUIParam,
						 LPTSTR lpszProfileName,
						 LPTSTR lpszPassword,
						 ULONG ulFlags,
						 LPMAPISESSION FAR * lppSession)
{
	CHECKLOAD(pfnMAPILogonEx);
	if (pfnMAPILogonEx) return pfnMAPILogonEx(
		ulUIParam,
		lpszProfileName,
		lpszPassword,
		ulFlags,
		lppSession);
	return MAPI_E_CALL_FAILED;
} // MAPILogonEx

STDMETHODIMP MAPIAdminProfiles(ULONG ulFlags,
							   LPPROFADMIN FAR *lppProfAdmin)
{
	CHECKLOAD(pfnMAPIAdminProfiles);
	if (pfnMAPIAdminProfiles) return pfnMAPIAdminProfiles(
		ulFlags,
		lppProfAdmin);
	return MAPI_E_CALL_FAILED;
} // MAPIAdminProfiles

STDAPI HrValidateIPMSubtree(LPMDB lpMDB,
							ULONG ulFlags,
							ULONG FAR *lpcValues,
							LPSPropValue FAR *lppValues,
							LPMAPIERROR FAR *lpperr)
{
	CHECKLOAD(pfnHrValidateIPMSubtree);
	if (pfnHrValidateIPMSubtree) return pfnHrValidateIPMSubtree(
		lpMDB,
		ulFlags,
		lpcValues,
		lppValues,
		lpperr);
	return MAPI_E_CALL_FAILED;
} // HrValidateIPMSubtree

STDAPI MAPIOpenLocalFormContainer(LPMAPIFORMCONTAINER FAR * ppfcnt)
{
	CHECKLOAD(pfnMAPIOpenLocalFormContainer);
	if (pfnMAPIOpenLocalFormContainer) return pfnMAPIOpenLocalFormContainer(
		ppfcnt);
	return MAPI_E_CALL_FAILED;
} // MAPIOpenLocalFormContainer

STDAPI HrDispatchNotifications(ULONG ulFlags)
{
	CHECKLOAD(pfnHrDispatchNotifications);
	if (pfnHrDispatchNotifications) return pfnHrDispatchNotifications(
		ulFlags);
	return MAPI_E_CALL_FAILED;
} // HrDispatchNotifications

STDAPI WrapStoreEntryID(ULONG ulFlags, __in LPTSTR lpszDLLName, ULONG cbOrigEntry,
						LPENTRYID lpOrigEntry, ULONG *lpcbWrappedEntry, LPENTRYID *lppWrappedEntry)
{
	CHECKLOAD(pfnWrapStoreEntryID);
	if (pfnWrapStoreEntryID) return pfnWrapStoreEntryID(ulFlags, lpszDLLName, cbOrigEntry, lpOrigEntry, lpcbWrappedEntry, lppWrappedEntry);
	return MAPI_E_CALL_FAILED;
} // WrapStoreEntryID

STDAPI_(SCODE)
CreateIProp( LPCIID					lpInterface,
			ALLOCATEBUFFER FAR *	lpAllocateBuffer,
			ALLOCATEMORE FAR *		lpAllocateMore,
			FREEBUFFER FAR *		lpFreeBuffer,
			LPVOID					lpvReserved,
			LPPROPDATA FAR *		lppPropData )
{
	CHECKLOAD(pfnCreateIProp);
	if (pfnCreateIProp) return pfnCreateIProp(lpInterface, lpAllocateBuffer, lpAllocateMore, lpFreeBuffer, lpvReserved, lppPropData);
	return MAPI_E_CALL_FAILED;
} // CreateIProp

STDAPI_(SCODE)
CreateTable( LPCIID					lpInterface,
			ALLOCATEBUFFER FAR *	lpAllocateBuffer,
			ALLOCATEMORE FAR *		lpAllocateMore,
			FREEBUFFER FAR *		lpFreeBuffer,
			LPVOID					lpvReserved,
			ULONG					ulTableType,
			ULONG					ulPropTagIndexColumn,
			LPSPropTagArray		lpSPropTagArrayColumns,
			LPTABLEDATA FAR *		lppTableData )
{
	CHECKLOAD(pfnCreateTable);
	if (pfnCreateTable) return pfnCreateTable(
		lpInterface,
		lpAllocateBuffer,
		lpAllocateMore,
		lpFreeBuffer,
		lpvReserved,
		ulTableType,
		ulPropTagIndexColumn,
		lpSPropTagArrayColumns,
		lppTableData);
	return MAPI_E_CALL_FAILED;
} // CreateTable

STDAPI HrValidateParameters( METHODS eMethod, LPVOID FAR *ppFirstArg)
{
	CHECKLOAD(pfnHrValidateParameters);
	if (pfnHrValidateParameters) return pfnHrValidateParameters(
		eMethod,
		ppFirstArg);
	return MAPI_E_CALL_FAILED;
} // HrValidateParameters

STDAPI
BuildDisplayTable(	LPALLOCATEBUFFER	lpAllocateBuffer,
				  LPALLOCATEMORE		lpAllocateMore,
				  LPFREEBUFFER		lpFreeBuffer,
				  LPMALLOC			lpMalloc,
				  HINSTANCE			hInstance,
				  UINT				cPages,
				  LPDTPAGE			lpPage,
				  ULONG				ulFlags,
				  LPMAPITABLE *		lppTable,
				  LPTABLEDATA	*		lppTblData )
{
	CHECKLOAD(pfnBuildDisplayTable);
	if (pfnBuildDisplayTable) return pfnBuildDisplayTable(
		lpAllocateBuffer,
		lpAllocateMore,
		lpFreeBuffer,
		lpMalloc,
		hInstance,
		cPages,
		lpPage,
		ulFlags,
		lppTable,
		lppTblData);
	return MAPI_E_CALL_FAILED;
} // BuildDisplayTable

int WINAPI MNLS_lstrlenW(LPCWSTR lpString)
{
	CHECKLOAD(pfnMNLS_lstrlenW);
	if (pfnMNLS_lstrlenW) return pfnMNLS_lstrlenW(
		lpString);
	return 0;
} // MNLS_lstrlenW

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