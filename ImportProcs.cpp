#include "stdafx.h"
#include "ImportProcs.h"
#include "String.h"

HMODULE hModAclui = nullptr;
HMODULE hModOle32 = nullptr;
HMODULE hModUxTheme = nullptr;
HMODULE hModInetComm = nullptr;
HMODULE hModMSI = nullptr;
HMODULE hModKernel32 = nullptr;
HMODULE hModShell32 = nullptr;

typedef HTHEME(STDMETHODCALLTYPE CLOSETHEMEDATA)
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
	DWORD dwFlags,
	LPCWSTR lpModuleName,
	HMODULE* phModule);
typedef GETMODULEHANDLEEXW* LPGETMODULEHANDLEEXW;

typedef HRESULT(STDMETHODCALLTYPE MIMEOLEGETCODEPAGECHARSET)
(
	CODEPAGEID cpiCodePage,
	CHARSETTYPE ctCsetType,
	LPHCHARSET phCharset
	);
typedef MIMEOLEGETCODEPAGECHARSET* LPMIMEOLEGETCODEPAGECHARSET;

typedef UINT(WINAPI MSIPROVIDECOMPONENTW)
(
	LPCWSTR szProduct,
	LPCWSTR szFeature,
	LPCWSTR szComponent,
	DWORD dwInstallMode,
	LPWSTR lpPathBuf,
	LPDWORD pcchPathBuf
	);
typedef MSIPROVIDECOMPONENTW FAR * LPMSIPROVIDECOMPONENTW;

typedef UINT(WINAPI MSIPROVIDEQUALIFIEDCOMPONENTW)
(
	LPCWSTR szCategory,
	LPCWSTR szQualifier,
	DWORD dwInstallMode,
	LPWSTR lpPathBuf,
	LPDWORD pcchPathBuf
	);
typedef MSIPROVIDEQUALIFIEDCOMPONENTW FAR * LPMSIPROVIDEQUALIFIEDCOMPONENTW;

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

// Exists to allow some logging
_Check_return_ HMODULE MyLoadLibraryA(_In_z_ LPCSTR lpszLibFileName)
{
	HMODULE hMod = nullptr;
	auto hRes = S_OK;
	DebugPrint(DBGLoadLibrary, L"MyLoadLibraryA - loading \"%hs\"\n", lpszLibFileName);
	WC_D(hMod, LoadLibraryA(lpszLibFileName));
	if (hMod)
	{
		DebugPrint(DBGLoadLibrary, L"MyLoadLibraryA - \"%hs\" loaded at %p\n", lpszLibFileName, hMod);
	}
	else
	{
		DebugPrint(DBGLoadLibrary, L"MyLoadLibraryA - \"%hs\" failed to load\n", lpszLibFileName);
	}
	return hMod;
}

// Exists to allow some logging
_Check_return_ HMODULE MyLoadLibraryW(_In_z_ LPCWSTR lpszLibFileName)
{
	HMODULE hMod = nullptr;
	auto hRes = S_OK;
	DebugPrint(DBGLoadLibrary, L"MyLoadLibraryW - loading \"%ws\"\n", lpszLibFileName);
	WC_D(hMod, LoadLibraryW(lpszLibFileName));
	if (hMod)
	{
		DebugPrint(DBGLoadLibrary, L"MyLoadLibraryW - \"%ws\" loaded at %p\n", lpszLibFileName, hMod);
	}
	else
	{
		DebugPrint(DBGLoadLibrary, L"MyLoadLibraryW - \"%ws\" failed to load\n", lpszLibFileName);
	}
	return hMod;
}

// Loads szModule at the handle given by lphModule, then looks for szEntryPoint.
// Will not load a module or entry point twice
void LoadProc(LPCTSTR szModule, HMODULE* lphModule, LPCSTR szEntryPoint, FARPROC* lpfn)
{
	if (!szEntryPoint || !lpfn || !lphModule) return;
	if (*lpfn) return;
	if (!*lphModule && szModule)
	{
		*lphModule = LoadFromSystemDir(szModule);
	}

	if (!*lphModule) return;

	auto hRes = S_OK;
	WC_D(*lpfn, GetProcAddress(
		*lphModule,
		szEntryPoint));
}

_Check_return_ HMODULE LoadFromSystemDir(_In_z_ LPCTSTR szDLLName)
{
	if (!szDLLName) return nullptr;

	auto hRes = S_OK;
	HMODULE hModRet = nullptr;
	TCHAR szDLLPath[MAX_PATH] = { 0 };
	UINT uiRet = NULL;

	static TCHAR szSystemDir[MAX_PATH] = { 0 };
	static auto bSystemDirLoaded = false;

	DebugPrint(DBGLoadLibrary, L"LoadFromSystemDir - loading \"%ws\"\n", LPCTSTRToWstring(szDLLName).c_str());

	if (!bSystemDirLoaded)
	{
		WC_D(uiRet, GetSystemDirectory(szSystemDir, MAX_PATH));
		bSystemDirLoaded = true;
	}

	WC_H(StringCchPrintf(szDLLPath, _countof(szDLLPath), _T("%s\\%s"), szSystemDir, szDLLName)); // STRING_OK
	DebugPrint(DBGLoadLibrary, L"LoadFromSystemDir - loading from \"%ws\"\n", LPCTSTRToWstring(szDLLPath).c_str());
	hModRet = MyLoadLibrary(szDLLPath);

	return hModRet;
}

_Check_return_ HMODULE LoadFromOLMAPIDir(_In_z_ LPCTSTR szDLLName)
{
	HMODULE hModRet = nullptr;

	DebugPrint(DBGLoadLibrary, L"LoadFromOLMAPIDir - loading \"%ws\"\n", LPCTSTRToWstring(szDLLName).c_str());

	for (auto i = oqcOfficeBegin; i < oqcOfficeEnd; i++)
	{
		auto szOutlookMAPIPath = GetInstalledOutlookMAPI(i);
		if (!szOutlookMAPIPath.empty())
		{
			auto hRes = S_OK;
			UINT ret = 0;
			WCHAR szDrive[_MAX_DRIVE] = { 0 };
			WCHAR szMAPIPath[MAX_PATH] = { 0 };
			WC_D(ret, _wsplitpath_s(szOutlookMAPIPath.c_str(), szDrive, _MAX_DRIVE, szMAPIPath, MAX_PATH, NULL, NULL, NULL, NULL));

			if (SUCCEEDED(hRes))
			{
				auto szFullPath = new WCHAR[MAX_PATH];
#ifdef UNICODE
					swprintf_s(szFullPath, MAX_PATH, L"%s%s%s", szDrive, szMAPIPath, szDLLName);
#else
				swprintf_s(szFullPath, MAX_PATH, L"%s%s%hs", szDrive, szMAPIPath, szDLLName);
#endif
				DebugPrint(DBGLoadLibrary, L"LoadFromOLMAPIDir - loading from \"%ws\"\n", szFullPath);
				WC_D(hModRet, MyLoadLibraryW(szFullPath));
				delete[] szFullPath;

			}
		}

		if (hModRet) break;
	}

	return hModRet;
}

void ImportProcs()
{
	LoadProc(_T("aclui.dll"), &hModAclui, "EditSecurity", reinterpret_cast<FARPROC*>(&pfnEditSecurity)); // STRING_OK;
	LoadProc(_T("ole32.dll"), &hModOle32, "StgCreateStorageEx", reinterpret_cast<FARPROC*>(&pfnStgCreateStorageEx)); // STRING_OK;
	LoadProc(_T("uxtheme.dll"), &hModUxTheme, "OpenThemeData", reinterpret_cast<FARPROC*>(&pfnOpenThemeData)); // STRING_OK;
	LoadProc(_T("uxtheme.dll"), &hModUxTheme, "CloseThemeData", reinterpret_cast<FARPROC*>(&pfnCloseThemeData)); // STRING_OK;
	LoadProc(_T("uxtheme.dll"), &hModUxTheme, "GetThemeMargins", reinterpret_cast<FARPROC*>(&pfnGetThemeMargins)); // STRING_OK;
	LoadProc(_T("uxtheme.dll"), &hModUxTheme, "SetWindowTheme", reinterpret_cast<FARPROC*>(&pfnSetWindowTheme)); // STRING_OK;
	LoadProc(_T("uxtheme.dll"), &hModUxTheme, "GetThemeSysSize", reinterpret_cast<FARPROC*>(&pfnGetThemeSysSize)); // STRING_OK;
	LoadProc(_T("msi.dll"), &hModMSI, "MsiGetFileVersionW", reinterpret_cast<FARPROC*>(&pfnMsiGetFileVersion)); // STRING_OK;
	LoadProc(_T("msi.dll"), &hModMSI, "MsiProvideQualifiedComponentW", reinterpret_cast<FARPROC*>(&pfnMsiProvideQualifiedComponent)); // STRING_OK;
	LoadProc(_T("shell32.dll"), &hModShell32, "SHGetPropertyStoreForWindow", reinterpret_cast<FARPROC*>(&pfnSHGetPropertyStoreForWindow)); // STRING_OK;
}

// Opens the mail key for the specified MAPI client, such as 'Microsoft Outlook' or 'ExchangeMAPI'
// Pass NULL to open the mail key for the default MAPI client
_Check_return_ HKEY GetMailKey(_In_opt_z_ LPCTSTR szClient)
{
	wstring lpszClient = L"Default";
	if (szClient) lpszClient = LPCTSTRToWstring(szClient);
	DebugPrint(DBGLoadLibrary, L"Enter GetMailKey(%ws)\n", lpszClient.c_str());
	auto hRes = S_OK;
	HKEY hMailKey = nullptr;
	auto bClientIsDefault = false;

	// If szClient is NULL, we need to read the name of the default MAPI client
	if (!szClient)
	{
		HKEY hDefaultMailKey = nullptr;
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
				(LPVOID*)&szClient));
			if (szClient) lpszClient = LPCTSTRToWstring(szClient);
			DebugPrint(DBGLoadLibrary, L"Default MAPI = %ws\n", lpszClient.c_str());
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
}

// Gets MSI IDs for the specified MAPI client, such as 'Microsoft Outlook' or 'ExchangeMAPI'
// Pass NULL to get the IDs for the default MAPI client
// Allocates with new, delete with delete[]
void GetMapiMsiIds(_In_opt_z_ LPCTSTR szClient, _Deref_out_opt_z_ LPSTR* lpszComponentID, _Deref_out_opt_z_ LPSTR* lpszAppLCID, _Deref_out_opt_z_ LPSTR* lpszOfficeLCID)
{
	DebugPrint(DBGLoadLibrary, L"GetMapiMsiIds(%ws)\n", LPCTSTRToWstring(szClient).c_str());
	auto hRes = S_OK;

	if (lpszComponentID) *lpszComponentID = nullptr;
	if (lpszAppLCID) *lpszAppLCID = nullptr;
	if (lpszOfficeLCID) *lpszOfficeLCID = nullptr;

	auto hKey = GetMailKey(szClient);

	if (hKey)
	{
		DWORD dwKeyType = NULL;

		if (lpszComponentID)
		{
			WC_H(HrGetRegistryValueA(
				hKey,
				"MSIComponentID", // STRING_OK
				&dwKeyType,
				reinterpret_cast<LPVOID*>(lpszComponentID)));
			DebugPrint(DBGLoadLibrary, L"MSIComponentID = %hs\n", *lpszComponentID ? *lpszComponentID : "<not found>");
			hRes = S_OK;
		}

		if (lpszAppLCID)
		{
			WC_H(HrGetRegistryValueA(
				hKey,
				"MSIApplicationLCID", // STRING_OK
				&dwKeyType,
				reinterpret_cast<LPVOID*>(lpszAppLCID)));
			DebugPrint(DBGLoadLibrary, L"MSIApplicationLCID = %hs\n", *lpszAppLCID ? *lpszAppLCID : "<not found>");
			hRes = S_OK;
		}

		if (lpszOfficeLCID)
		{
			WC_H(HrGetRegistryValueA(
				hKey,
				"MSIOfficeLCID", // STRING_OK
				&dwKeyType,
				reinterpret_cast<LPVOID*>(lpszOfficeLCID)));
			DebugPrint(DBGLoadLibrary, L"MSIOfficeLCID = %hs\n", *lpszOfficeLCID ? *lpszOfficeLCID : "<not found>");
			hRes = S_OK;
		}

		EC_W32(RegCloseKey(hKey));
	}
}

wstring GetMAPIPath(wstring szClient)
{
	auto hRes = S_OK;
	CHAR lpszPath[MAX_PATH] = { 0 };
	ULONG cchPath = _countof(lpszPath);

	// Find some strings:
	LPSTR szComponentID = nullptr;
	LPSTR szAppLCID = nullptr;
	LPSTR szOfficeLCID = nullptr;

	GetMapiMsiIds(wstringToCString(szClient), &szComponentID, &szAppLCID, &szOfficeLCID);

	if (szComponentID)
	{
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
	}

	delete[] szComponentID;
	delete[] szOfficeLCID;
	delete[] szAppLCID;

	return LPCSTRToWstring(lpszPath);
}

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
	auto hRes = OpenStreamOnFileW(
		lpAllocateBuffer,
		lpFreeBuffer,
		ulFlags,
		lpszFileName,
		nullptr,
		lppStream);
	if (MAPI_E_CALL_FAILED == hRes)
	{
		hRes = OpenStreamOnFile(
			lpAllocateBuffer,
			lpFreeBuffer,
			ulFlags,
			wstringToCString(lpszFileName),
			nullptr,
			lppStream);
	}
	return hRes;
}

_Check_return_ HRESULT HrDupPropset(
	int cprop,
	_In_count_(cprop) LPSPropValue rgprop,
	_In_ LPVOID lpObject,
	_In_ LPSPropValue* prgprop)
{
	ULONG cb = NULL;
	auto hRes = S_OK;

	// Find out how much memory we need
	EC_MAPI(ScCountProps(cprop, rgprop, &cb));

	if (SUCCEEDED(hRes) && cb)
	{
		// Obtain memory
		if (lpObject != nullptr)
		{
			EC_H(MAPIAllocateMore(cb, lpObject, reinterpret_cast<LPVOID*>(prgprop)));
		}
		else
		{
			EC_H(MAPIAllocateBuffer(cb, reinterpret_cast<LPVOID*>(prgprop)));
		}

		if (SUCCEEDED(hRes) && prgprop)
		{
			// Copy the properties
			EC_MAPI(ScCopyProps(cprop, rgprop, *prgprop, &cb));
		}
	}

	return hRes;
}

typedef ACTIONS* LPACTIONS;

// swiped from EDK rules sample
_Check_return_ STDAPI HrCopyActions(
	_In_ LPACTIONS lpActsSrc, // source action ptr
	_In_ LPVOID lpObject, // ptr to existing MAPI buffer
	_In_ LPACTIONS* lppActsDst) // ptr to destination ACTIONS buffer
{
	if (!lpActsSrc || !lppActsDst) return MAPI_E_INVALID_PARAMETER;
	if (lpActsSrc->cActions <= 0 || lpActsSrc->lpAction == nullptr) return MAPI_E_INVALID_PARAMETER;

	auto fNullObject = lpObject == nullptr;
	auto hRes = S_OK;

	*lppActsDst = nullptr;

	if (lpObject != nullptr)
	{
		WC_H(MAPIAllocateMore(sizeof(ACTIONS), lpObject, reinterpret_cast<LPVOID*>(lppActsDst)));
	}
	else
	{
		WC_H(MAPIAllocateBuffer(sizeof(ACTIONS), reinterpret_cast<LPVOID*>(lppActsDst)));
		lpObject = *lppActsDst;
	}

	if (FAILED(hRes)) return hRes;
	// no short circuit returns after here

	auto lpActsDst = *lppActsDst;
	*lpActsDst = *lpActsSrc;
	lpActsDst->lpAction = nullptr;

	WC_H(MAPIAllocateMore(sizeof(ACTION) * lpActsDst->cActions,
		lpObject,
		reinterpret_cast<LPVOID*>(&lpActsDst->lpAction)));
	if (SUCCEEDED(hRes) && lpActsDst->lpAction)
	{
		// Initialize acttype values for all members of the array to a value
		// that will not cause deallocation errors should the copy fail.
		for (ULONG i = 0; i < lpActsDst->cActions; i++)
			lpActsDst->lpAction[i].acttype = OP_BOUNCE;

		// Now actually copy all the members of the array.
		for (ULONG i = 0; i < lpActsDst->cActions; i++)
		{
			auto lpActDst = &lpActsDst->lpAction[i];
			auto lpActSrc = &lpActsSrc->lpAction[i];

			*lpActDst = *lpActSrc;

			switch (lpActSrc->acttype)
			{
			case OP_MOVE: // actMoveCopy
			case OP_COPY:
				if (lpActDst->actMoveCopy.cbStoreEntryId &&
					lpActDst->actMoveCopy.lpStoreEntryId)
				{
					WC_H(MAPIAllocateMore(lpActDst->actMoveCopy.cbStoreEntryId,
						lpObject,
						reinterpret_cast<LPVOID*>(&lpActDst->actMoveCopy.lpStoreEntryId)));
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
						reinterpret_cast<LPVOID*>(&lpActDst->actMoveCopy.lpFldEntryId)));
					if (FAILED(hRes)) break;

					memcpy(lpActDst->actMoveCopy.lpFldEntryId,
						lpActSrc->actMoveCopy.lpFldEntryId,
						lpActSrc->actMoveCopy.cbFldEntryId);
				}

				break;

			case OP_REPLY: // actReply
			case OP_OOF_REPLY:
				if (lpActDst->actReply.cbEntryId &&
					lpActDst->actReply.lpEntryId)
				{
					WC_H(MAPIAllocateMore(lpActDst->actReply.cbEntryId,
						lpObject,
						reinterpret_cast<LPVOID*>(&lpActDst->actReply.lpEntryId)));
					if (FAILED(hRes)) break;

					memcpy(lpActDst->actReply.lpEntryId,
						lpActSrc->actReply.lpEntryId,
						lpActSrc->actReply.cbEntryId);
				}
				break;

			case OP_DEFER_ACTION: // actDeferAction
				if (lpActSrc->actDeferAction.pbData &&
					lpActSrc->actDeferAction.cbData)
				{
					WC_H(MAPIAllocateMore(lpActDst->actDeferAction.cbData,
						lpObject,
						reinterpret_cast<LPVOID*>(&lpActDst->actDeferAction.pbData)));
					if (FAILED(hRes)) break;

					memcpy(lpActDst->actDeferAction.pbData,
						lpActSrc->actDeferAction.pbData,
						lpActDst->actDeferAction.cbData);
				}
				break;

			case OP_FORWARD: // lpadrlist
			case OP_DELEGATE:
				lpActDst->lpadrlist = nullptr;

				if (lpActSrc->lpadrlist && lpActSrc->lpadrlist->cEntries)
				{
					WC_H(MAPIAllocateMore(CbADRLIST(lpActSrc->lpadrlist),
						lpObject,
						reinterpret_cast<LPVOID*>(&lpActDst->lpadrlist)));
					if (FAILED(hRes)) break;

					lpActDst->lpadrlist->cEntries = lpActSrc->lpadrlist->cEntries;

					// Initialize the new ADRENTRYs and validate cValues.
					for (ULONG j = 0; j < lpActSrc->lpadrlist->cEntries; j++)
					{
						lpActDst->lpadrlist->aEntries[j] = lpActSrc->lpadrlist->aEntries[j];
						lpActDst->lpadrlist->aEntries[j].rgPropVals = nullptr;

						if (lpActDst->lpadrlist->aEntries[j].cValues == 0)
						{
							hRes = MAPI_E_INVALID_PARAMETER;
							break;
						}
					}

					// Copy the rgPropVals.
					for (ULONG j = 0; j < lpActSrc->lpadrlist->cEntries; j++)
					{
						WC_MAPI(HrDupPropset(
							lpActDst->lpadrlist->aEntries[j].cValues,
							lpActSrc->lpadrlist->aEntries[j].rgPropVals,
							lpObject,
							&lpActDst->lpadrlist->aEntries[j].rgPropVals));
						if (FAILED(hRes)) break;
					}
				}

				break;

			case OP_TAG: // propTag
				WC_H(MyPropCopyMore(
					&lpActDst->propTag,
					&lpActSrc->propTag,
					MAPIAllocateMore,
					lpObject));
				if (FAILED(hRes)) break;

				break;

			case OP_BOUNCE: // scBounceCode
			case OP_DELETE: // union not used
			case OP_MARK_AS_READ:
				break; // Nothing to do!

			default: // error!
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
}

_Check_return_ HRESULT HrCopyRestrictionArray(
	_In_ LPSRestriction lpResSrc, // source restriction
	_In_ LPVOID lpObject, // ptr to existing MAPI buffer
	ULONG cRes, // # elements in array
	_In_count_(cRes) LPSRestriction lpResDest) // destination restriction
{
	if (!lpResSrc || !lpResDest || !lpObject) return MAPI_E_INVALID_PARAMETER;
	auto hRes = S_OK;

	for (ULONG i = 0; i < cRes; i++)
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
				if (lpResSrc[i].res.resAnd.cRes > ULONG_MAX / sizeof(SRestriction))
				{
					hRes = MAPI_E_CALL_FAILED;
					break;
				}
				WC_H(MAPIAllocateMore(sizeof(SRestriction) * lpResSrc[i].res.resAnd.cRes,
					lpObject,
					reinterpret_cast<LPVOID*>(&lpResDest[i].res.resAnd.lpRes)));
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
					reinterpret_cast<LPVOID*>(&lpResDest[i].res.resNot.lpRes)));
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
				WC_MAPI(HrDupPropset(
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
					reinterpret_cast<LPVOID*>(&lpResDest[i].res.resSub.lpRes)));
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
					reinterpret_cast<LPVOID*>(&lpResDest[i].res.resComment.lpRes)));
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
				WC_MAPI(HrDupPropset(
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
}

_Check_return_ STDAPI HrCopyRestriction(
	_In_ LPSRestriction lpResSrc, // source restriction ptr
	_In_opt_ LPVOID lpObject, // ptr to existing MAPI buffer
	_In_ LPSRestriction* lppResDest) // dest restriction buffer ptr
{
	if (!lppResDest) return MAPI_E_INVALID_PARAMETER;
	*lppResDest = nullptr;
	if (!lpResSrc) return S_OK;

	auto fNullObject = lpObject == nullptr;
	auto hRes = S_OK;

	if (lpObject != nullptr)
	{
		WC_H(MAPIAllocateMore(sizeof(SRestriction),
			lpObject,
			reinterpret_cast<LPVOID*>(lppResDest)));
	}
	else
	{
		WC_H(MAPIAllocateBuffer(sizeof(SRestriction),
			reinterpret_cast<LPVOID*>(lppResDest)));
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
}

// This augmented PropCopyMore is implicitly tied to the built-in MAPIAllocateMore and MAPIAllocateBuffer through
// the calls to HrCopyRestriction and HrCopyActions. Rewriting those functions to accept function pointers is
// expensive for no benefit here. So if you borrow this code, be careful if you plan on using other allocators.
_Check_return_ STDAPI_(SCODE) MyPropCopyMore(_In_ LPSPropValue lpSPropValueDest,
	_In_ LPSPropValue lpSPropValueSrc,
	_In_ ALLOCATEMORE * lpfAllocMore,
	_In_ LPVOID lpvObject)
{
	auto hRes = S_OK;
	switch (PROP_TYPE(lpSPropValueSrc->ulPropTag))
	{
	case PT_SRESTRICTION:
	case PT_ACTIONS:
	{
		// It's an action or restriction - we know how to copy those:
		memcpy(reinterpret_cast<BYTE *>(lpSPropValueDest),
			reinterpret_cast<BYTE *>(lpSPropValueSrc),
			sizeof(SPropValue));
		if (PT_SRESTRICTION == PROP_TYPE(lpSPropValueSrc->ulPropTag))
		{
			LPSRestriction lpNewRes = nullptr;
			WC_H(HrCopyRestriction(
				reinterpret_cast<LPSRestriction>(lpSPropValueSrc->Value.lpszA),
				lpvObject,
				&lpNewRes));
			lpSPropValueDest->Value.lpszA = reinterpret_cast<LPSTR>(lpNewRes);
		}
		else
		{
			ACTIONS* lpNewAct = nullptr;
			WC_H(HrCopyActions(
				reinterpret_cast<ACTIONS*>(lpSPropValueSrc->Value.lpszA),
				lpvObject,
				&lpNewAct));
			lpSPropValueDest->Value.lpszA = reinterpret_cast<LPSTR>(lpNewAct);
		}
		break;
	}
	default:
		WC_MAPI(PropCopyMore(lpSPropValueDest, lpSPropValueSrc, lpfAllocMore, lpvObject));
	}
	return hRes;
}

void WINAPI MyHeapSetInformation(_In_opt_ HANDLE HeapHandle,
	_In_ HEAP_INFORMATION_CLASS HeapInformationClass,
	_In_opt_count_(HeapInformationLength) PVOID HeapInformation,
	_In_ SIZE_T HeapInformationLength)
{
	if (!pfnHeapSetInformation)
	{
		LoadProc(_T("kernel32.dll"), &hModKernel32, "HeapSetInformation", reinterpret_cast<FARPROC*>(&pfnHeapSetInformation)); // STRING_OK;
	}

	if (pfnHeapSetInformation) pfnHeapSetInformation(HeapHandle, HeapInformationClass, HeapInformation, HeapInformationLength);
}

HRESULT WINAPI MyMimeOleGetCodePageCharset(
	CODEPAGEID cpiCodePage,
	CHARSETTYPE ctCsetType,
	LPHCHARSET phCharset)
{
	if (!pfnMimeOleGetCodePageCharset)
	{
		LoadProc(_T("inetcomm.dll"), &hModInetComm, "MimeOleGetCodePageCharset", reinterpret_cast<FARPROC*>(&pfnMimeOleGetCodePageCharset)); // STRING_OK;
	}

	if (pfnMimeOleGetCodePageCharset) return pfnMimeOleGetCodePageCharset(cpiCodePage, ctCsetType, phCharset);
	return MAPI_E_CALL_FAILED;
}

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
		LoadProc(_T("msi.dll"), &hModMSI, "MimeOleGetCodePageCharset", reinterpret_cast<FARPROC*>(&pfnMsiProvideComponentW)); // STRING_OK;
	}

	if (pfnMsiProvideComponentW) return pfnMsiProvideComponentW(szProduct, szFeature, szComponent, dwInstallMode, lpPathBuf, pcchPathBuf);
	return ERROR_NOT_SUPPORTED;
}

STDAPI_(UINT) MsiProvideQualifiedComponentW(
	LPCWSTR szCategory,
	LPCWSTR szQualifier,
	DWORD dwInstallMode,
	LPWSTR lpPathBuf,
	LPDWORD pcchPathBuf)
{
	if (!pfnMsiProvideQualifiedComponentW)
	{
		LoadProc(_T("msi.dll"), &hModMSI, "MsiProvideQualifiedComponentW", reinterpret_cast<FARPROC*>(&pfnMsiProvideQualifiedComponentW)); // STRING_OK;
	}

	if (pfnMsiProvideQualifiedComponentW) return pfnMsiProvideQualifiedComponentW(szCategory, szQualifier, dwInstallMode, lpPathBuf, pcchPathBuf);
	return ERROR_NOT_SUPPORTED;
}

BOOL WINAPI MyGetModuleHandleExW(
	DWORD dwFlags,
	LPCWSTR lpModuleName,
	HMODULE* phModule)
{
	if (!pfnGetModuleHandleExW)
	{
		LoadProc(_T("kernel32.dll"), &hModMSI, "GetModuleHandleExW", reinterpret_cast<FARPROC*>(&pfnGetModuleHandleExW)); // STRING_OK;
	}

	if (pfnGetModuleHandleExW) return pfnGetModuleHandleExW(dwFlags, lpModuleName, phModule);
	*phModule = GetModuleHandleW(lpModuleName);
	return *phModule != nullptr;
}