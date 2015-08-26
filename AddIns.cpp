// AddIns.cpp : Functions supporting AddIns

#include "stdafx.h"
#include "BaseDialog.h"
#include "MFCMAPI.h"
#include "ImportProcs.h"
#include "String.h"
#include "Editor.h"
#include "Guids.h"
#ifndef MRMAPI
#include "UIFunctions.h"
#endif

// Our built in arrays, which get merged into the arrays declared in mfcmapi.h
#include "GenTagArray.h"
#include "Flags.h"
#include "GuidArray.h"
#include "NameIDArray.h"
#include "PropTypeArray.h"
#include "SmartView\SmartViewParsers.h"

LPADDIN g_lpMyAddins = NULL;

// Count of built-in arrays
ULONG g_ulPropTagArray = _countof(g_PropTagArray);
ULONG g_ulFlagArray = _countof(g_FlagArray);
ULONG g_ulPropGuidArray = _countof(g_PropGuidArray);
ULONG g_ulNameIDArray = _countof(g_NameIDArray);
ULONG g_ulPropTypeArray = _countof(g_PropTypeArray);
ULONG g_ulSmartViewParserArray = _countof(g_SmartViewParserArray);
ULONG g_ulSmartViewParserTypeArray = _countof(g_SmartViewParserTypeArray);

// Arrays and counts declared in mfcmapi.h
LPNAME_ARRAY_ENTRY PropTypeArray = NULL;
ULONG ulPropTypeArray = NULL;

LPNAME_ARRAY_ENTRY_V2 PropTagArray = NULL;
ULONG ulPropTagArray = NULL;

LPGUID_ARRAY_ENTRY PropGuidArray = NULL;
ULONG ulPropGuidArray = NULL;

LPNAMEID_ARRAY_ENTRY NameIDArray = NULL;
ULONG ulNameIDArray;

LPFLAG_ARRAY_ENTRY FlagArray = NULL;
ULONG ulFlagArray = NULL;

LPSMARTVIEW_PARSER_ARRAY_ENTRY SmartViewParserArray = NULL;
ULONG ulSmartViewParserArray;

LPNAME_ARRAY_ENTRY SmartViewParserTypeArray = NULL;
ULONG ulSmartViewParserTypeArray;

_Check_return_ ULONG GetAddinVersion(HMODULE hMod)
{
	HRESULT hRes = S_OK;
	LPGETAPIVERSION pfnGetAPIVersion = NULL;
	WC_D(pfnGetAPIVersion, (LPGETAPIVERSION)GetProcAddress(hMod, szGetAPIVersion));
	if (pfnGetAPIVersion)
	{
		return pfnGetAPIVersion();
	}

	// Default case for unversioned add-ins
	return MFCMAPI_HEADER_V1;
} // GetAddinVersion

// Load NAME_ARRAY_ENTRY style props from the add-in and upconvert them to V2
void LoadLegacyPropTags(
	HMODULE hMod,
	_In_ ULONG* lpulPropTags, // Number of entries in lppPropTags
	_In_ LPNAME_ARRAY_ENTRY_V2* lppPropTags // Array of NAME_ARRAY_ENTRY_V2 structures
	)
{
	HRESULT hRes = S_OK;
	LPGETPROPTAGS pfnGetPropTags = NULL;
	LPNAME_ARRAY_ENTRY lpPropTags = NULL;
	WC_D(pfnGetPropTags, (LPGETPROPTAGS)GetProcAddress(hMod, szGetPropTags));
	if (pfnGetPropTags)
	{
		pfnGetPropTags(lpulPropTags, &lpPropTags);
		if (lpPropTags && *lpulPropTags)
		{
			LPNAME_ARRAY_ENTRY_V2 lpPropTagsV2 = new NAME_ARRAY_ENTRY_V2[*lpulPropTags];
			if (lpPropTagsV2)
			{
				ULONG i = 0;
				for (i = 0; i < *lpulPropTags; i++)
				{
					lpPropTagsV2[i].lpszName = lpPropTags[i].lpszName;
					lpPropTagsV2[i].ulValue = lpPropTags[i].ulValue;
					lpPropTagsV2[i].ulSortOrder = 0;
				}
				*lppPropTags = lpPropTagsV2;
			}
		}
	}
} // LoadLegacyPropTags

void LoadSingleAddIn(_In_ LPADDIN lpAddIn, HMODULE hMod, _In_ LPLOADADDIN pfnLoadAddIn)
{
	DebugPrint(DBGAddInPlumbing, L"Loading AddIn\n");
	if (!lpAddIn) return;
	if (!pfnLoadAddIn) return;
	HRESULT hRes = S_OK;
	lpAddIn->hMod = hMod;
	pfnLoadAddIn(&lpAddIn->szName);
	if (lpAddIn->szName)
	{
		DebugPrint(DBGAddInPlumbing, L"Loading \"%ws\"\n", lpAddIn->szName);
	}

	ULONG ulVersion = GetAddinVersion(hMod);
	DebugPrint(DBGAddInPlumbing, L"AddIn version = %u\n", ulVersion);

	LPGETMENUS pfnGetMenus = NULL;
	WC_D(pfnGetMenus, (LPGETMENUS)GetProcAddress(hMod, szGetMenus));
	if (pfnGetMenus)
	{
		pfnGetMenus(&lpAddIn->ulMenu, &lpAddIn->lpMenu);
		if (!lpAddIn->ulMenu || !lpAddIn->lpMenu)
		{
			DebugPrint(DBGAddInPlumbing, L"AddIn returned invalid menus\n");
			lpAddIn->ulMenu = NULL;
			lpAddIn->lpMenu = NULL;
		}
		if (lpAddIn->ulMenu && lpAddIn->lpMenu)
		{
			ULONG ulMenu = 0;
			for (ulMenu = 0; ulMenu < lpAddIn->ulMenu; ulMenu++)
			{
				// Save off our add-in struct
				lpAddIn->lpMenu[ulMenu].lpAddIn = lpAddIn;
				if (lpAddIn->lpMenu[ulMenu].szMenu)
					DebugPrint(DBGAddInPlumbing, L"Menu: %ws\n", lpAddIn->lpMenu[ulMenu].szMenu);
				if (lpAddIn->lpMenu[ulMenu].szHelp)
					DebugPrint(DBGAddInPlumbing, L"Help: %ws\n", lpAddIn->lpMenu[ulMenu].szHelp);
				DebugPrint(DBGAddInPlumbing, L"ID: 0x%08X\n", lpAddIn->lpMenu[ulMenu].ulID);
				DebugPrint(DBGAddInPlumbing, L"Context: 0x%08X\n", lpAddIn->lpMenu[ulMenu].ulContext);
				DebugPrint(DBGAddInPlumbing, L"Flags: 0x%08X\n", lpAddIn->lpMenu[ulMenu].ulFlags);
			}
		}
	}

	hRes = S_OK;
	lpAddIn->bLegacyPropTags = false;
	LPGETPROPTAGSV2 pfnGetPropTagsV2 = NULL;
	WC_D(pfnGetPropTagsV2, (LPGETPROPTAGSV2)GetProcAddress(hMod, szGetPropTagsV2));
	if (pfnGetPropTagsV2)
	{
		pfnGetPropTagsV2(&lpAddIn->ulPropTags, &lpAddIn->lpPropTags);
	}
	else
	{
		LoadLegacyPropTags(hMod, &lpAddIn->ulPropTags, &lpAddIn->lpPropTags);
		lpAddIn->bLegacyPropTags = true;
	}

	hRes = S_OK;
	LPGETPROPTYPES pfnGetPropTypes = NULL;
	WC_D(pfnGetPropTypes, (LPGETPROPTYPES)GetProcAddress(hMod, szGetPropTypes));
	if (pfnGetPropTypes)
	{
		pfnGetPropTypes(&lpAddIn->ulPropTypes, &lpAddIn->lpPropTypes);
	}

	hRes = S_OK;
	LPGETPROPGUIDS pfnGetPropGuids = NULL;
	WC_D(pfnGetPropGuids, (LPGETPROPGUIDS)GetProcAddress(hMod, szGetPropGuids));
	if (pfnGetPropGuids)
	{
		pfnGetPropGuids(&lpAddIn->ulPropGuids, &lpAddIn->lpPropGuids);
	}

	// v2 changed the LPNAMEID_ARRAY_ENTRY structure
	if (MFCMAPI_HEADER_V2 <= ulVersion)
	{
		hRes = S_OK;
		LPGETNAMEIDS pfnGetNameIDs = NULL;
		WC_D(pfnGetNameIDs, (LPGETNAMEIDS)GetProcAddress(hMod, szGetNameIDs));
		if (pfnGetNameIDs)
		{
			pfnGetNameIDs(&lpAddIn->ulNameIDs, &lpAddIn->lpNameIDs);
		}
	}

	hRes = S_OK;
	LPGETPROPFLAGS pfnGetPropFlags = NULL;
	WC_D(pfnGetPropFlags, (LPGETPROPFLAGS)GetProcAddress(hMod, szGetPropFlags));
	if (pfnGetPropFlags)
	{
		pfnGetPropFlags(&lpAddIn->ulPropFlags, &lpAddIn->lpPropFlags);
	}

	hRes = S_OK;
	LPGETSMARTVIEWPARSERARRAY pfnGetSmartViewParserArray = NULL;
	WC_D(pfnGetSmartViewParserArray, (LPGETSMARTVIEWPARSERARRAY)GetProcAddress(hMod, szGetSmartViewParserArray));
	if (pfnGetSmartViewParserArray)
	{
		pfnGetSmartViewParserArray(&lpAddIn->ulSmartViewParsers, &lpAddIn->lpSmartViewParsers);
	}

	hRes = S_OK;
	LPGETSMARTVIEWPARSERTYPEARRAY pfnGetSmartViewParserTypeArray = NULL;
	WC_D(pfnGetSmartViewParserTypeArray, (LPGETSMARTVIEWPARSERTYPEARRAY)GetProcAddress(hMod, szGetSmartViewParserTypeArray));
	if (pfnGetSmartViewParserTypeArray)
	{
		pfnGetSmartViewParserTypeArray(&lpAddIn->ulSmartViewParserTypes, &lpAddIn->lpSmartViewParserTypes);
	}

	DebugPrint(DBGAddInPlumbing, L"Done loading AddIn\n");
}

class CFileList
{
	struct DLLEntry
	{
		TCHAR*	szDLL;
		DLLEntry* lpNext;
	};
public:
	CFileList(_In_z_ LPCTSTR szKey);
	~CFileList();
	void AddToList(_In_z_ LPTSTR szDLL);
	bool IsOnList(_In_z_ LPTSTR szDLL);

	LPCTSTR m_szKey;
	HKEY m_hRootKey;
	DLLEntry* m_lpList;
};
#define EXCLUSION_LIST _T("AddInExclusionList") // STRING_OK
#define INCLUSION_LIST _T("AddInInclusionList") // STRING_OK
#define SEPARATOR _T(";") // STRING_OK

// Read in registry and build a list of invalid add-in DLLs
CFileList::CFileList(_In_z_ LPCTSTR szKey)
{
	HRESULT hRes = S_OK;
	LPTSTR lpszReg = NULL;

	m_lpList = NULL;
	m_hRootKey = CreateRootKey();
	m_szKey = szKey;

	if (m_hRootKey)
	{
		DWORD dwKeyType = NULL;
		WC_H(HrGetRegistryValue(
			m_hRootKey,
			m_szKey,
			&dwKeyType,
			(LPVOID*)&lpszReg));
	}

	if (lpszReg)
	{
		LPTSTR szContext = NULL;
		LPTSTR szDLL = _tcstok_s(lpszReg, SEPARATOR, &szContext);
		while (szDLL)
		{
			AddToList(szDLL);
			szDLL = _tcstok_s(NULL, SEPARATOR, &szContext);
		}
	}
	delete[] lpszReg;
} // CFileList::CFileList

// Write the list back to registry
CFileList::~CFileList()
{
	HRESULT hRes = S_OK;
	DLLEntry* lpCur = m_lpList;
	LPTSTR szList = NULL;
	size_t cchList = 0;
	size_t cchDLL = 0;

	while (lpCur)
	{
		if (lpCur->szDLL)
		{
			hRes = StringCchLength(lpCur->szDLL, STRSAFE_MAX_CCH, &cchDLL);

			if (SUCCEEDED(hRes) && cchDLL)
			{
				cchList += cchDLL + 1; // length + ; or null terminator
			}
		}
		lpCur = lpCur->lpNext;
	}

	if (cchList)
	{
		szList = new TCHAR[cchList];
	}

	if (szList)
	{
		szList[0] = NULL;
		lpCur = m_lpList;
		while (lpCur)
		{
			if (lpCur->szDLL)
			{
				hRes = StringCchCat(szList, cchList, lpCur->szDLL);
				if (lpCur->lpNext)
				{
					hRes = StringCchCat(szList, cchList, SEPARATOR);
				}
			}
			lpCur = lpCur->lpNext;
		}
		WriteStringToRegistry(
			m_hRootKey,
			m_szKey,
			szList);

		delete[] szList;
	}

	while (m_lpList)
	{
		delete[] m_lpList->szDLL;
		DLLEntry* lpNext = m_lpList->lpNext;
		delete m_lpList;
		m_lpList = lpNext;
	}
	EC_W32(RegCloseKey(m_hRootKey));
} // CFileList::~CFileList

// Add the DLL to the list
void CFileList::AddToList(_In_z_ LPTSTR szDLL)
{
	if (!szDLL) return;

	DLLEntry* lpNewEntry = new DLLEntry;
	if (lpNewEntry)
	{
		lpNewEntry->szDLL = NULL;
		lpNewEntry->lpNext = m_lpList;
		m_lpList = lpNewEntry;

		size_t cchDLL = NULL;
		HRESULT hRes = S_OK;
		hRes = StringCchLength(szDLL, STRSAFE_MAX_CCH, &cchDLL);
		if (SUCCEEDED(hRes) && cchDLL)
		{
			lpNewEntry->szDLL = new TCHAR[cchDLL + 1];

			if (lpNewEntry->szDLL)
			{
				hRes = StringCchCopy(lpNewEntry->szDLL, cchDLL + 1, szDLL);
			}
		}
	}
} // CFileList::AddToList

// Check this DLL name against the list
bool CFileList::IsOnList(_In_z_ LPTSTR szDLL)
{
	if (!szDLL) return true;
	DLLEntry* lpCur = m_lpList;
	while (lpCur)
	{
		if (!_tcsicmp(szDLL, lpCur->szDLL)) return true;
		lpCur = lpCur->lpNext;
	}
	return false;
} // CFileList::IsOnList

void LoadAddIns()
{
	DebugPrint(DBGAddInPlumbing, L"Loading AddIns\n");
	// First, we look at each DLL in the current dir and see if it exports 'LoadAddIn'
	LPADDIN lpCurAddIn = NULL;
	// Allocate space to hold information on all DLLs in the directory

	if (!RegKeys[regkeyLOADADDINS].ulCurDWORD)
	{
		DebugPrint(DBGAddInPlumbing, L"Bypassing add-in loading\n");
	}
	else
	{
		HRESULT hRes = S_OK;
		TCHAR szFilePath[MAX_PATH];
		DWORD dwDir = NULL;
		EC_D(dwDir, GetModuleFileName(NULL, szFilePath, _countof(szFilePath)));
		if (!dwDir) return;

		if (szFilePath[0])
		{
			// We got the path to mfcmapi.exe - need to strip it
			if (SUCCEEDED(hRes) && szFilePath[0] != _T('\0'))
			{
				LPTSTR lpszSlash = NULL;
				LPTSTR lpszCur = szFilePath;

				for (lpszSlash = lpszCur; *lpszCur; lpszCur = lpszCur++)
				{
					if (*lpszCur == _T('\\')) lpszSlash = lpszCur;
				}
				*lpszSlash = _T('\0');
			}
		}

		CFileList ExclusionList(EXCLUSION_LIST);
		CFileList InclusionList(INCLUSION_LIST);

#define SPECLEN 6 // for '\\*.dll'
		if (szFilePath[0])
		{
			DebugPrint(DBGAddInPlumbing, L"Current dir = \"%ws\"\n", LPCTSTRToWstring(szFilePath).c_str());
			hRes = StringCchCatN(szFilePath, MAX_PATH, _T("\\*.dll"), SPECLEN); // STRING_OK
			if (SUCCEEDED(hRes))
			{
				DebugPrint(DBGAddInPlumbing, L"File spec = \"%ws\"\n", LPCTSTRToWstring(szFilePath).c_str());

				WIN32_FIND_DATA FindFileData = { 0 };
				HANDLE hFind = FindFirstFile(szFilePath, &FindFileData);

				if (hFind == INVALID_HANDLE_VALUE)
				{
					DebugPrint(DBGAddInPlumbing, L"Invalid file handle. Error is %u.\n", GetLastError());
				}
				else
				{
					for (;;)
					{
						hRes = S_OK;
						DebugPrint(DBGAddInPlumbing, L"Examining \"%ws\"\n", LPCTSTRToWstring(FindFileData.cFileName).c_str());
						HMODULE hMod = NULL;

						// If we know the Add-in is good, just load it.
						// If we know it's bad, skip it.
						// Otherwise, we have to check if it's good.
						// LoadLibrary calls DLLMain, which can get expensive just to see if a function is exported
						// So we use DONT_RESOLVE_DLL_REFERENCES to see if we're interested in the DLL first
						// Only if we're interested do we reload the DLL for real
						if (InclusionList.IsOnList(FindFileData.cFileName))
						{
							hMod = MyLoadLibrary(FindFileData.cFileName);
						}
						else
						{
							if (!ExclusionList.IsOnList(FindFileData.cFileName))
							{
								hMod = LoadLibraryEx(FindFileData.cFileName, NULL, DONT_RESOLVE_DLL_REFERENCES);
								if (hMod)
								{
									LPLOADADDIN pfnLoadAddIn = NULL;
									WC_D(pfnLoadAddIn, (LPLOADADDIN)GetProcAddress(hMod, szLoadAddIn));
									FreeLibrary(hMod);
									hMod = NULL;

									if (pfnLoadAddIn)
									{
										// Remember this as a good add-in
										InclusionList.AddToList(FindFileData.cFileName);
										// We found a candidate, load it for real now
										hMod = MyLoadLibrary(FindFileData.cFileName);
										// GetProcAddress again just in case we loaded at a different address
										WC_D(pfnLoadAddIn, (LPLOADADDIN)GetProcAddress(hMod, szLoadAddIn));
									}
								}
								// If we still don't have a DLL loaded, exclude it
								if (!hMod)
								{
									ExclusionList.AddToList(FindFileData.cFileName);
								}
							}
						}
						if (hMod)
						{
							DebugPrint(DBGAddInPlumbing, L"Opened module\n");
							LPLOADADDIN pfnLoadAddIn = NULL;
							WC_D(pfnLoadAddIn, (LPLOADADDIN)GetProcAddress(hMod, szLoadAddIn));

							if (pfnLoadAddIn && GetAddinVersion(hMod) == MFCMAPI_HEADER_CURRENT_VERSION)
							{
								DebugPrint(DBGAddInPlumbing, L"Found an add-in\n");
								// Add a node
								if (!lpCurAddIn)
								{
									if (!g_lpMyAddins)
									{
										g_lpMyAddins = new _AddIn;
										ZeroMemory(g_lpMyAddins, sizeof(_AddIn));
									}
									lpCurAddIn = g_lpMyAddins;
								}
								else if (lpCurAddIn)
								{
									lpCurAddIn->lpNextAddIn = new _AddIn;
									ZeroMemory(lpCurAddIn->lpNextAddIn, sizeof(_AddIn));
									lpCurAddIn = lpCurAddIn->lpNextAddIn;
								}

								// Now that we have a node, populate it
								if (lpCurAddIn)
								{
									LoadSingleAddIn(lpCurAddIn, hMod, pfnLoadAddIn);
								}
							}
							else
							{
								// Not an add-in, release the hMod
								FreeLibrary(hMod);
							}
						}
						// get next file
						if (!FindNextFile(hFind, &FindFileData)) break;
					}

					DWORD dwRet = GetLastError();
					FindClose(hFind);
					if (dwRet != ERROR_NO_MORE_FILES)
					{
						DebugPrint(DBGAddInPlumbing, L"FindNextFile error. Error is %u.\n", dwRet);
					}
				}
			}
		}
	}

	MergeAddInArrays();
	DebugPrint(DBGAddInPlumbing, L"Done loading AddIns\n");
} // LoadAddIns

void ResetArrays()
{
	if (PropTypeArray != g_PropTypeArray) delete[] PropTypeArray;
	if (PropTagArray != g_PropTagArray) delete[] PropTagArray;
	if (PropGuidArray != g_PropGuidArray) delete[] PropGuidArray;
	if (NameIDArray != g_NameIDArray) delete[] NameIDArray;
	if (FlagArray != g_FlagArray) delete[] FlagArray;
	if (SmartViewParserArray != g_SmartViewParserArray) delete[] SmartViewParserArray;
	if (SmartViewParserTypeArray != g_SmartViewParserTypeArray) delete[] SmartViewParserTypeArray;

	ulPropTypeArray = g_ulPropTypeArray;
	PropTypeArray = g_PropTypeArray;
	ulPropTagArray = g_ulPropTagArray;
	PropTagArray = g_PropTagArray;
	ulPropGuidArray = g_ulPropGuidArray;
	PropGuidArray = g_PropGuidArray;
	ulNameIDArray = g_ulNameIDArray;
	NameIDArray = g_NameIDArray;
	ulFlagArray = g_ulFlagArray;
	FlagArray = g_FlagArray;
	ulSmartViewParserArray = g_ulSmartViewParserArray;
	SmartViewParserArray = g_SmartViewParserArray;
	ulSmartViewParserTypeArray = g_ulSmartViewParserTypeArray;
	SmartViewParserTypeArray = g_SmartViewParserTypeArray;
}

void UnloadAddIns()
{
	DebugPrint(DBGAddInPlumbing, L"Unloading AddIns\n");
	if (g_lpMyAddins)
	{
		LPADDIN lpCurAddIn = g_lpMyAddins;
		while (lpCurAddIn)
		{
			DebugPrint(DBGAddInPlumbing, L"Freeing add-in\n");
			if (lpCurAddIn->bLegacyPropTags)
			{
				delete[] lpCurAddIn->lpPropTags;
			}
			if (lpCurAddIn->hMod)
			{
				HRESULT hRes = S_OK;
				if (lpCurAddIn->szName)
				{
					DebugPrint(DBGAddInPlumbing, L"Unloading \"%ws\"\n", lpCurAddIn->szName);
				}
				LPUNLOADADDIN pfnUnLoadAddIn = NULL;
				WC_D(pfnUnLoadAddIn, (LPUNLOADADDIN)GetProcAddress(lpCurAddIn->hMod, szUnloadAddIn));
				if (pfnUnLoadAddIn) pfnUnLoadAddIn();

				FreeLibrary(lpCurAddIn->hMod);
			}
			LPADDIN lpAddInToFree = lpCurAddIn;
			lpCurAddIn = lpCurAddIn->lpNextAddIn;
			delete lpAddInToFree;
		}
	}

	ResetArrays();

	DebugPrint(DBGAddInPlumbing, L"Done unloading AddIns\n");
} // UnloadAddIns

#ifndef MRMAPI
// Adds menu items appropriate to the context
// Returns number of menu items added
_Check_return_ ULONG ExtendAddInMenu(HMENU hMenu, ULONG ulAddInContext)
{
	DebugPrint(DBGAddInPlumbing, L"Extending menus, ulAddInContext = 0x%08X\n", ulAddInContext);
	HMENU hAddInMenu = NULL;

	UINT uidCurMenu = ID_ADDINMENU;

	if (MENU_CONTEXT_PROPERTY == ulAddInContext)
	{
		uidCurMenu = ID_ADDINPROPERTYMENU;
	}

	if (g_lpMyAddins)
	{
		LPADDIN lpCurAddIn = g_lpMyAddins;
		while (lpCurAddIn)
		{
			DebugPrint(DBGAddInPlumbing, L"Examining add-in for menus\n");
			if (lpCurAddIn->hMod)
			{
				HRESULT hRes = S_OK;
				if (lpCurAddIn->szName)
				{
					DebugPrint(DBGAddInPlumbing, L"Examining \"%ws\"\n", lpCurAddIn->szName);
				}
				ULONG ulMenu = 0;
				for (ulMenu = 0; ulMenu < lpCurAddIn->ulMenu && SUCCEEDED(hRes); ulMenu++)
				{
					if ((lpCurAddIn->lpMenu[ulMenu].ulFlags & MENU_FLAGS_SINGLESELECT) &&
						(lpCurAddIn->lpMenu[ulMenu].ulFlags & MENU_FLAGS_MULTISELECT))
					{
						// Invalid combo of flags - don't add the menu
						DebugPrint(DBGAddInPlumbing, L"Invalid flags on menu \"%ws\" in add-in \"%ws\"\n", lpCurAddIn->lpMenu[ulMenu].szMenu, lpCurAddIn->szName);
						DebugPrint(DBGAddInPlumbing, L"MENU_FLAGS_SINGLESELECT and MENU_FLAGS_MULTISELECT cannot be combined\n");
						continue;
					}
					if (lpCurAddIn->lpMenu[ulMenu].ulContext & ulAddInContext)
					{
						// Add the Add-Ins menu if we haven't added it already
						if (!hAddInMenu)
						{
							WCHAR szAddInTitle[8] = { 0 }; // The length of IDS_ADDINSMENU
							int iRet = NULL;
							EC_D(iRet, LoadStringW(GetModuleHandle(NULL),
								IDS_ADDINSMENU,
								szAddInTitle,
								_countof(szAddInTitle)));

							hAddInMenu = CreatePopupMenu();
							::InsertMenuW(
								hMenu,
								(UINT)-1,
								MF_BYPOSITION | MF_POPUP | MF_ENABLED,
								(UINT_PTR)hAddInMenu,
								(LPCWSTR)szAddInTitle);
						}

						// Now add each of the menu entries
						if (SUCCEEDED(hRes))
						{
							LPMENUENTRY lpMenu = CreateMenuEntry(lpCurAddIn->lpMenu[ulMenu].szMenu);
							if (lpMenu)
							{
								lpMenu->m_AddInData = (ULONG_PTR)&lpCurAddIn->lpMenu[ulMenu];
							}

							EC_B(AppendMenu(
								hAddInMenu,
								MF_ENABLED | MF_OWNERDRAW,
								uidCurMenu,
								(LPCTSTR)lpMenu));
							uidCurMenu++;
						}
					}
				}
			}
			lpCurAddIn = lpCurAddIn->lpNextAddIn;
		}
	}
	DebugPrint(DBGAddInPlumbing, L"Done extending menus\n");
	return uidCurMenu - ID_ADDINMENU;
} // ExtendAddInMenu

_Check_return_ LPMENUITEM GetAddinMenuItem(HWND hWnd, UINT uidMsg)
{
	if (uidMsg < ID_ADDINMENU) return NULL;

	MENUITEMINFOW subMenu = { 0 };
	subMenu.cbSize = sizeof(MENUITEMINFO);
	subMenu.fMask = MIIM_STATE | MIIM_ID | MIIM_DATA;

	if (::GetMenuItemInfoW(
		::GetMenu(hWnd),
		uidMsg,
		false,
		&subMenu) &&
		subMenu.dwItemData)
	{
		return (LPMENUITEM)((LPMENUENTRY)subMenu.dwItemData)->m_AddInData;
	}

	return NULL;
} // GetAddinMenuItem

void InvokeAddInMenu(_In_opt_ LPADDINMENUPARAMS lpParams)
{
	HRESULT hRes = S_OK;

	if (!lpParams) return;
	if (!lpParams->lpAddInMenu) return;
	if (!lpParams->lpAddInMenu->lpAddIn) return;

	if (!lpParams->lpAddInMenu->lpAddIn->pfnCallMenu)
	{
		if (!lpParams->lpAddInMenu->lpAddIn->hMod) return;
		WC_D(lpParams->lpAddInMenu->lpAddIn->pfnCallMenu, (LPCALLMENU)GetProcAddress(
			lpParams->lpAddInMenu->lpAddIn->hMod,
			szCallMenu));
	}

	if (!lpParams->lpAddInMenu->lpAddIn->pfnCallMenu)
	{
		DebugPrint(DBGAddInPlumbing, L"InvokeAddInMenu: CallMenu not found\n");
		return;
	}

	WC_H(lpParams->lpAddInMenu->lpAddIn->pfnCallMenu(lpParams));
} // InvokeAddInMenu
#endif // MRMAPI

// Compare type arrays.
int _cdecl CompareTypes(_In_ const void* a1, _In_ const void* a2)
{
	LPNAME_ARRAY_ENTRY lpType1 = (LPNAME_ARRAY_ENTRY)a1;
	LPNAME_ARRAY_ENTRY lpType2 = (LPNAME_ARRAY_ENTRY)a2;

	if (lpType1->ulValue > lpType2->ulValue) return 1;
	if (lpType1->ulValue == lpType2->ulValue)
	{
		return wcscmp(lpType1->lpszName, lpType2->lpszName);
	}
	return -1;
} // CompareTypes

// Compare tag arrays. Pay no attention to sort order - we'll sort on sort order during output.
int _cdecl CompareTags(_In_ const void* a1, _In_ const void* a2)
{
	LPNAME_ARRAY_ENTRY_V2 lpTag1 = (LPNAME_ARRAY_ENTRY_V2)a1;
	LPNAME_ARRAY_ENTRY_V2 lpTag2 = (LPNAME_ARRAY_ENTRY_V2)a2;

	if (lpTag1->ulValue > lpTag2->ulValue) return 1;
	if (lpTag1->ulValue == lpTag2->ulValue)
	{
		return wcscmp(lpTag1->lpszName, lpTag2->lpszName);
	}
	return -1;
} // CompareTags

int _cdecl CompareNameID(_In_ const void* a1, _In_ const void* a2)
{
	LPNAMEID_ARRAY_ENTRY lpID1 = (LPNAMEID_ARRAY_ENTRY)a1;
	LPNAMEID_ARRAY_ENTRY lpID2 = (LPNAMEID_ARRAY_ENTRY)a2;

	if (lpID1->lValue > lpID2->lValue) return 1;
	if (lpID1->lValue == lpID2->lValue)
	{
		int iCmp = wcscmp(lpID1->lpszName, lpID2->lpszName);
		if (iCmp) return iCmp;
		if (IsEqualGUID(*lpID1->lpGuid, *lpID2->lpGuid)) return 0;
	}
	return -1;
} // CompareNameID

int _cdecl CompareSmartViewParser(_In_ const void* a1, _In_ const void* a2)
{
	LPSMARTVIEW_PARSER_ARRAY_ENTRY lpParser1 = (LPSMARTVIEW_PARSER_ARRAY_ENTRY)a1;
	LPSMARTVIEW_PARSER_ARRAY_ENTRY lpParser2 = (LPSMARTVIEW_PARSER_ARRAY_ENTRY)a2;

	if (lpParser1->ulIndex > lpParser2->ulIndex) return 1;
	if (lpParser1->ulIndex == lpParser2->ulIndex)
	{
		if (lpParser1->iStructType > lpParser2->iStructType) return 1;
		if (lpParser1->iStructType == lpParser2->iStructType)
		{
			if (lpParser1->bMV && !lpParser2->bMV) return 1;
			if (!lpParser1->bMV && lpParser2->bMV) return -1;
			return 0;
		}
	}
	return -1;
} // CompareSmartViewParser

void MergeArrays(
	_Inout_bytecap_x_(cIn1 * width) LPVOID In1,
	_In_ size_t cIn1,
	_Inout_bytecap_x_(cIn2 * width) LPVOID In2,
	_In_ size_t cIn2,
	_Out_ _Deref_post_bytecap_x_(*lpcOut * width) LPVOID* lpOut,
	_Out_ size_t* lpcOut,
	_In_ size_t width,
	_In_ int(_cdecl *comp)(const void *, const void *))
{
	if (!In1 && !In2) return;
	if (!lpOut || !lpcOut) return;

	// Assume no duplicates
	*lpcOut = cIn1 + cIn2;
	*lpOut = new char[*lpcOut * width];

	if (*lpOut)
	{
		memset(*lpOut, 0, *lpcOut * width);
		char* iIn1 = (char*)In1;
		char* iIn2 = (char*)In2;
		LPVOID endIn1 = iIn1 + width * (cIn1 - 1);
		LPVOID endIn2 = iIn2 + width * (cIn2 - 1);
		char* iOut = (char*)*lpOut;
		int iComp = 0;

		while (iIn1 <= endIn1 && iIn2 <= endIn2)
		{
			iComp = comp(iIn1, iIn2);
			if (iComp < 0)
			{
				memcpy(iOut, iIn1, width);
				iIn1 += width;
			}
			else if (iComp > 0)
			{
				memcpy(iOut, iIn2, width);
				iIn2 += width;
			}
			else
			{
				// They're the same - copy one over and skip past both
				memcpy(iOut, iIn1, width);
				iIn1 += width;
				iIn2 += width;
			}
			iOut += width;
		}
		while (iIn1 <= endIn1)
		{
			memcpy(iOut, iIn1, width);
			iIn1 += width;
			iOut += width;
		}
		while (iIn2 <= endIn2)
		{
			memcpy(iOut, iIn2, width);
			iIn2 += width;
			iOut += width;
		}

		*lpcOut = (iOut - (char*)*lpOut) / width;
	}
} // MergeArrays

// Flags are difficult to sort since we need to have a stable sort
// Records with the same key must appear in the output in the same order as the input
// qsort doesn't guarantee this, so we do it manually with an insertion sort
void SortFlagArray(_In_count_(ulFlags) LPFLAG_ARRAY_ENTRY lpFlags, _In_ ULONG ulFlags)
{
	ULONG i = 0;
	ULONG iLoc = 0;
	for (i = 1; i < ulFlags; i++)
	{
		FLAG_ARRAY_ENTRY NextItem = lpFlags[i];
		for (iLoc = i; iLoc > 0; iLoc--)
		{
			if (lpFlags[iLoc - 1].ulFlagName <= NextItem.ulFlagName) break;
			lpFlags[iLoc] = lpFlags[iLoc - 1];
		}
		lpFlags[iLoc] = NextItem;
	}
} // SortFlagArray

// Consults the end of the supplied array to find a match to the passed in entry
// If no dupe is found, copies lpSource[*lpiSource] to lpTarget and increases *lpcArray
// Increases *lpiSource regardless
void AppendFlagIfNotDupe(_In_count_(*lpcArray) LPFLAG_ARRAY_ENTRY lpTarget, _In_ size_t* lpcArray, _In_count_(*lpiSource + 1)  LPFLAG_ARRAY_ENTRY lpSource, _In_ size_t* lpiSource)
{
	size_t iTarget = *lpcArray;
	size_t iSource = *lpiSource;
	(*lpiSource)++;
	while (iTarget)
	{
		iTarget--;
		// Stop searching when ulFlagName doesn't match
		// Assumes lpTarget is sorted
		if (lpTarget[iTarget].ulFlagName != lpSource[iSource].ulFlagName) break;
		if (lpTarget[iTarget].lFlagValue == lpSource[iSource].lFlagValue &&
			lpTarget[iTarget].ulFlagType == lpSource[iSource].ulFlagType &&
			!wcscmp(lpTarget[iTarget].lpszName, lpSource[iSource].lpszName))
		{
			return;
		}
	}
	lpTarget[*lpcArray] = lpSource[iSource];
	(*lpcArray)++;
} // AppendFlagIfNotDupe

// Similar to MergeArrays, but using AppendFlagIfNotDupe logic
void MergeFlagArrays(
	_In_count_(cIn1) LPFLAG_ARRAY_ENTRY In1,
	_In_ size_t cIn1,
	_In_count_(cIn2) LPFLAG_ARRAY_ENTRY In2,
	_In_ size_t cIn2,
	_Out_ _Deref_post_count_(*lpcOut) LPFLAG_ARRAY_ENTRY* lpOut,
	_Out_ size_t* lpcOut)
{
	if (!In1 && !In2) return;
	if (!lpOut || !lpcOut) return;

	// Assume no duplicates
	*lpcOut = cIn1 + cIn2;
	LPFLAG_ARRAY_ENTRY Out = new FLAG_ARRAY_ENTRY[*lpcOut];

	if (Out)
	{
		size_t iIn1 = 0;
		size_t iIn2 = 0;
		size_t iOut = 0;

		while (iIn1 < cIn1 && iIn2 < cIn2)
		{
			// Add from In1 first, then In2 when In2 is bigger than In2
			if (In1[iIn1].ulFlagName <= In2[iIn2].ulFlagName)
			{
				AppendFlagIfNotDupe(Out, &iOut, In1, &iIn1);
			}
			else
			{
				AppendFlagIfNotDupe(Out, &iOut, In2, &iIn2);
			}
		}
		while (iIn1 < cIn1)
		{
			AppendFlagIfNotDupe(Out, &iOut, In1, &iIn1);
		}
		while (iIn2 < cIn2)
		{
			AppendFlagIfNotDupe(Out, &iOut, In2, &iIn2);
		}

		*lpcOut = iOut;
		*lpOut = Out;
	}
} // MergeFlagArrays

// Assumes built in arrays are already sorted!
void MergeAddInArrays()
{
	DebugPrint(DBGAddInPlumbing, L"Merging Add-In arrays\n");

	ResetArrays();

	DebugPrint(DBGAddInPlumbing, L"Found 0x%08X built in prop tags.\n", g_ulPropTagArray);
	DebugPrint(DBGAddInPlumbing, L"Found 0x%08X built in prop types.\n", g_ulPropTypeArray);
	DebugPrint(DBGAddInPlumbing, L"Found 0x%08X built in guids.\n", g_ulPropGuidArray);
	DebugPrint(DBGAddInPlumbing, L"Found 0x%08X built in named ids.\n", g_ulNameIDArray);
	DebugPrint(DBGAddInPlumbing, L"Found 0x%08X built in flags.\n", g_ulFlagArray);
	DebugPrint(DBGAddInPlumbing, L"Found 0x%08X built in Smart View parsers.\n", g_ulSmartViewParserArray);
	DebugPrint(DBGAddInPlumbing, L"Found 0x%08X built in Smart View parser types.\n", g_ulSmartViewParserTypeArray);

	// No add-in == nothing to merge
	if (!g_lpMyAddins) return;

	// First pass - count up any the size of the guid array
	ULONG ulAddInPropGuidArray = g_ulPropGuidArray;

	LPADDIN lpCurAddIn = g_lpMyAddins;
	while (lpCurAddIn)
	{
		DebugPrint(DBGAddInPlumbing, L"Looking at %ws\n", lpCurAddIn->szName);
		DebugPrint(DBGAddInPlumbing, L"Found 0x%08X prop tags.\n", lpCurAddIn->ulPropTags);
		DebugPrint(DBGAddInPlumbing, L"Found 0x%08X prop types.\n", lpCurAddIn->ulPropTypes);
		DebugPrint(DBGAddInPlumbing, L"Found 0x%08X guids.\n", lpCurAddIn->ulPropGuids);
		DebugPrint(DBGAddInPlumbing, L"Found 0x%08X named ids.\n", lpCurAddIn->ulNameIDs);
		DebugPrint(DBGAddInPlumbing, L"Found 0x%08X flags.\n", lpCurAddIn->ulPropFlags);
		DebugPrint(DBGAddInPlumbing, L"Found 0x%08X Smart View parsers.\n", lpCurAddIn->ulSmartViewParsers);
		DebugPrint(DBGAddInPlumbing, L"Found 0x%08X Smart View parser types.\n", lpCurAddIn->ulSmartViewParserTypes);
		ulAddInPropGuidArray += lpCurAddIn->ulPropGuids;
		lpCurAddIn = lpCurAddIn->lpNextAddIn;
	}

	// if count has gone up - we need to merge
	// Allocate a larger array and initialize it with our hardcoded data
	ULONG i = 0;
	if (ulAddInPropGuidArray != g_ulPropGuidArray)
	{
		PropGuidArray = new GUID_ARRAY_ENTRY[ulAddInPropGuidArray];
		if (PropGuidArray)
		{
			for (i = 0; i < g_ulPropGuidArray; i++)
			{
				PropGuidArray[i] = g_PropGuidArray[i];
			}
		}
	}

	// Second pass - merge our arrays to the hardcoded arrays
	ULONG ulCurPropGuid = g_ulPropGuidArray;
	lpCurAddIn = g_lpMyAddins;
	while (lpCurAddIn)
	{
		if (lpCurAddIn->ulPropTypes)
		{
			qsort(lpCurAddIn->lpPropTypes, lpCurAddIn->ulPropTypes, sizeof(NAME_ARRAY_ENTRY), &CompareTypes);
			LPNAME_ARRAY_ENTRY newPropTypeArray = NULL;
			size_t ulnewPropTypeArray = NULL;
			MergeArrays(PropTypeArray, ulPropTypeArray,
				lpCurAddIn->lpPropTypes, lpCurAddIn->ulPropTypes,
				(LPVOID*)&newPropTypeArray, &ulnewPropTypeArray,
				sizeof(NAME_ARRAY_ENTRY),
				CompareTypes);
			if (PropTypeArray != g_PropTypeArray) delete[] PropTypeArray;
			PropTypeArray = newPropTypeArray;
			ulPropTypeArray = (ULONG)ulnewPropTypeArray;
		}

		if (lpCurAddIn->ulPropTags)
		{
			qsort(lpCurAddIn->lpPropTags, lpCurAddIn->ulPropTags, sizeof(NAME_ARRAY_ENTRY_V2), &CompareTags);
			LPNAME_ARRAY_ENTRY_V2 newPropTagArray = NULL;
			size_t ulnewPropTagArray = NULL;
			MergeArrays(PropTagArray, ulPropTagArray,
				lpCurAddIn->lpPropTags, lpCurAddIn->ulPropTags,
				(LPVOID*)&newPropTagArray, &ulnewPropTagArray,
				sizeof(NAME_ARRAY_ENTRY_V2),
				CompareTags);
			if (PropTagArray != g_PropTagArray) delete[] PropTagArray;
			PropTagArray = newPropTagArray;
			ulPropTagArray = (ULONG)ulnewPropTagArray;
		}

		if (lpCurAddIn->ulNameIDs)
		{
			qsort(lpCurAddIn->lpNameIDs, lpCurAddIn->ulNameIDs, sizeof(NAMEID_ARRAY_ENTRY), &CompareNameID);
			LPNAMEID_ARRAY_ENTRY newNameIDArray = NULL;
			size_t ulnewNameIDArray = NULL;
			MergeArrays(NameIDArray, ulNameIDArray,
				lpCurAddIn->lpNameIDs, lpCurAddIn->ulNameIDs,
				(LPVOID*)&newNameIDArray, &ulnewNameIDArray,
				sizeof(NAMEID_ARRAY_ENTRY),
				CompareNameID);
			if (NameIDArray != g_NameIDArray) delete[] NameIDArray;
			NameIDArray = newNameIDArray;
			ulNameIDArray = (ULONG)ulnewNameIDArray;
		}

		if (lpCurAddIn->ulPropFlags)
		{
			SortFlagArray(lpCurAddIn->lpPropFlags, lpCurAddIn->ulPropFlags);
			LPFLAG_ARRAY_ENTRY newFlagArray = NULL;
			size_t ulnewFlagArray = NULL;
			MergeFlagArrays(FlagArray, ulFlagArray,
				lpCurAddIn->lpPropFlags, lpCurAddIn->ulPropFlags,
				&newFlagArray, &ulnewFlagArray);
			if (FlagArray != g_FlagArray) delete[] FlagArray;
			FlagArray = newFlagArray;
			ulFlagArray = (ULONG)ulnewFlagArray;
		}

		if (lpCurAddIn->ulSmartViewParsers)
		{
			qsort(lpCurAddIn->lpSmartViewParsers, lpCurAddIn->ulSmartViewParsers, sizeof(SMARTVIEW_PARSER_ARRAY_ENTRY), &CompareSmartViewParser);
			LPSMARTVIEW_PARSER_ARRAY_ENTRY newSmartViewParserArray = NULL;
			size_t ulnewSmartViewParserArray = NULL;
			MergeArrays(SmartViewParserArray, ulSmartViewParserArray,
				lpCurAddIn->lpSmartViewParsers, lpCurAddIn->ulSmartViewParsers,
				(LPVOID*)&newSmartViewParserArray, &ulnewSmartViewParserArray,
				sizeof(SMARTVIEW_PARSER_ARRAY_ENTRY),
				CompareSmartViewParser);
			if (SmartViewParserArray != g_SmartViewParserArray) delete[] SmartViewParserArray;
			SmartViewParserArray = newSmartViewParserArray;
			ulSmartViewParserArray = (ULONG)ulnewSmartViewParserArray;
		}

		// We add our new parsers to the end of the array, assigning ids starting with IDS_STEND
		static ULONG s_ulNextParser = IDS_STEND;
		if (lpCurAddIn->ulSmartViewParserTypes)
		{
			ULONG ulNewArray = lpCurAddIn->ulSmartViewParserTypes + ulSmartViewParserTypeArray;
			LPNAME_ARRAY_ENTRY lpNewArray = new NAME_ARRAY_ENTRY[ulNewArray];

			if (lpNewArray)
			{
				for (i = 0; i < ulSmartViewParserTypeArray; i++)
				{
					lpNewArray[i] = SmartViewParserTypeArray[i];
				}

				for (i = 0; i < lpCurAddIn->ulSmartViewParserTypes; i++)
				{
					lpNewArray[i + ulSmartViewParserTypeArray].ulValue = s_ulNextParser++;
					lpNewArray[i + ulSmartViewParserTypeArray].lpszName = lpCurAddIn->lpSmartViewParserTypes[i];
				}
			}

			SmartViewParserTypeArray = lpNewArray;
			ulSmartViewParserTypeArray = ulNewArray;
		}

		if (lpCurAddIn->ulPropGuids)
		{
			// Copy guids from lpCurAddIn->lpPropGuids, checking for dupes on the way
			for (i = 0; i < lpCurAddIn->ulPropGuids; i++)
			{
				ULONG iCur = 0;
				bool bDupe = false;
				// Since this array isn't sorted, we have to compare against all valid entries for dupes
				for (iCur = 0; iCur < ulCurPropGuid; iCur++)
				{
					if (IsEqualGUID(*lpCurAddIn->lpPropGuids[i].lpGuid, *PropGuidArray[iCur].lpGuid))
					{
						bDupe = true;
						break;
					}
				}

				if (!bDupe)
				{
					PropGuidArray[ulCurPropGuid] = lpCurAddIn->lpPropGuids[i];
					ulCurPropGuid++;
				}
			}
			ulPropGuidArray = ulCurPropGuid;
		}
		lpCurAddIn = lpCurAddIn->lpNextAddIn;
	}

	DebugPrint(DBGAddInPlumbing, L"After merge, 0x%08X prop tags.\n", ulPropTagArray);
	DebugPrint(DBGAddInPlumbing, L"After merge, 0x%08X prop types.\n", ulPropTypeArray);
	DebugPrint(DBGAddInPlumbing, L"After merge, 0x%08X guids.\n", ulPropGuidArray);
	DebugPrint(DBGAddInPlumbing, L"After merge, 0x%08X flags.\n", ulFlagArray);
	DebugPrint(DBGAddInPlumbing, L"After merge, 0x%08X Smart View parsers.\n", ulSmartViewParserArray);
	DebugPrint(DBGAddInPlumbing, L"After merge, 0x%08X Smart View parser types.\n", ulSmartViewParserTypeArray);

	DebugPrint(DBGAddInPlumbing, L"Done merging add-in arrays\n");
}

__declspec(dllexport) void __cdecl AddInLog(bool bPrintThreadTime, _Printf_format_string_ LPWSTR szMsg, ...)
{
	if (!fIsSet(DBGAddIn)) return;

	va_list argList = NULL;
	va_start(argList, szMsg);
	wstring szAddInLogString = formatV(szMsg, argList);
	va_end(argList);

	_OutputW(DBGAddIn, NULL, bPrintThreadTime, szAddInLogString);
}

#ifndef MRMAPI
_Check_return_ __declspec(dllexport) HRESULT __cdecl SimpleDialog(_In_z_ LPWSTR szTitle, _Printf_format_string_ LPWSTR szMsg, ...)
{
	HRESULT hRes = S_OK;

	CEditor MySimpleDialog(
		NULL,
		NULL,
		NULL,
		0,
		CEDITOR_BUTTON_OK);
	MySimpleDialog.SetAddInTitle(szTitle);

	va_list argList = NULL;
	va_start(argList, szMsg);
	wstring szDialogString = formatV(szMsg, argList);
	va_end(argList);

	LPTSTR lpszDialogString = wstringToLPTSTR(szDialogString);
	MySimpleDialog.SetPromptPostFix(lpszDialogString);
	delete[] lpszDialogString;

	WC_H(MySimpleDialog.DisplayDialog());
	return hRes;
} // SimpleDialog

_Check_return_ __declspec(dllexport) HRESULT __cdecl ComplexDialog(_In_ LPADDINDIALOG lpDialog, _Out_ LPADDINDIALOGRESULT* lppDialogResult)
{
	if (!lpDialog) return MAPI_E_INVALID_PARAMETER;
	// Reject any flags except CEDITOR_BUTTON_OK and CEDITOR_BUTTON_CANCEL
	if (lpDialog->ulButtonFlags & ~(CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)) return MAPI_E_INVALID_PARAMETER;
	HRESULT hRes = S_OK;
	ULONG i = 0;

	CEditor MyComplexDialog(
		NULL,
		NULL,
		NULL,
		lpDialog->ulNumControls,
		lpDialog->ulButtonFlags);
	MyComplexDialog.SetAddInTitle(lpDialog->szTitle);

#ifdef UNICODE
	MyComplexDialog.SetPromptPostFix(lpDialog->szPrompt);
#else
	char *szAnsiPrompt = NULL;
	EC_H(UnicodeToAnsi(lpDialog->szPrompt, &szAnsiPrompt));
	MyComplexDialog.SetPromptPostFix(szAnsiPrompt);
	delete[] szAnsiPrompt;
#endif

	if (lpDialog->ulNumControls && lpDialog->lpDialogControls)
	{
		for (i = 0; i < lpDialog->ulNumControls; i++)
		{
			switch (lpDialog->lpDialogControls[i].cType)
			{
			case ADDIN_CTRL_CHECK:
				MyComplexDialog.InitPane(i, CreateCheckPane(
					NULL,
					lpDialog->lpDialogControls[i].bDefaultCheckState,
					lpDialog->lpDialogControls[i].bReadOnly));
				break;
			case ADDIN_CTRL_EDIT_TEXT:
			{
				LPTSTR szDefaultText = NULL;
#ifdef UNICODE
				szDefaultText = lpDialog->lpDialogControls[i].szDefaultText;
#else
				EC_H(UnicodeToAnsi(lpDialog->lpDialogControls[i].szDefaultText, &szDefaultText));
#endif
				if (lpDialog->lpDialogControls[i].bMultiLine)
				{
					MyComplexDialog.InitPane(i, CreateCollapsibleTextPane(
						NULL,
						lpDialog->lpDialogControls[i].bReadOnly));
					MyComplexDialog.SetString(i, szDefaultText);
				}
				else
				{
					MyComplexDialog.InitPane(i, CreateSingleLinePane(
						NULL,
						szDefaultText,
						lpDialog->lpDialogControls[i].bReadOnly));

				}
#ifndef UNICODE
				delete[] szDefaultText;
#endif
				break;
			}
			case ADDIN_CTRL_EDIT_BINARY:
				if (lpDialog->lpDialogControls[i].bMultiLine)
				{
					MyComplexDialog.InitPane(i, CreateCollapsibleTextPane(
						NULL,
						lpDialog->lpDialogControls[i].bReadOnly));
				}
				else
				{
					MyComplexDialog.InitPane(i, CreateSingleLinePane(
						NULL,
						NULL,
						lpDialog->lpDialogControls[i].bReadOnly));

				}
				MyComplexDialog.SetBinary(
					i,
					lpDialog->lpDialogControls[i].lpBin,
					lpDialog->lpDialogControls[i].cbBin);
				break;
			case ADDIN_CTRL_EDIT_NUM_DECIMAL:
				MyComplexDialog.InitPane(i, CreateSingleLinePane(
					NULL,
					NULL,
					lpDialog->lpDialogControls[i].bReadOnly));
				MyComplexDialog.SetDecimal(
					i,
					lpDialog->lpDialogControls[i].ulDefaultNum);
				break;
			case ADDIN_CTRL_EDIT_NUM_HEX:
				MyComplexDialog.InitPane(i, CreateSingleLinePane(
					NULL,
					NULL,
					lpDialog->lpDialogControls[i].bReadOnly));
				MyComplexDialog.SetHex(
					i,
					lpDialog->lpDialogControls[i].ulDefaultNum);
				break;
			}

			// Do this after initializing controls so we have our label status set correctly.
			MyComplexDialog.SetAddInLabel(i, lpDialog->lpDialogControls[i].szLabel);
		}
	}

	WC_H(MyComplexDialog.DisplayDialog());

	// Put together results if needed
	if (SUCCEEDED(hRes) && lppDialogResult && lpDialog->ulNumControls && lpDialog->lpDialogControls)
	{
		LPADDINDIALOGRESULT lpResults = new _AddInDialogResult;
		if (lpResults)
		{
			lpResults->ulNumControls = lpDialog->ulNumControls;
			lpResults->lpDialogControlResults = new _AddInDialogControlResult[lpDialog->ulNumControls];
			if (lpResults->lpDialogControlResults)
			{
				ZeroMemory(lpResults->lpDialogControlResults, sizeof(_AddInDialogControlResult)*lpDialog->ulNumControls);
				for (i = 0; i < lpDialog->ulNumControls; i++)
				{
					lpResults->lpDialogControlResults[i].cType = lpDialog->lpDialogControls[i].cType;
					switch (lpDialog->lpDialogControls[i].cType)
					{
					case ADDIN_CTRL_CHECK:
						lpResults->lpDialogControlResults[i].bCheckState = MyComplexDialog.GetCheck(i);
						break;
					case ADDIN_CTRL_EDIT_TEXT:
					{
						LPWSTR szText = MyComplexDialog.GetStringW(i);
						if (szText)
						{
							size_t cchText = NULL;
							WC_H(StringCchLengthW(szText, STRSAFE_MAX_CCH, &cchText));

							cchText++;
							lpResults->lpDialogControlResults[i].szText = new WCHAR[cchText];

							if (lpResults->lpDialogControlResults[i].szText)
							{
								EC_H(StringCchCopyW(
									lpResults->lpDialogControlResults[i].szText,
									cchText,
									szText));
							}
						}
						break;
					}
					case ADDIN_CTRL_EDIT_BINARY:
						// GetEntryID does just what we want - abuse it
						WC_H(MyComplexDialog.GetEntryID(
							i,
							false,
							&lpResults->lpDialogControlResults[i].cbBin,
							(LPENTRYID*)&lpResults->lpDialogControlResults[i].lpBin));
						break;
					case ADDIN_CTRL_EDIT_NUM_DECIMAL:
						lpResults->lpDialogControlResults[i].ulVal = MyComplexDialog.GetDecimal(i);
						break;
					case ADDIN_CTRL_EDIT_NUM_HEX:
						lpResults->lpDialogControlResults[i].ulVal = MyComplexDialog.GetHex(i);
						break;
					}
				}
			}
			if (SUCCEEDED(hRes))
			{
				*lppDialogResult = lpResults;
			}
			else FreeDialogResult(lpResults);

		}
	}
	return hRes;
} // ComplexDialog

__declspec(dllexport) void __cdecl FreeDialogResult(_In_ LPADDINDIALOGRESULT lpDialogResult)
{
	if (lpDialogResult)
	{
		if (lpDialogResult->lpDialogControlResults)
		{
			ULONG i = 0;
			for (i = 0; i < lpDialogResult->ulNumControls; i++)
			{
				delete[] lpDialogResult->lpDialogControlResults[i].lpBin;
				delete[] lpDialogResult->lpDialogControlResults[i].szText;
			}
		}
		delete[] lpDialogResult;
	}
} // FreeDialogResult
#endif

__declspec(dllexport) void __cdecl GetMAPIModule(_In_ HMODULE* lphModule, bool bForce)
{
	if (!lphModule) return;
	*lphModule = GetMAPIHandle();
	if (!*lphModule && bForce)
	{
		// No MAPI loaded - load it
		*lphModule = GetPrivateMAPI();
	}
} // GetMAPIModule

wstring AddInStructTypeToString(__ParsingTypeEnum iStructType)
{
	ULONG i = 0;
	for (i = 0; i < ulSmartViewParserTypeArray; i++)
	{
		if (SmartViewParserTypeArray[i].ulValue == (ULONG)iStructType)
		{
			return SmartViewParserTypeArray[i].lpszName;
		}
	}

	return L"";
}

wstring AddInSmartView(__ParsingTypeEnum iStructType, ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	// Don't let add-ins hijack our built in types
	if (iStructType <= IDS_STEND - 1) return L"";

	wstring szStructType = AddInStructTypeToString(iStructType);
	if (szStructType.empty()) return L"";

	LPADDIN lpCurAddIn = g_lpMyAddins;
	while (lpCurAddIn)
	{
		if (lpCurAddIn->ulSmartViewParserTypes)
		{
			ULONG i = 0;
			for (i = 0; i < lpCurAddIn->ulSmartViewParserTypes; i++)
			{
				if (0 == szStructType.compare(lpCurAddIn->lpSmartViewParserTypes[i]))
				{
					HRESULT hRes = S_OK;
					LPSMARTVIEWPARSE pfnSmartViewParse = NULL;
					WC_D(pfnSmartViewParse, (LPSMARTVIEWPARSE)GetProcAddress(lpCurAddIn->hMod, szSmartViewParse));

					LPFREEPARSE pfnFreeParse = NULL;
					WC_D(pfnFreeParse, (LPFREEPARSE)GetProcAddress(lpCurAddIn->hMod, szFreeParse));

					if (pfnSmartViewParse && pfnFreeParse)
					{
						LPWSTR szParse = pfnSmartViewParse(szStructType.c_str(), cbBin, lpBin);
						wstring szRet = szParse;

						pfnFreeParse(szParse);
						return szRet;
					}
				}
			}
		}

		lpCurAddIn = lpCurAddIn->lpNextAddIn;
	}

	return L"No parser found";
}