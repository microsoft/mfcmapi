// AddIns.cpp : Functions supporting AddIns

#include "stdafx.h"
#include "BaseDialog.h"
#include "MFCMAPI.h"
#include "ImportProcs.h"
#include "MAPIFunctions.h"
#include "PropTagArray.h"
#include "Editor.h"
#include "Guids.h"

LPADDIN g_lpMyAddins = NULL;

ULONG GetAddinVersion(HMODULE hMod)
{
	HRESULT hRes = S_OK;
	LPGETAPIVERSION pfnGetAPIVersion = NULL;
	WC_D(pfnGetAPIVersion, (LPGETAPIVERSION) GetProcAddress(hMod,szGetAPIVersion));
	if (pfnGetAPIVersion)
	{
		return pfnGetAPIVersion();
	}

	// Default case for unversioned add-ins
	return MFCMAPI_HEADER_V1;
}

void LoadSingleAddIn(LPADDIN lpAddIn, HMODULE hMod, LPLOADADDIN pfnLoadAddIn)
{
	DebugPrint(DBGAddInPlumbing,_T("Loading AddIn\n"));
	if (!lpAddIn) return;
	if (!pfnLoadAddIn) return;
	HRESULT hRes = S_OK;
	lpAddIn->hMod = hMod;
	pfnLoadAddIn(&lpAddIn->szName);
	if (lpAddIn->szName)
	{
		DebugPrint(DBGAddInPlumbing,_T("Loading \"%ws\"\n"),lpAddIn->szName);
	}

	ULONG ulVersion = GetAddinVersion(hMod);
	DebugPrint(DBGAddInPlumbing,_T("AddIn version = %d\n"),ulVersion);

	LPGETMENUS pfnGetMenus = NULL;
	WC_D(pfnGetMenus, (LPGETMENUS) GetProcAddress(hMod,szGetMenus));
	if (pfnGetMenus)
	{
		pfnGetMenus(&lpAddIn->ulMenu,&lpAddIn->lpMenu);
		if (!lpAddIn->ulMenu || !lpAddIn->lpMenu)
		{
			DebugPrint(DBGAddInPlumbing,_T("AddIn returned invalid menus\n"));
			lpAddIn->ulMenu = NULL;
			lpAddIn->lpMenu = NULL;
		}
		if (lpAddIn->ulMenu && lpAddIn->lpMenu)
		{
			ULONG ulMenu = 0;
			for (ulMenu = 0 ; ulMenu < lpAddIn->ulMenu ; ulMenu++)
			{
				// Save off our add-in struct
				lpAddIn->lpMenu[ulMenu].lpAddIn = lpAddIn;
				if (lpAddIn->lpMenu[ulMenu].szMenu)
					DebugPrint(DBGAddInPlumbing,_T("Menu: %ws\n"),lpAddIn->lpMenu[ulMenu].szMenu);
				if (lpAddIn->lpMenu[ulMenu].szHelp)
					DebugPrint(DBGAddInPlumbing,_T("Help: %ws\n"),lpAddIn->lpMenu[ulMenu].szHelp);
				DebugPrint(DBGAddInPlumbing,_T("ID: 0x%08X\n"),lpAddIn->lpMenu[ulMenu].ulID);
				DebugPrint(DBGAddInPlumbing,_T("Context: 0x%08X\n"),lpAddIn->lpMenu[ulMenu].ulContext);
				DebugPrint(DBGAddInPlumbing,_T("Flags: 0x%08X\n"),lpAddIn->lpMenu[ulMenu].ulFlags);
			}
		}
	}

	hRes = S_OK;
	LPGETPROPTAGS pfnGetPropTags = NULL;
	WC_D(pfnGetPropTags, (LPGETPROPTAGS) GetProcAddress(hMod,szGetPropTags));
	if (pfnGetPropTags)
	{
		pfnGetPropTags(&lpAddIn->ulPropTags,&lpAddIn->lpPropTags);
	}

	hRes = S_OK;
	LPGETPROPTYPES pfnGetPropTypes = NULL;
	WC_D(pfnGetPropTypes, (LPGETPROPTYPES) GetProcAddress(hMod,szGetPropTypes));
	if (pfnGetPropTypes)
	{
		pfnGetPropTypes(&lpAddIn->ulPropTypes,&lpAddIn->lpPropTypes);
	}

	hRes = S_OK;
	LPGETPROPGUIDS pfnGetPropGuids = NULL;
	WC_D(pfnGetPropGuids, (LPGETPROPGUIDS) GetProcAddress(hMod,szGetPropGuids));
	if (pfnGetPropGuids)
	{
		pfnGetPropGuids(&lpAddIn->ulPropGuids,&lpAddIn->lpPropGuids);
	}

	// v2 changed the LPNAMEID_ARRAY_ENTRY structure
	if (MFCMAPI_HEADER_V2 <= ulVersion)
	{
		hRes = S_OK;
		LPGETNAMEIDS pfnGetNameIDs = NULL;
		WC_D(pfnGetNameIDs, (LPGETNAMEIDS) GetProcAddress(hMod,szGetNameIDs));
		if (pfnGetNameIDs)
		{
			pfnGetNameIDs(&lpAddIn->ulNameIDs,&lpAddIn->lpNameIDs);
		}
	}

	hRes = S_OK;
	LPGETPROPFLAGS pfnGetPropFlags = NULL;
	WC_D(pfnGetPropFlags, (LPGETPROPFLAGS) GetProcAddress(hMod,szGetPropFlags));
	if (pfnGetPropFlags)
	{
		pfnGetPropFlags(&lpAddIn->ulPropFlags,&lpAddIn->lpPropFlags);
	}

	DebugPrint(DBGAddInPlumbing,_T("Done loading AddIn\n"));
}

void LoadAddIns()
{
	DebugPrint(DBGAddInPlumbing,_T("Loading AddIns\n"));
	// First, we look at each DLL in the current dir and see if it exports 'LoadAddIn'
	LPADDIN lpCurAddIn = NULL;
	// Allocate space to hold information on all DLLs in the directory

	HRESULT hRes = S_OK;
	TCHAR szFilePath[MAX_PATH];
	DWORD dwDir = NULL;
	EC_D(dwDir,GetModuleFileName(NULL,szFilePath,_countof(szFilePath)));
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

#define SPECLEN 6 // for '\\*.dll'
	if (szFilePath[0])
	{
		DebugPrint(DBGAddInPlumbing,_T("Current dir = \"%s\"\n"),szFilePath);
		hRes = StringCchCatN(szFilePath, MAX_PATH, _T("\\*.dll"), SPECLEN); // STRING_OK
		if (SUCCEEDED(hRes))
		{
			DebugPrint(DBGAddInPlumbing,_T("File spec = \"%s\"\n"),szFilePath);

			WIN32_FIND_DATA FindFileData = {0};
			HANDLE hFind = FindFirstFile(szFilePath, &FindFileData);

			if (hFind == INVALID_HANDLE_VALUE)
			{
				DebugPrint(DBGAddInPlumbing,_T("Invalid file handle. Error is %u.\n"),GetLastError());
			}
			else
			{
				while (true)
				{
					hRes = S_OK;
					DebugPrint(DBGAddInPlumbing,_T("Examining \"%s\"\n"),FindFileData.cFileName);
					HMODULE hMod = NULL;

					// LoadLibrary calls DLLMain, which can get expensive just to see if a function is exported
					// So we use DONT_RESOLVE_DLL_REFERENCES to see if we're interested in the DLL first
					// Only if we're interested do we reload the DLL for real
					hMod = LoadLibraryEx(FindFileData.cFileName,NULL,DONT_RESOLVE_DLL_REFERENCES);
					if (hMod)
					{
						LPLOADADDIN pfnLoadAddIn = NULL;
						WC_D(pfnLoadAddIn, (LPLOADADDIN) GetProcAddress(hMod,szLoadAddIn));
						FreeLibrary(hMod);
						hMod = NULL;

						if (pfnLoadAddIn)
						{
							// We found a candidate, load it for real now
							hMod = MyLoadLibrary(FindFileData.cFileName);
							// GetProcAddress again just in case we loaded at a different address
							WC_D(pfnLoadAddIn, (LPLOADADDIN) GetProcAddress(hMod,szLoadAddIn));
						}
					}
					if (hMod)
					{
						DebugPrint(DBGAddInPlumbing,_T("Opened module\n"));
						LPLOADADDIN pfnLoadAddIn = NULL;
						WC_D(pfnLoadAddIn, (LPLOADADDIN) GetProcAddress(hMod,szLoadAddIn));

						if (pfnLoadAddIn)
						{
							DebugPrint(DBGAddInPlumbing,_T("Found an add-in\n"));
							// Add a node
							if (!lpCurAddIn)
							{
								if (!g_lpMyAddins)
								{
									g_lpMyAddins = new _AddIn;
									ZeroMemory(g_lpMyAddins,sizeof(_AddIn));
								}
								lpCurAddIn = g_lpMyAddins;
							}
							else if (lpCurAddIn)
							{
								lpCurAddIn->lpNextAddIn = new _AddIn;
								ZeroMemory(lpCurAddIn->lpNextAddIn,sizeof(_AddIn));
								lpCurAddIn = lpCurAddIn->lpNextAddIn;
							}

							// Now that we have a node, populate it
							if (lpCurAddIn)
							{
								LoadSingleAddIn(lpCurAddIn,hMod,pfnLoadAddIn);
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
					DebugPrint(DBGAddInPlumbing,_T("FindNextFile error. Error is %u.\n"),dwRet);
				}
			}
		}
	}

	MergeAddInArrays();
	SortAddInArrays();
	DebugPrint(DBGAddInPlumbing,_T("Done loading AddIns\n"));
}

void ResetArrays()
{
	if (PropTypeArray != g_PropTypeArray) delete PropTypeArray;
	if (PropTagArray != g_PropTagArray) delete PropTagArray;
	if (PropGuidArray != g_PropGuidArray) delete PropGuidArray;
	if (NameIDArray != g_NameIDArray) delete NameIDArray;
	if (FlagArray != g_FlagArray) delete FlagArray;

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
}

void UnloadAddIns()
{
	DebugPrint(DBGAddInPlumbing,_T("Unloading AddIns\n"));
	if (g_lpMyAddins)
	{
		LPADDIN lpCurAddIn = g_lpMyAddins;
		while (lpCurAddIn)
		{
			DebugPrint(DBGAddInPlumbing,_T("Freeing add-in\n"));
			if (lpCurAddIn->hMod)
			{
				HRESULT hRes = S_OK;
				if (lpCurAddIn->szName)
				{
					DebugPrint(DBGAddInPlumbing,_T("Unloading \"%ws\"\n"),lpCurAddIn->szName);
				}
				LPUNLOADADDIN pfnUnLoadAddIn = NULL;
				WC_D(pfnUnLoadAddIn, (LPUNLOADADDIN) GetProcAddress(lpCurAddIn->hMod,szUnloadAddIn));
				if (pfnUnLoadAddIn) pfnUnLoadAddIn();

				FreeLibrary(lpCurAddIn->hMod);
			}
			LPADDIN lpAddInToFree = lpCurAddIn;
			lpCurAddIn = lpCurAddIn->lpNextAddIn;
			delete lpAddInToFree;
		}
	}

	ResetArrays();

	DebugPrint(DBGAddInPlumbing,_T("Done unloading AddIns\n"));
}

// Adds menu items appropriate to the context
// Returns number of menu items added
ULONG ExtendAddInMenu(HMENU hMenu, ULONG ulAddInContext)
{
	DebugPrint(DBGAddInPlumbing,_T("Extending menus, ulAddInContext = 0x%08X\n"),ulAddInContext);
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
			DebugPrint(DBGAddInPlumbing,_T("Examining add-in for menus\n"));
			if (lpCurAddIn->hMod)
			{
				HRESULT hRes = S_OK;
				if (lpCurAddIn->szName)
				{
					DebugPrint(DBGAddInPlumbing,_T("Examining \"%ws\"\n"),lpCurAddIn->szName);
				}
				ULONG ulMenu = 0;
				for (ulMenu = 0 ; ulMenu <lpCurAddIn->ulMenu && SUCCEEDED(hRes); ulMenu++)
				{
					if ((lpCurAddIn->lpMenu[ulMenu].ulFlags & MENU_FLAGS_SINGLESELECT) &&
						(lpCurAddIn->lpMenu[ulMenu].ulFlags & MENU_FLAGS_MULTISELECT))
					{
						// Invalid combo of flags - don't add the menu
						DebugPrint(DBGAddInPlumbing,_T("Invalid flags on menu \"%ws\" in add-in \"%ws\"\n"),lpCurAddIn->lpMenu[ulMenu].szMenu,lpCurAddIn->szName);
						DebugPrint(DBGAddInPlumbing,_T("MENU_FLAGS_SINGLESELECT and MENU_FLAGS_MULTISELECT cannot be combined\n"));
						continue;
					}
					if (lpCurAddIn->lpMenu[ulMenu].ulContext & ulAddInContext)
					{
						if (!hAddInMenu)
						{
							WCHAR szAddInTitle[8] = {0}; // The length of IDS_ADDINSMENU
							int iRet = NULL;
							// CString doesn't provide a way to extract just Unicode strings, so we do this manually
							EC_D(iRet,LoadStringW(GetModuleHandle(NULL),
								IDS_ADDINSMENU,
								szAddInTitle,
								_countof(szAddInTitle)));

							hAddInMenu = CreatePopupMenu();

							MENUITEMINFOW topMenu = {0};
							topMenu.cbSize = sizeof(MENUITEMINFO);
							topMenu.fMask = MIIM_SUBMENU|MIIM_TYPE|MIIM_STATE|MIIM_ID;
							topMenu.hSubMenu = hAddInMenu;
							topMenu.fState = MFS_ENABLED;
							topMenu.dwTypeData = szAddInTitle;
							EC_B(InsertMenuItemW(
								hMenu,
								(UINT) -1,
								MF_BYPOSITION | MF_POPUP | MF_STRING,
								&topMenu));
						}
						if (SUCCEEDED(hRes))
						{
							MENUITEMINFOW subMenu = {0};
							subMenu.cbSize = sizeof(MENUITEMINFO);
							subMenu.fMask = MIIM_TYPE|MIIM_STATE|MIIM_ID|MIIM_DATA;
							subMenu.fType = MFT_STRING;
							subMenu.fState = MFS_ENABLED;
							subMenu.wID = uidCurMenu;
							subMenu.dwTypeData = lpCurAddIn->lpMenu[ulMenu].szMenu;
							subMenu.dwItemData = (ULONG_PTR) &lpCurAddIn->lpMenu[ulMenu];

							EC_B(InsertMenuItemW(
								hAddInMenu,
								(UINT) -1,
								MF_BYPOSITION | MF_POPUP | MF_STRING,
								&subMenu));
							uidCurMenu++;
						}
					}
				}
			}
			lpCurAddIn = lpCurAddIn->lpNextAddIn;
		}
	}
	DestroyMenu(hAddInMenu);
	DebugPrint(DBGAddInPlumbing,_T("Done extending menus\n"));
	return uidCurMenu-ID_ADDINMENU;
}

LPMENUITEM GetAddinMenuItem(HWND hWnd, UINT uidMsg)
{
	if (uidMsg < ID_ADDINMENU) return NULL;

	HRESULT hRes = S_OK;
	MENUITEMINFOW subMenu = {0};
	subMenu.cbSize = sizeof(MENUITEMINFO);
	subMenu.fMask = MIIM_STATE|MIIM_ID|MIIM_DATA;

	WC_B(GetMenuItemInfoW(
		::GetMenu(hWnd),
		uidMsg,
		false,
		&subMenu));
	return (LPMENUITEM) subMenu.dwItemData;
}

void InvokeAddInMenu(LPADDINMENUPARAMS lpParams)
{
	HRESULT hRes = S_OK;

	if (!lpParams) return;
	if (!lpParams->lpAddInMenu) return;
	if (!lpParams->lpAddInMenu->lpAddIn) return;

	if (!lpParams->lpAddInMenu->lpAddIn->pfnCallMenu)
	{
		if (!lpParams->lpAddInMenu->lpAddIn->hMod) return;
		WC_D(lpParams->lpAddInMenu->lpAddIn->pfnCallMenu, (LPCALLMENU) GetProcAddress(
			lpParams->lpAddInMenu->lpAddIn->hMod,
			szCallMenu));
	}

	if (!lpParams->lpAddInMenu->lpAddIn->pfnCallMenu)
	{
		DebugPrint(DBGAddInPlumbing,_T("InvokeAddInMenu: CallMenu not found\n"));
		return;
	}

	WC_H(lpParams->lpAddInMenu->lpAddIn->pfnCallMenu(lpParams));

	return;
}

LPNAME_ARRAY_ENTRY PropTypeArray = NULL;
ULONG ulPropTypeArray = NULL;

LPNAME_ARRAY_ENTRY PropTagArray = NULL;
ULONG ulPropTagArray = NULL;

LPGUID_ARRAY_ENTRY PropGuidArray = NULL;
ULONG ulPropGuidArray = NULL;

LPNAMEID_ARRAY_ENTRY NameIDArray = NULL;
ULONG ulNameIDArray;

LPFLAG_ARRAY_ENTRY FlagArray = NULL;
ULONG ulFlagArray = NULL;

void MergeAddInArrays()
{
	DebugPrint(DBGAddInPlumbing,_T("Merging Add-In arrays\n"));

	ResetArrays();

	DebugPrint(DBGAddInPlumbing,_T("Found 0x%08X built in prop tags.\n"),g_ulPropTagArray);
	DebugPrint(DBGAddInPlumbing,_T("Found 0x%08X built in prop types.\n"),g_ulPropTypeArray);
	DebugPrint(DBGAddInPlumbing,_T("Found 0x%08X built in guids.\n"),g_ulPropGuidArray);
	DebugPrint(DBGAddInPlumbing,_T("Found 0x%08X built in named ids.\n"),g_ulNameIDArray);
	DebugPrint(DBGAddInPlumbing,_T("Found 0x%08X built in flags.\n"),g_ulFlagArray);

	// No add-in == nothing to merge
	if (!g_lpMyAddins) return;

	// First pass - count up any arrays we have to merge
	ULONG ulAddInPropTypeArray = g_ulPropTypeArray;
	ULONG ulAddInPropTagArray = g_ulPropTagArray;
	ULONG ulAddInPropGuidArray = g_ulPropGuidArray;
	ULONG ulAddInNameIDArray = g_ulNameIDArray;
	ULONG ulAddInFlagArray = g_ulFlagArray;

	LPADDIN lpCurAddIn = g_lpMyAddins;
	while (lpCurAddIn)
	{
		DebugPrint(DBGAddInPlumbing,_T("Looking at %ws\n"),lpCurAddIn->szName);
		DebugPrint(DBGAddInPlumbing,_T("Found 0x%08X prop tags.\n"),lpCurAddIn->ulPropTags);
		DebugPrint(DBGAddInPlumbing,_T("Found 0x%08X prop types.\n"),lpCurAddIn->ulPropTypes);
		DebugPrint(DBGAddInPlumbing,_T("Found 0x%08X guids.\n"),lpCurAddIn->ulPropGuids);
		DebugPrint(DBGAddInPlumbing,_T("Found 0x%08X named ids.\n"),lpCurAddIn->ulNameIDs);
		DebugPrint(DBGAddInPlumbing,_T("Found 0x%08X flags.\n"),lpCurAddIn->ulPropFlags);
		ulAddInPropTypeArray += lpCurAddIn->ulPropTypes;
		ulAddInPropTagArray += lpCurAddIn->ulPropTags;
		ulAddInPropGuidArray += lpCurAddIn->ulPropGuids;
		ulAddInNameIDArray += lpCurAddIn->ulNameIDs;
		ulAddInFlagArray += lpCurAddIn->ulPropFlags;
		lpCurAddIn = lpCurAddIn->lpNextAddIn;
	}

	// if count has gone up - we need to merge
	// Allocate a larger array and initialize it with our hardcoded data
	ULONG i = 0;
	if (ulAddInPropTypeArray != g_ulPropTypeArray)
	{
		PropTypeArray = new NAME_ARRAY_ENTRY[ulAddInPropTypeArray];
		for (i = 0 ; i < g_ulPropTypeArray ; i++)
		{
			PropTypeArray[i] = g_PropTypeArray[i];
		}
	}
	if (ulAddInPropTagArray != g_ulPropTagArray)
	{
		PropTagArray = new NAME_ARRAY_ENTRY[ulAddInPropTagArray];
		for (i = 0 ; i < g_ulPropTagArray ; i++)
		{
			PropTagArray[i] = g_PropTagArray[i];
		}
	}
	if (ulAddInPropGuidArray != g_ulPropGuidArray)
	{
		PropGuidArray = new GUID_ARRAY_ENTRY[ulAddInPropGuidArray];
		for (i = 0 ; i < g_ulPropGuidArray ; i++)
		{
			PropGuidArray[i] = g_PropGuidArray[i];
		}
	}
	if (ulAddInNameIDArray != g_ulNameIDArray)
	{
		NameIDArray = new NAMEID_ARRAY_ENTRY[ulAddInNameIDArray];
		for (i = 0 ; i < g_ulNameIDArray ; i++)
		{
			NameIDArray[i] = g_NameIDArray[i];
		}
	}
	if (ulAddInFlagArray != g_ulFlagArray)
	{
		FlagArray = new FLAG_ARRAY_ENTRY[ulAddInFlagArray];
		for (i = 0 ; i < g_ulFlagArray ; i++)
		{
			FlagArray[i] = g_FlagArray[i];
		}
	}

	// Second pass - append the new arrays to the hardcoded arrays
	ULONG ulCurPropType = g_ulPropTypeArray;
	ULONG ulCurPropTag = g_ulPropTagArray;
	ULONG ulCurPropGuid = g_ulPropGuidArray;
	ULONG ulCurNameID = g_ulNameIDArray;
	ULONG ulCurFlag = g_ulFlagArray;
	lpCurAddIn = g_lpMyAddins;
	while (lpCurAddIn)
	{
		if (lpCurAddIn->ulPropTypes)
		{
			for (i = ulCurPropType; i < ulCurPropType+lpCurAddIn->ulPropTypes ; i++)
			{
				PropTypeArray[i] = lpCurAddIn->lpPropTypes[i-ulCurPropType];
			}
			ulCurPropType += lpCurAddIn->ulPropTypes;
		}
		if (lpCurAddIn->ulPropTags)
		{
			for (i = ulCurPropTag; i < ulCurPropTag+lpCurAddIn->ulPropTags ; i++)
			{
				PropTagArray[i] = lpCurAddIn->lpPropTags[i-ulCurPropTag];
			}
			ulCurPropTag += lpCurAddIn->ulPropTags;
		}
		if (lpCurAddIn->ulPropGuids)
		{
			for (i = ulCurPropGuid; i < ulCurPropGuid+lpCurAddIn->ulPropGuids ; i++)
			{
				PropGuidArray[i] = lpCurAddIn->lpPropGuids[i-ulCurPropGuid];
			}
			ulCurPropGuid += lpCurAddIn->ulPropGuids;
		}
		if (lpCurAddIn->ulNameIDs)
		{
			for (i = ulCurNameID; i < ulCurNameID+lpCurAddIn->ulNameIDs ; i++)
			{
				NameIDArray[i] = lpCurAddIn->lpNameIDs[i-ulCurNameID];
			}
			ulCurNameID += lpCurAddIn->ulNameIDs;
		}
		if (lpCurAddIn->ulPropFlags)
		{
			for (i = ulCurFlag; i < ulCurFlag+lpCurAddIn->ulPropFlags ; i++)
			{
				FlagArray[i] = lpCurAddIn->lpPropFlags[i-ulCurFlag];
			}
			ulCurFlag += lpCurAddIn->ulPropFlags;
		}
		lpCurAddIn = lpCurAddIn->lpNextAddIn;
	}
	ulPropTypeArray = ulCurPropType;
	ulPropTagArray = ulCurPropTag;
	ulPropGuidArray = ulCurPropGuid;
	ulNameIDArray = ulCurNameID;
	ulFlagArray = ulCurFlag;

	DebugPrint(DBGAddInPlumbing,_T("Done merging add-in arrays\n"));
}

// Sort arrays and eliminate duplicates
void SortAddInArrays()
{
	DebugPrint(DBGAddInPlumbing,_T("Sorting add-in arrays\n"));

	// insertion sort on PropTagArray/ulPropTagArray
	ULONG i = 0;
	ULONG iUnsorted = 0;
	ULONG iLoc = 0;
	for (iUnsorted = 1;iUnsorted<ulPropTagArray;iUnsorted++)
	{
		NAME_ARRAY_ENTRY NextItem = PropTagArray[iUnsorted];
		for (iLoc = iUnsorted;iLoc > 0;iLoc--)
		{
			if (PropTagArray[iLoc-1].ulValue < NextItem.ulValue) break;
			if (PropTagArray[iLoc-1].ulValue == NextItem.ulValue &&
				wcscmp(NextItem.lpszName,PropTagArray[iLoc-1].lpszName) >= 0) break;
			PropTagArray[iLoc] = PropTagArray[iLoc-1];
		}
		PropTagArray[iLoc] = NextItem;
	}

	// Move the dupes to the end of the array and truncate the array
	for (i = 0;i < ulPropTagArray-1;)
	{
		if ((PropTagArray[i].ulValue == PropTagArray[i+1].ulValue) &&
			!wcscmp(PropTagArray[i].lpszName,PropTagArray[i+1].lpszName))
		{
			for (iLoc = i+1;iLoc < ulPropTagArray-1; iLoc++)
			{
				PropTagArray[iLoc] = PropTagArray[iLoc+1];
			}
			ulPropTagArray--;
		}
		else i++;
	}
	DebugPrint(DBGAddInPlumbing,_T("After sort, 0x%08X prop tags.\n"),ulPropTagArray);

	// insertion sort on PropTypeArray/ulPropTypeArray
	for (iUnsorted = 1;iUnsorted<ulPropTypeArray;iUnsorted++)
	{
		NAME_ARRAY_ENTRY NextItem = PropTypeArray[iUnsorted];
		for (iLoc = iUnsorted;iLoc > 0;iLoc--)
		{
			if (PropTypeArray[iLoc-1].ulValue < NextItem.ulValue) break;
			if (PropTypeArray[iLoc-1].ulValue == NextItem.ulValue &&
				wcscmp(NextItem.lpszName,PropTypeArray[iLoc-1].lpszName) >= 0) break;
			PropTypeArray[iLoc] = PropTypeArray[iLoc-1];
		}
		PropTypeArray[iLoc] = NextItem;
	}

	// Move the dupes to the end of the array and truncate the array
	for (i = 0;i < ulPropTypeArray-1;)
	{
		if ((PropTypeArray[i].ulValue == PropTypeArray[i+1].ulValue) &&
			!wcscmp(PropTypeArray[i].lpszName,PropTypeArray[i+1].lpszName))
		{
			for (iLoc = i+1;iLoc < ulPropTypeArray-1; iLoc++)
			{
				PropTypeArray[iLoc] = PropTypeArray[iLoc+1];
			}
			ulPropTypeArray--;
		}
		else i++;
	}
	DebugPrint(DBGAddInPlumbing,_T("After sort, 0x%08X prop types.\n"),ulPropTypeArray);

	// No sort on PropGuidArray/ulPropGuidArray
	// Move the dupes to the end of the array and truncate the array
	for (i = 0;i < ulPropGuidArray-1;i++)
	{
		ULONG iCur = 0;
		for (iCur = i+1;iCur < ulPropGuidArray;)
		{
			if (IsEqualGUID(*PropGuidArray[i].lpGuid,*PropGuidArray[iCur].lpGuid))
			{
				for (iLoc = iCur;iLoc < ulPropGuidArray-1; iLoc++)
				{
					PropGuidArray[iLoc] = PropGuidArray[iLoc+1];
				}
				ulPropGuidArray--;
			}
			else iCur++;
		}
	}
	DebugPrint(DBGAddInPlumbing,_T("After sort, 0x%08X guids.\n"),ulPropGuidArray);

	// insertion sort on NameIDArray/ulNameIDArray
	for (iUnsorted = 1;iUnsorted<ulNameIDArray;iUnsorted++)
	{
		NAMEID_ARRAY_ENTRY NextItem = NameIDArray[iUnsorted];
		for (iLoc = iUnsorted;iLoc > 0;iLoc--)
		{
			if (NameIDArray[iLoc-1].lValue < NextItem.lValue) break;
			if (NameIDArray[iLoc-1].lValue == NextItem.lValue &&
				wcscmp(NextItem.lpszName,NameIDArray[iLoc-1].lpszName) >= 0) break;
			NameIDArray[iLoc] = NameIDArray[iLoc-1];
		}

		NameIDArray[iLoc] = NextItem;
	}

	// Move the dupes to the end of the array and truncate the array
	for (i = 0;i < ulNameIDArray-1;)
	{
		ULONG iMatch = i+1;
		for (; iMatch < ulNameIDArray-1 ; iMatch++)
		{
			if (NameIDArray[i].lValue != NameIDArray[iMatch].lValue) break;
			if (wcscmp(NameIDArray[i].lpszName,NameIDArray[iMatch].lpszName)) break;
			if (IsEqualGUID(*NameIDArray[i].lpGuid,*NameIDArray[iMatch].lpGuid))
			{
				for (iLoc = iMatch;iLoc < ulNameIDArray-1; iLoc++)
				{
					NameIDArray[iLoc] = NameIDArray[iLoc+1];
				}
				ulNameIDArray--;
			}
		}
		i++;
	}
	DebugPrint(DBGAddInPlumbing,_T("After sort, 0x%08X named ids.\n"),ulNameIDArray);

	// Flags are difficult to sort
	// Sort order
	// 1 - prev.ulFlagName <= next.ulFlagName
	// 2 - ?? We have to preserve a stable order - so just delete dupes here
	for (iUnsorted = 1;iUnsorted<ulFlagArray;iUnsorted++)
	{
		FLAG_ARRAY_ENTRY NextItem = FlagArray[iUnsorted];
		for (iLoc = iUnsorted;iLoc > 0;iLoc--)
		{
			if (FlagArray[iLoc-1].ulFlagName <= NextItem.ulFlagName) break;
			FlagArray[iLoc] = FlagArray[iLoc-1];
		}
		FlagArray[iLoc] = NextItem;
	}

	// Move the dupes to the end of the array and truncate the array
	for (i = 0;i < ulFlagArray-1;)
	{
		ULONG iMatch = i+1;
		for (; iMatch < ulFlagArray-1 ; iMatch++)
		{
			if (FlagArray[i].ulFlagName != FlagArray[iMatch].ulFlagName) break;
			if (FlagArray[i].lFlagValue == FlagArray[iMatch].lFlagValue &&
				FlagArray[i].ulFlagType == FlagArray[iMatch].ulFlagType &&
				!wcscmp(FlagArray[i].lpszName,FlagArray[iMatch].lpszName))
			{
				for (iLoc = iMatch;iLoc < ulFlagArray-1; iLoc++)
				{
					FlagArray[iLoc] = FlagArray[iLoc+1];
				}
				ulFlagArray--;
			}
		}
		i++;
	}
	DebugPrint(DBGAddInPlumbing,_T("After sort, 0x%08X flags.\n"),ulFlagArray);

	DebugPrint(DBGAddInPlumbing,_T("Done sorting add-in arrays\n"));
}

__declspec(dllexport) void __cdecl AddInLog(BOOL bPrintThreadTime, LPWSTR szMsg,...)
{
	if (!fIsSet(DBGAddIn)) return;
	HRESULT hRes = S_OK;

	if (!szMsg)
	{
		_Output(DBGGeneric, NULL, true, _T("AddInLog called with NULL szMsg!\n"));
		return;
	}

	va_list argList = NULL;
	va_start(argList, szMsg);

	WCHAR szAddInLogString[4096];
	WC_H(StringCchVPrintfW(szAddInLogString, _countof(szAddInLogString), szMsg, argList));
	if (FAILED(hRes))
	{
		_Output(DBGFatalError,NULL, true,_T("Debug output string not large enough to print everything to it\n"));
		// Since this function was 'safe', we've still got something we can print - send it on.
	}
	va_end(argList);

#ifdef UNICODE
	_Output(DBGAddIn, NULL, bPrintThreadTime, szAddInLogString);
#else
	char *szAnsiAddInLogString = NULL;
	EC_H(UnicodeToAnsi(szAddInLogString,&szAnsiAddInLogString));
	_Output(DBGAddIn, NULL, bPrintThreadTime, szAnsiAddInLogString);
	delete[] szAnsiAddInLogString;
#endif
}

__declspec(dllexport) HRESULT __cdecl SimpleDialog(LPWSTR szTitle, LPWSTR szMsg,...)
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

	WCHAR szDialogString[4096];
	WC_H(StringCchVPrintfW(szDialogString, _countof(szDialogString), szMsg, argList));
	if (FAILED(hRes))
	{
		_Output(DBGFatalError,NULL, true,_T("Debug output string not large enough to print everything to it\n"));
		// Since this function was 'safe', we've still got something we can print - send it on.
	}
	va_end(argList);

#ifdef UNICODE
	MySimpleDialog.SetPromptPostFix(szDialogString);
#else
	char *szAnsiDialogString = NULL;
	EC_H(UnicodeToAnsi(szDialogString,&szAnsiDialogString));
	MySimpleDialog.SetPromptPostFix(szAnsiDialogString);
	delete[] szAnsiDialogString;
#endif
	WC_H(MySimpleDialog.DisplayDialog());
	return hRes;
}

__declspec(dllexport) HRESULT __cdecl ComplexDialog(LPADDINDIALOG lpDialog, LPADDINDIALOGRESULT* lppDialogResult)
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
	EC_H(UnicodeToAnsi(lpDialog->szPrompt,&szAnsiPrompt));
	MyComplexDialog.SetPromptPostFix(szAnsiPrompt);
	delete[] szAnsiPrompt;
#endif

	if (lpDialog->ulNumControls && lpDialog->lpDialogControls)
	{
		for (i = 0 ; i < lpDialog->ulNumControls ; i++)
		{
			MyComplexDialog.SetAddInLabel(i,lpDialog->lpDialogControls[i].szLabel);
			switch (lpDialog->lpDialogControls[i].cType)
			{
			case ADDIN_CTRL_CHECK:
				{
					MyComplexDialog.InitCheck(
						i,
						NULL,
						lpDialog->lpDialogControls[i].bDefaultCheckState,
						lpDialog->lpDialogControls[i].bReadOnly);
					break;
				}
			case ADDIN_CTRL_EDIT_TEXT:
				{
					LPTSTR szDefaultText = NULL;
#ifdef UNICODE
					szDefaultText = lpDialog->lpDialogControls[i].szDefaultText;
#else
					EC_H(UnicodeToAnsi(lpDialog->lpDialogControls[i].szDefaultText,&szDefaultText));
#endif
					if (lpDialog->lpDialogControls[i].bMultiLine)
					{
						MyComplexDialog.InitMultiLine(
							i,
							NULL,
							szDefaultText,
							lpDialog->lpDialogControls[i].bReadOnly);
					}
					else
					{
						MyComplexDialog.InitSingleLineSz(
							i,
							NULL,
							szDefaultText,
							lpDialog->lpDialogControls[i].bReadOnly);

					}
#ifndef UNICODE
					delete[] szDefaultText;
#endif
					break;
				}
			case ADDIN_CTRL_EDIT_BINARY:
				{
					if (lpDialog->lpDialogControls[i].bMultiLine)
					{
						MyComplexDialog.InitMultiLine(
							i,
							NULL,
							NULL,
							lpDialog->lpDialogControls[i].bReadOnly);
					}
					else
					{
						MyComplexDialog.InitSingleLineSz(
							i,
							NULL,
							NULL,
							lpDialog->lpDialogControls[i].bReadOnly);

					}
					MyComplexDialog.SetBinary(
						i,
						lpDialog->lpDialogControls[i].lpBin,
						lpDialog->lpDialogControls[i].cbBin);
					break;
				}
			case ADDIN_CTRL_EDIT_NUM_DECIMAL:
				{
					MyComplexDialog.InitSingleLineSz(
						i,
						NULL,
						NULL,
						lpDialog->lpDialogControls[i].bReadOnly);
					MyComplexDialog.SetDecimal(
						i,
						lpDialog->lpDialogControls[i].ulDefaultNum);
					break;
				}
			case ADDIN_CTRL_EDIT_NUM_HEX:
				{
					MyComplexDialog.InitSingleLineSz(
						i,
						NULL,
						NULL,
						lpDialog->lpDialogControls[i].bReadOnly);
					MyComplexDialog.SetHex(
						i,
						lpDialog->lpDialogControls[i].ulDefaultNum);
					break;
				}
			}
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
				ZeroMemory(lpResults->lpDialogControlResults,sizeof(_AddInDialogControlResult)*lpDialog->ulNumControls);
				for (i = 0 ; i < lpDialog->ulNumControls ; i++)
				{
					lpResults->lpDialogControlResults[i].cType = lpDialog->lpDialogControls[i].cType;
					switch (lpDialog->lpDialogControls[i].cType)
					{
					case ADDIN_CTRL_CHECK:
						{
							lpResults->lpDialogControlResults[i].bCheckState = MyComplexDialog.GetCheck(i);
							break;
						}
					case ADDIN_CTRL_EDIT_TEXT:
						{
							LPWSTR szText = MyComplexDialog.GetStringW(i);
							if (szText)
							{
								size_t cchText = NULL;
								StringCchLengthW(szText,STRSAFE_MAX_CCH,&cchText);

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
						{
							// GetEntryID does just what we want - abuse it
							WC_H(MyComplexDialog.GetEntryID(
								i,
								false,
								&lpResults->lpDialogControlResults[i].cbBin,
								(LPENTRYID*)&lpResults->lpDialogControlResults[i].lpBin));
							break;
						}
					case ADDIN_CTRL_EDIT_NUM_DECIMAL:
						{
							lpResults->lpDialogControlResults[i].ulVal = MyComplexDialog.GetDecimal(i);
							break;
						}
					case ADDIN_CTRL_EDIT_NUM_HEX:
						{
							lpResults->lpDialogControlResults[i].ulVal = MyComplexDialog.GetHex(i);
							break;
						}
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
}

__declspec(dllexport) void __cdecl FreeDialogResult(LPADDINDIALOGRESULT lpDialogResult)
{
	if (lpDialogResult)
	{
		if (lpDialogResult->lpDialogControlResults)
		{
			ULONG i = 0;
			for (i = 0 ; i < lpDialogResult->ulNumControls ; i++)
			{
				delete[] lpDialogResult->lpDialogControlResults[i].lpBin;
				delete[] lpDialogResult->lpDialogControlResults[i].szText;
			}
		}
		delete[] lpDialogResult;
	}
}

__declspec(dllexport) void __cdecl GetMAPIModule(HMODULE* lphModule, BOOL bForce)
{
	if (!lphModule) return;
	*lphModule = NULL;
	if (hModMSMAPI)
	{
		*lphModule = hModMSMAPI;
	}
	else if (hModMAPI)
	{
		*lphModule = hModMSMAPI;
	}
	else if (bForce)
	{
		// No MAPI loaded - load it then try again
		AutoLoadMAPI();
		GetMAPIModule(lphModule,false);
	}
}


// Search for properties matching lpszPropName on a substring
LPNAMEID_ARRAY_ENTRY GetDispIDFromName(LPCWSTR lpszDispIDName)
{
	if (!lpszDispIDName) return NULL;

	ULONG ulCur = 0;

	for (ulCur = 0 ; ulCur < ulNameIDArray ; ulCur++)
	{
		if (0 == wcscmp(NameIDArray[ulCur].lpszName,lpszDispIDName))
		{
			// PSUNKNOWN is used as a placeholder in NameIDArray - don't return matching entries
			if (IsEqualGUID(*NameIDArray[ulCur].lpGuid,PSUNKNOWN)) return NULL;

			return &NameIDArray[ulCur];
		}
	}
	return NULL;
} // GetDispIDFromName