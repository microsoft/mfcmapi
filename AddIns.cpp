#include "stdafx.h"
#include "MFCMAPI.h"
#include "ImportProcs.h"
#include <Interpret/String.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <Interpret/Guids.h>
#ifndef MRMAPI
#include <UI/UIFunctions.h>
#endif

// Our built in arrays, which get merged into the arrays declared in mfcmapi.h
#include <Interpret/GenTagArray.h>
#include <Interpret/Flags.h>
#include <Interpret/GUIDArray.h>
#include <Interpret/NameIDArray.h>
#include <Interpret/PropTypeArray.h>
#include <Interpret/SmartView/SmartViewParsers.h>

vector<NAME_ARRAY_ENTRY_V2> PropTagArray;
vector<NAME_ARRAY_ENTRY> PropTypeArray;
vector<GUID_ARRAY_ENTRY> PropGuidArray;
vector<NAMEID_ARRAY_ENTRY> NameIDArray;
vector<FLAG_ARRAY_ENTRY> FlagArray;
vector<SMARTVIEW_PARSER_ARRAY_ENTRY> SmartViewParserArray;
vector<NAME_ARRAY_ENTRY> SmartViewParserTypeArray;

vector<_AddIn> g_lpMyAddins;

template <typename T> T GetFunction(
	HMODULE hMod,
	LPCSTR szFuncName)
{
	auto hRes = S_OK;
	T pObj = nullptr;
	WC_D(pObj, reinterpret_cast<T>(GetProcAddress(hMod, szFuncName)));
	return pObj;
}

_Check_return_ ULONG GetAddinVersion(HMODULE hMod)
{
	auto pfnGetAPIVersion = GetFunction<LPGETAPIVERSION>(hMod, szGetAPIVersion);
	if (pfnGetAPIVersion)
	{
		return pfnGetAPIVersion();
	}

	// Default case for unversioned add-ins
	return MFCMAPI_HEADER_V1;
}

// Load NAME_ARRAY_ENTRY style props from the add-in and upconvert them to V2
void LoadLegacyPropTags(
	HMODULE hMod,
	_In_ ULONG* lpulPropTags, // Number of entries in lppPropTags
	_In_ LPNAME_ARRAY_ENTRY_V2* lppPropTags // Array of NAME_ARRAY_ENTRY_V2 structures
)
{
	auto pfnGetPropTags = GetFunction<LPGETPROPTAGS>(hMod, szGetPropTags);
	if (pfnGetPropTags)
	{
		LPNAME_ARRAY_ENTRY lpPropTags = nullptr;
		pfnGetPropTags(lpulPropTags, &lpPropTags);
		if (lpPropTags && *lpulPropTags)
		{
			auto lpPropTagsV2 = new NAME_ARRAY_ENTRY_V2[*lpulPropTags];
			if (lpPropTagsV2)
			{
				for (ULONG i = 0; i < *lpulPropTags; i++)
				{
					lpPropTagsV2[i].lpszName = lpPropTags[i].lpszName;
					lpPropTagsV2[i].ulValue = lpPropTags[i].ulValue;
					lpPropTagsV2[i].ulSortOrder = 0;
				}

				*lppPropTags = lpPropTagsV2;
			}
		}
	}
}

void LoadSingleAddIn(_In_ _AddIn& addIn, HMODULE hMod, _In_ LPLOADADDIN pfnLoadAddIn)
{
	DebugPrint(DBGAddInPlumbing, L"Loading AddIn\n");
	if (!pfnLoadAddIn) return;
	addIn.hMod = hMod;
	pfnLoadAddIn(&addIn.szName);
	if (addIn.szName)
	{
		DebugPrint(DBGAddInPlumbing, L"Loading \"%ws\"\n", addIn.szName);
	}

	auto ulVersion = GetAddinVersion(hMod);
	DebugPrint(DBGAddInPlumbing, L"AddIn version = %u\n", ulVersion);

	auto pfnGetMenus = GetFunction<LPGETMENUS>(hMod, szGetMenus);
	if (pfnGetMenus)
	{
		pfnGetMenus(&addIn.ulMenu, &addIn.lpMenu);
		if (!addIn.ulMenu || !addIn.lpMenu)
		{
			DebugPrint(DBGAddInPlumbing, L"AddIn returned invalid menus\n");
			addIn.ulMenu = NULL;
			addIn.lpMenu = nullptr;
		}
		if (addIn.ulMenu && addIn.lpMenu)
		{
			for (ULONG ulMenu = 0; ulMenu < addIn.ulMenu; ulMenu++)
			{
				// Save off our add-in struct
				addIn.lpMenu[ulMenu].lpAddIn = &addIn;
				if (addIn.lpMenu[ulMenu].szMenu)
					DebugPrint(DBGAddInPlumbing, L"Menu: %ws\n", addIn.lpMenu[ulMenu].szMenu);
				if (addIn.lpMenu[ulMenu].szHelp)
					DebugPrint(DBGAddInPlumbing, L"Help: %ws\n", addIn.lpMenu[ulMenu].szHelp);
				DebugPrint(DBGAddInPlumbing, L"ID: 0x%08X\n", addIn.lpMenu[ulMenu].ulID);
				DebugPrint(DBGAddInPlumbing, L"Context: 0x%08X\n", addIn.lpMenu[ulMenu].ulContext);
				DebugPrint(DBGAddInPlumbing, L"Flags: 0x%08X\n", addIn.lpMenu[ulMenu].ulFlags);
			}
		}
	}

	addIn.bLegacyPropTags = false;
	auto pfnGetPropTagsV2 = GetFunction<LPGETPROPTAGSV2>(hMod, szGetPropTagsV2);
	if (pfnGetPropTagsV2)
	{
		pfnGetPropTagsV2(&addIn.ulPropTags, &addIn.lpPropTags);
	}
	else
	{
		LoadLegacyPropTags(hMod, &addIn.ulPropTags, &addIn.lpPropTags);
		addIn.bLegacyPropTags = true;
	}

	auto pfnGetPropTypes = GetFunction<LPGETPROPTYPES>(hMod, szGetPropTypes);
	if (pfnGetPropTypes)
	{
		pfnGetPropTypes(&addIn.ulPropTypes, &addIn.lpPropTypes);
	}

	auto pfnGetPropGuids = GetFunction<LPGETPROPGUIDS>(hMod, szGetPropGuids);
	if (pfnGetPropGuids)
	{
		pfnGetPropGuids(&addIn.ulPropGuids, &addIn.lpPropGuids);
	}

	// v2 changed the LPNAMEID_ARRAY_ENTRY structure
	if (MFCMAPI_HEADER_V2 <= ulVersion)
	{
		auto pfnGetNameIDs = GetFunction<LPGETNAMEIDS>(hMod, szGetNameIDs);
		if (pfnGetNameIDs)
		{
			pfnGetNameIDs(&addIn.ulNameIDs, &addIn.lpNameIDs);
		}
	}

	auto pfnGetPropFlags = GetFunction<LPGETPROPFLAGS>(hMod, szGetPropFlags);
	if (pfnGetPropFlags)
	{
		pfnGetPropFlags(&addIn.ulPropFlags, &addIn.lpPropFlags);
	}

	auto pfnGetSmartViewParserArray = GetFunction<LPGETSMARTVIEWPARSERARRAY>(hMod, szGetSmartViewParserArray);
	if (pfnGetSmartViewParserArray)
	{
		pfnGetSmartViewParserArray(&addIn.ulSmartViewParsers, &addIn.lpSmartViewParsers);
	}

	auto pfnGetSmartViewParserTypeArray = GetFunction<LPGETSMARTVIEWPARSERTYPEARRAY>(hMod, szGetSmartViewParserTypeArray);
	if (pfnGetSmartViewParserTypeArray)
	{
		pfnGetSmartViewParserTypeArray(&addIn.ulSmartViewParserTypes, &addIn.lpSmartViewParserTypes);
	}

	DebugPrint(DBGAddInPlumbing, L"Done loading AddIn\n");
}

class CFileList
{
public:
	wstring m_szKey;
	HKEY m_hRootKey;
	vector<wstring> m_lpList;

#define EXCLUSION_LIST L"AddInExclusionList" // STRING_OK
#define INCLUSION_LIST L"AddInInclusionList" // STRING_OK
#define SEPARATOR L";" // STRING_OK

	// Read in registry and build a list of invalid add-in DLLs
	CFileList(_In_ const wstring& szKey)
	{
		wstring lpszReg;

		m_hRootKey = CreateRootKey();
		m_szKey = szKey;

		if (m_hRootKey)
		{
			lpszReg = ReadStringFromRegistry(
				m_hRootKey,
				m_szKey);
		}

		if (!lpszReg.empty())
		{
			LPWSTR szContext = nullptr;
			auto szDLL = wcstok_s(LPWSTR(lpszReg.c_str()), SEPARATOR, &szContext);
			while (szDLL)
			{
				AddToList(szDLL);
				szDLL = wcstok_s(nullptr, SEPARATOR, &szContext);
			}
		}
	}

	// Write the list back to registry
	~CFileList()
	{
		auto hRes = S_OK;
		wstring szList;

		if (!m_lpList.empty())
		{
			for (const auto& dll : m_lpList)
			{
				szList += dll;
				szList += SEPARATOR;
			}

			WriteStringToRegistry(
				m_hRootKey,
				m_szKey,
				szList);
		}

		EC_W32(RegCloseKey(m_hRootKey));
	}

	// Add the DLL to the list
	void AddToList(_In_ const wstring& szDLL)
	{
		if (szDLL.empty()) return;
		m_lpList.push_back(szDLL);
	}

	// Check this DLL name against the list
	bool IsOnList(_In_ const wstring& szDLL) const
	{
		if (szDLL.empty()) return true;
		for (const auto& dll : m_lpList)
		{
			if (wstringToLower(dll) == wstringToLower(szDLL)) return true;
		}

		return false;
	}
};

void LoadAddIns()
{
	DebugPrint(DBGAddInPlumbing, L"Loading AddIns\n");
	// First, we look at each DLL in the current dir and see if it exports 'LoadAddIn'

	if (!RegKeys[regkeyLOADADDINS].ulCurDWORD)
	{
		DebugPrint(DBGAddInPlumbing, L"Bypassing add-in loading\n");
	}
	else
	{
		auto hRes = S_OK;
		WCHAR szFilePath[MAX_PATH];
		DWORD dwDir = NULL;
		EC_D(dwDir, GetModuleFileNameW(nullptr, szFilePath, _countof(szFilePath)));
		if (!dwDir) return;

		// We got the path to mfcmapi.exe - need to strip it
		wstring szDir = szFilePath;
		szDir = szDir.substr(0, szDir.find_last_of(L'\\'));

		CFileList ExclusionList(EXCLUSION_LIST);
		CFileList InclusionList(INCLUSION_LIST);

		if (!szDir.empty())
		{
			DebugPrint(DBGAddInPlumbing, L"Current dir = \"%ws\"\n", szDir.c_str());
			auto szSpec = szDir + L"\\*.dll"; // STRING_OK
			if (SUCCEEDED(hRes))
			{
				DebugPrint(DBGAddInPlumbing, L"File spec = \"%ws\"\n", szSpec.c_str());

				WIN32_FIND_DATAW FindFileData = { 0 };
				auto hFind = FindFirstFileW(szSpec.c_str(), &FindFileData);

				if (hFind == INVALID_HANDLE_VALUE)
				{
					DebugPrint(DBGAddInPlumbing, L"Invalid file handle. Error is %u.\n", GetLastError());
				}
				else
				{
					for (;;)
					{
						DebugPrint(DBGAddInPlumbing, L"Examining \"%ws\"\n", FindFileData.cFileName);
						HMODULE hMod = nullptr;

						// If we know the Add-in is good, just load it.
						// If we know it's bad, skip it.
						// Otherwise, we have to check if it's good.
						// LoadLibrary calls DLLMain, which can get expensive just to see if a function is exported
						// So we use DONT_RESOLVE_DLL_REFERENCES to see if we're interested in the DLL first
						// Only if we're interested do we reload the DLL for real
						if (InclusionList.IsOnList(FindFileData.cFileName))
						{
							hMod = MyLoadLibraryW(FindFileData.cFileName);
						}
						else
						{
							if (!ExclusionList.IsOnList(FindFileData.cFileName))
							{
								hMod = LoadLibraryExW(FindFileData.cFileName, nullptr, DONT_RESOLVE_DLL_REFERENCES);
								if (hMod)
								{
									auto pfnLoadAddIn = GetFunction<LPLOADADDIN>(hMod, szLoadAddIn);
									FreeLibrary(hMod);
									hMod = nullptr;

									if (pfnLoadAddIn)
									{
										// Remember this as a good add-in
										InclusionList.AddToList(FindFileData.cFileName);
										// We found a candidate, load it for real now
										hMod = MyLoadLibraryW(FindFileData.cFileName);
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
							auto pfnLoadAddIn = GetFunction<LPLOADADDIN>(hMod, szLoadAddIn);
							if (pfnLoadAddIn && GetAddinVersion(hMod) == MFCMAPI_HEADER_CURRENT_VERSION)
							{
								DebugPrint(DBGAddInPlumbing, L"Found an add-in\n");
								g_lpMyAddins.push_back(_AddIn());
								LoadSingleAddIn(g_lpMyAddins.back(), hMod, pfnLoadAddIn);
							}
							else
							{
								// Not an add-in, release the hMod
								FreeLibrary(hMod);
							}
						}

						// get next file
						if (!FindNextFileW(hFind, &FindFileData)) break;
					}

					auto dwRet = GetLastError();
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
}

void UnloadAddIns()
{
	DebugPrint(DBGAddInPlumbing, L"Unloading AddIns\n");

	for (const auto& addIn : g_lpMyAddins)
	{
		DebugPrint(DBGAddInPlumbing, L"Freeing add-in\n");
		if (addIn.bLegacyPropTags)
		{
			delete[] addIn.lpPropTags;
		}

		if (addIn.hMod)
		{
			if (addIn.szName)
			{
				DebugPrint(DBGAddInPlumbing, L"Unloading \"%ws\"\n", addIn.szName);
			}

			auto pfnUnLoadAddIn = GetFunction<LPUNLOADADDIN>(addIn.hMod, szUnloadAddIn);
			if (pfnUnLoadAddIn) pfnUnLoadAddIn();

			FreeLibrary(addIn.hMod);
		}
	}

	DebugPrint(DBGAddInPlumbing, L"Done unloading AddIns\n");
}

#ifndef MRMAPI
// Adds menu items appropriate to the context
// Returns number of menu items added
_Check_return_ ULONG ExtendAddInMenu(HMENU hMenu, ULONG ulAddInContext)
{
	DebugPrint(DBGAddInPlumbing, L"Extending menus, ulAddInContext = 0x%08X\n", ulAddInContext);
	HMENU hAddInMenu = nullptr;

	UINT uidCurMenu = ID_ADDINMENU;

	if (MENU_CONTEXT_PROPERTY == ulAddInContext)
	{
		uidCurMenu = ID_ADDINPROPERTYMENU;
	}

	for (const auto& addIn : g_lpMyAddins)
	{
		DebugPrint(DBGAddInPlumbing, L"Examining add-in for menus\n");
		if (addIn.hMod)
		{
			auto hRes = S_OK;
			if (addIn.szName)
			{
				DebugPrint(DBGAddInPlumbing, L"Examining \"%ws\"\n", addIn.szName);
			}

			for (ULONG ulMenu = 0; ulMenu < addIn.ulMenu && SUCCEEDED(hRes); ulMenu++)
			{
				if (addIn.lpMenu[ulMenu].ulFlags & MENU_FLAGS_SINGLESELECT &&
					addIn.lpMenu[ulMenu].ulFlags & MENU_FLAGS_MULTISELECT)
				{
					// Invalid combo of flags - don't add the menu
					DebugPrint(DBGAddInPlumbing, L"Invalid flags on menu \"%ws\" in add-in \"%ws\"\n", addIn.lpMenu[ulMenu].szMenu, addIn.szName);
					DebugPrint(DBGAddInPlumbing, L"MENU_FLAGS_SINGLESELECT and MENU_FLAGS_MULTISELECT cannot be combined\n");
					continue;
				}

				if (addIn.lpMenu[ulMenu].ulContext & ulAddInContext)
				{
					// Add the Add-Ins menu if we haven't added it already
					if (!hAddInMenu)
					{
						hAddInMenu = CreatePopupMenu();
						if (hAddInMenu)
						{
							InsertMenuW(
								hMenu,
								static_cast<UINT>(-1),
								MF_BYPOSITION | MF_POPUP | MF_ENABLED,
								reinterpret_cast<UINT_PTR>(hAddInMenu),
								loadstring(IDS_ADDINSMENU).c_str());
						}
						else continue;
					}

					// Now add each of the menu entries
					if (SUCCEEDED(hRes))
					{
						auto lpMenu = CreateMenuEntry(addIn.lpMenu[ulMenu].szMenu);
						if (lpMenu)
						{
							lpMenu->m_AddInData = reinterpret_cast<ULONG_PTR>(&addIn.lpMenu[ulMenu]);
						}

						EC_B(AppendMenuW(
							hAddInMenu,
							MF_ENABLED | MF_OWNERDRAW,
							uidCurMenu,
							reinterpret_cast<LPCWSTR>(lpMenu)));
						uidCurMenu++;
					}
				}
			}
		}
	}

	DebugPrint(DBGAddInPlumbing, L"Done extending menus\n");
	return uidCurMenu - ID_ADDINMENU;
}

_Check_return_ LPMENUITEM GetAddinMenuItem(HWND hWnd, UINT uidMsg)
{
	if (uidMsg < ID_ADDINMENU) return nullptr;

	MENUITEMINFOW subMenu = { 0 };
	subMenu.cbSize = sizeof(MENUITEMINFO);
	subMenu.fMask = MIIM_STATE | MIIM_ID | MIIM_DATA;

	if (GetMenuItemInfoW(
		GetMenu(hWnd),
		uidMsg,
		false,
		&subMenu) &&
		subMenu.dwItemData)
	{
		return reinterpret_cast<LPMENUITEM>(reinterpret_cast<LPMENUENTRY>(subMenu.dwItemData)->m_AddInData);
	}

	return nullptr;
}

void InvokeAddInMenu(_In_opt_ LPADDINMENUPARAMS lpParams)
{
	auto hRes = S_OK;

	if (!lpParams) return;
	if (!lpParams->lpAddInMenu) return;
	if (!lpParams->lpAddInMenu->lpAddIn) return;

	if (!lpParams->lpAddInMenu->lpAddIn->pfnCallMenu)
	{
		if (!lpParams->lpAddInMenu->lpAddIn->hMod) return;
		lpParams->lpAddInMenu->lpAddIn->pfnCallMenu = GetFunction<LPCALLMENU>(
			lpParams->lpAddInMenu->lpAddIn->hMod,
			szCallMenu);
	}

	if (!lpParams->lpAddInMenu->lpAddIn->pfnCallMenu)
	{
		DebugPrint(DBGAddInPlumbing, L"InvokeAddInMenu: CallMenu not found\n");
		return;
	}

	WC_H(lpParams->lpAddInMenu->lpAddIn->pfnCallMenu(lpParams));
}
#endif // MRMAPI

// Compare type arrays.
int _cdecl CompareTypes(_In_ const void* a1, _In_ const void* a2)
{
	auto lpType1 = LPNAME_ARRAY_ENTRY(a1);
	auto lpType2 = LPNAME_ARRAY_ENTRY(a2);

	if (lpType1->ulValue > lpType2->ulValue) return 1;
	if (lpType1->ulValue == lpType2->ulValue)
	{
		return wcscmp(lpType1->lpszName, lpType2->lpszName);
	}

	return -1;
}

// Compare tag arrays. Pay no attention to sort order - we'll sort on sort order during output.
int _cdecl CompareTags(_In_ const void* a1, _In_ const void* a2)
{
	auto lpTag1 = LPNAME_ARRAY_ENTRY_V2(a1);
	auto lpTag2 = LPNAME_ARRAY_ENTRY_V2(a2);

	if (lpTag1->ulValue > lpTag2->ulValue) return 1;
	if (lpTag1->ulValue == lpTag2->ulValue)
	{
		return wcscmp(lpTag1->lpszName, lpTag2->lpszName);
	}

	return -1;
}

int _cdecl CompareNameID(_In_ const void* a1, _In_ const void* a2)
{
	auto lpID1 = LPNAMEID_ARRAY_ENTRY(a1);
	auto lpID2 = LPNAMEID_ARRAY_ENTRY(a2);

	if (lpID1->lValue > lpID2->lValue) return 1;
	if (lpID1->lValue == lpID2->lValue)
	{
		auto iCmp = wcscmp(lpID1->lpszName, lpID2->lpszName);
		if (iCmp) return iCmp;
		if (IsEqualGUID(*lpID1->lpGuid, *lpID2->lpGuid)) return 0;
	}

	return -1;
}

int _cdecl CompareSmartViewParser(_In_ const void* a1, _In_ const void* a2)
{
	auto lpParser1 = LPSMARTVIEW_PARSER_ARRAY_ENTRY(a1);
	auto lpParser2 = LPSMARTVIEW_PARSER_ARRAY_ENTRY(a2);

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
}

template <typename T> void MergeArrays(
	vector<T> &Target,
	_Inout_bytecap_x_(cSource* width) T* Source,
	_In_ size_t cSource,
	_In_ int(_cdecl *Comparison)(const void *, const void *))
{
	// Sort the source array
	qsort(Source, cSource, sizeof T, Comparison);

	// Append any entries in the source not already in the target to the target
	for (ULONG i = 0; i < cSource; i++)
	{
		if (end(Target) == find_if(begin(Target), end(Target), [&](T &entry)
		{
			return Comparison(&Source[i], &entry) == 0;
		}))
		{
			Target.push_back(Source[i]);
		}
	}

	// Stable sort the resulting array
	std::stable_sort(begin(Target), end(Target), [Comparison](const T& a, const T& b) -> bool
	{
		return Comparison(&a, &b) < 0;
	});
}

// Flags are difficult to sort since we need to have a stable sort
// Records with the same key must appear in the output in the same order as the input
// qsort doesn't guarantee this, so we do it manually with an insertion sort
void SortFlagArray(_In_count_(ulFlags) LPFLAG_ARRAY_ENTRY lpFlags, _In_ ULONG ulFlags)
{
	for (ULONG i = 1; i < ulFlags; i++)
	{
		auto NextItem = lpFlags[i];
		ULONG iLoc = 0;
		for (iLoc = i; iLoc > 0; iLoc--)
		{
			if (lpFlags[iLoc - 1].ulFlagName <= NextItem.ulFlagName) break;
			lpFlags[iLoc] = lpFlags[iLoc - 1];
		}

		lpFlags[iLoc] = NextItem;
	}
}

// Consults the end of the target array to find a match to the source
// If no dupe is found, copies source to target
void AppendFlagIfNotDupe(vector<FLAG_ARRAY_ENTRY>& target, FLAG_ARRAY_ENTRY source)
{
	auto iTarget = target.size();

	while (iTarget)
	{
		iTarget--;
		// Stop searching when ulFlagName doesn't match
		// Assumes lpTarget is sorted
		if (target[iTarget].ulFlagName != source.ulFlagName) break;
		if (target[iTarget].lFlagValue == source.lFlagValue &&
			target[iTarget].ulFlagType == source.ulFlagType &&
			!wcscmp(target[iTarget].lpszName, source.lpszName))
		{
			return;
		}
	}

	target.push_back(source);
}

// Similar to MergeArrays, but using AppendFlagIfNotDupe logic
void MergeFlagArrays(
	vector<FLAG_ARRAY_ENTRY> &In1,
	_In_count_(cIn2) LPFLAG_ARRAY_ENTRY In2,
	_In_ size_t cIn2)
{
	if (!In2) return;

	vector<FLAG_ARRAY_ENTRY> Out;

	size_t iIn1 = 0;
	size_t iIn2 = 0;

	while (iIn1 < In1.size() && iIn2 < cIn2)
	{
		// Add from In1 first, then In2 when In2 is bigger than In2
		if (In1[iIn1].ulFlagName <= In2[iIn2].ulFlagName)
		{
			AppendFlagIfNotDupe(Out, In1[iIn1++]);
		}
		else
		{
			AppendFlagIfNotDupe(Out, In2[iIn2++]);
		}
	}

	while (iIn1 < In1.size())
	{
		AppendFlagIfNotDupe(Out, In1[iIn1++]);
	}

	while (iIn2 < cIn2)
	{
		AppendFlagIfNotDupe(Out, In2[iIn2++]);
	}

	In1 = Out;
}

// Assumes built in arrays are already sorted!
void MergeAddInArrays()
{
	DebugPrint(DBGAddInPlumbing, L"Loading default arrays\n");
	PropTagArray = vector<NAME_ARRAY_ENTRY_V2>(std::begin(g_PropTagArray), std::end(g_PropTagArray));
	PropTypeArray = vector<NAME_ARRAY_ENTRY>(std::begin(g_PropTypeArray), std::end(g_PropTypeArray));
	PropGuidArray = vector<GUID_ARRAY_ENTRY>(std::begin(g_PropGuidArray), std::end(g_PropGuidArray));
	NameIDArray = vector<NAMEID_ARRAY_ENTRY>(std::begin(g_NameIDArray), std::end(g_NameIDArray));
	FlagArray = vector<FLAG_ARRAY_ENTRY>(std::begin(g_FlagArray), std::end(g_FlagArray));
	SmartViewParserArray = vector<SMARTVIEW_PARSER_ARRAY_ENTRY>(std::begin(g_SmartViewParserArray), std::end(g_SmartViewParserArray));
	SmartViewParserTypeArray = vector<NAME_ARRAY_ENTRY>(std::begin(g_SmartViewParserTypeArray), std::end(g_SmartViewParserTypeArray));

	DebugPrint(DBGAddInPlumbing, L"Found 0x%08X built in prop tags.\n", PropTagArray.size());
	DebugPrint(DBGAddInPlumbing, L"Found 0x%08X built in prop types.\n", PropTypeArray.size());
	DebugPrint(DBGAddInPlumbing, L"Found 0x%08X built in guids.\n", PropGuidArray.size());
	DebugPrint(DBGAddInPlumbing, L"Found 0x%08X built in named ids.\n", NameIDArray.size());
	DebugPrint(DBGAddInPlumbing, L"Found 0x%08X built in flags.\n", FlagArray.size());
	DebugPrint(DBGAddInPlumbing, L"Found 0x%08X built in Smart View parsers.\n", SmartViewParserArray.size());
	DebugPrint(DBGAddInPlumbing, L"Found 0x%08X built in Smart View parser types.\n", SmartViewParserTypeArray.size());

	// No add-in == nothing to merge
	if (g_lpMyAddins.empty()) return;

	DebugPrint(DBGAddInPlumbing, L"Merging Add-In arrays\n");
	for (const auto& addIn : g_lpMyAddins)
	{
		DebugPrint(DBGAddInPlumbing, L"Looking at %ws\n", addIn.szName);
		DebugPrint(DBGAddInPlumbing, L"Found 0x%08X prop tags.\n", addIn.ulPropTags);
		DebugPrint(DBGAddInPlumbing, L"Found 0x%08X prop types.\n", addIn.ulPropTypes);
		DebugPrint(DBGAddInPlumbing, L"Found 0x%08X guids.\n", addIn.ulPropGuids);
		DebugPrint(DBGAddInPlumbing, L"Found 0x%08X named ids.\n", addIn.ulNameIDs);
		DebugPrint(DBGAddInPlumbing, L"Found 0x%08X flags.\n", addIn.ulPropFlags);
		DebugPrint(DBGAddInPlumbing, L"Found 0x%08X Smart View parsers.\n", addIn.ulSmartViewParsers);
		DebugPrint(DBGAddInPlumbing, L"Found 0x%08X Smart View parser types.\n", addIn.ulSmartViewParserTypes);
	}

	// Second pass - merge our arrays to the hardcoded arrays
	for (const auto& addIn : g_lpMyAddins)
	{
		if (addIn.ulPropTypes)
		{
			MergeArrays<NAME_ARRAY_ENTRY>(PropTypeArray, addIn.lpPropTypes, addIn.ulPropTypes, CompareTypes);
		}

		if (addIn.ulPropTags)
		{
			MergeArrays<NAME_ARRAY_ENTRY_V2>(PropTagArray, addIn.lpPropTags, addIn.ulPropTags, CompareTags);
		}

		if (addIn.ulNameIDs)
		{
			MergeArrays<NAMEID_ARRAY_ENTRY>(NameIDArray, addIn.lpNameIDs, addIn.ulNameIDs, CompareNameID);
		}

		if (addIn.ulPropFlags)
		{
			SortFlagArray(addIn.lpPropFlags, addIn.ulPropFlags);
			MergeFlagArrays(FlagArray, addIn.lpPropFlags, addIn.ulPropFlags);
		}

		if (addIn.ulSmartViewParsers)
		{
			MergeArrays<SMARTVIEW_PARSER_ARRAY_ENTRY>(SmartViewParserArray, addIn.lpSmartViewParsers, addIn.ulSmartViewParsers, CompareSmartViewParser);
		}

		// We add our new parsers to the end of the array, assigning ids starting with IDS_STEND
		static ULONG s_ulNextParser = IDS_STEND;
		if (addIn.ulSmartViewParserTypes)
		{
			for (ULONG i = 0; i < addIn.ulSmartViewParserTypes; i++)
			{
				NAME_ARRAY_ENTRY addinType;
				addinType.ulValue = s_ulNextParser++;
				addinType.lpszName = addIn.lpSmartViewParserTypes[i];
				SmartViewParserTypeArray.push_back(addinType);
			}
		}

		if (addIn.ulPropGuids)
		{
			// Copy guids from addIn.lpPropGuids, checking for dupes on the way
			for (ULONG i = 0; i < addIn.ulPropGuids; i++)
			{
				auto bDupe = false;
				// Since this array isn't sorted, we have to compare against all valid entries for dupes
				for (const auto& guid : PropGuidArray)
				{
					if (IsEqualGUID(*addIn.lpPropGuids[i].lpGuid, *guid.lpGuid))
					{
						bDupe = true;
						break;
					}
				}

				if (!bDupe)
				{
					PropGuidArray.push_back(addIn.lpPropGuids[i]);
				}
			}
		}
	}

	DebugPrint(DBGAddInPlumbing, L"After merge, 0x%08X prop tags.\n", PropTagArray.size());
	DebugPrint(DBGAddInPlumbing, L"After merge, 0x%08X prop types.\n", PropTypeArray.size());
	DebugPrint(DBGAddInPlumbing, L"After merge, 0x%08X guids.\n", PropGuidArray.size());
	DebugPrint(DBGAddInPlumbing, L"After merge, 0x%08X flags.\n", FlagArray.size());
	DebugPrint(DBGAddInPlumbing, L"After merge, 0x%08X Smart View parsers.\n", SmartViewParserArray.size());
	DebugPrint(DBGAddInPlumbing, L"After merge, 0x%08X Smart View parser types.\n", SmartViewParserTypeArray.size());

	DebugPrint(DBGAddInPlumbing, L"Done merging add-in arrays\n");
}

__declspec(dllexport) void __cdecl AddInLog(bool bPrintThreadTime, _Printf_format_string_ LPWSTR szMsg, ...)
{
	if (!fIsSet(DBGAddIn)) return;

	va_list argList = nullptr;
	va_start(argList, szMsg);
	auto szAddInLogString = formatV(szMsg, argList);
	va_end(argList);

	Output(DBGAddIn, nullptr, bPrintThreadTime, szAddInLogString);
}

#ifndef MRMAPI
_Check_return_ __declspec(dllexport) HRESULT __cdecl SimpleDialog(_In_z_ LPWSTR szTitle, _Printf_format_string_ LPWSTR szMsg, ...)
{
	auto hRes = S_OK;

	CEditor MySimpleDialog(
		nullptr,
		NULL,
		NULL,
		CEDITOR_BUTTON_OK);
	MySimpleDialog.SetAddInTitle(szTitle);

	va_list argList = nullptr;
	va_start(argList, szMsg);
	auto szDialogString = formatV(szMsg, argList);
	va_end(argList);

	MySimpleDialog.SetPromptPostFix(szDialogString);

	WC_H(MySimpleDialog.DisplayDialog());
	return hRes;
}

_Check_return_ __declspec(dllexport) HRESULT __cdecl ComplexDialog(_In_ LPADDINDIALOG lpDialog, _Out_ LPADDINDIALOGRESULT* lppDialogResult)
{
	if (!lpDialog) return MAPI_E_INVALID_PARAMETER;
	// Reject any flags except CEDITOR_BUTTON_OK and CEDITOR_BUTTON_CANCEL
	if (lpDialog->ulButtonFlags & ~(CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL)) return MAPI_E_INVALID_PARAMETER;
	auto hRes = S_OK;

	CEditor MyComplexDialog(
		nullptr,
		NULL,
		NULL,
		lpDialog->ulButtonFlags);
	MyComplexDialog.SetAddInTitle(lpDialog->szTitle);
	MyComplexDialog.SetPromptPostFix(lpDialog->szPrompt);

	if (lpDialog->ulNumControls && lpDialog->lpDialogControls)
	{
		for (ULONG i = 0; i < lpDialog->ulNumControls; i++)
		{
			switch (lpDialog->lpDialogControls[i].cType)
			{
			case ADDIN_CTRL_CHECK:
				MyComplexDialog.InitPane(i, CheckPane::Create(
					NULL,
					lpDialog->lpDialogControls[i].bDefaultCheckState,
					lpDialog->lpDialogControls[i].bReadOnly));
				break;
			case ADDIN_CTRL_EDIT_TEXT:
			{
				if (lpDialog->lpDialogControls[i].bMultiLine)
				{
					MyComplexDialog.InitPane(i, TextPane::CreateCollapsibleTextPane(
						NULL,
						lpDialog->lpDialogControls[i].bReadOnly));
					MyComplexDialog.SetStringW(i, lpDialog->lpDialogControls[i].szDefaultText);
				}
				else
				{
					MyComplexDialog.InitPane(i, TextPane::CreateSingleLinePane(NULL, wstring(lpDialog->lpDialogControls[i].szDefaultText), lpDialog->lpDialogControls[i].bReadOnly));
				}

				break;
			}
			case ADDIN_CTRL_EDIT_BINARY:
				if (lpDialog->lpDialogControls[i].bMultiLine)
				{
					MyComplexDialog.InitPane(i, TextPane::CreateCollapsibleTextPane(
						NULL,
						lpDialog->lpDialogControls[i].bReadOnly));
				}
				else
				{
					MyComplexDialog.InitPane(i, TextPane::CreateSingleLinePane(NULL, lpDialog->lpDialogControls[i].bReadOnly));

				}
				MyComplexDialog.SetBinary(
					i,
					lpDialog->lpDialogControls[i].lpBin,
					lpDialog->lpDialogControls[i].cbBin);
				break;
			case ADDIN_CTRL_EDIT_NUM_DECIMAL:
				MyComplexDialog.InitPane(i, TextPane::CreateSingleLinePane(NULL, lpDialog->lpDialogControls[i].bReadOnly));
				MyComplexDialog.SetDecimal(
					i,
					lpDialog->lpDialogControls[i].ulDefaultNum);
				break;
			case ADDIN_CTRL_EDIT_NUM_HEX:
				MyComplexDialog.InitPane(i, TextPane::CreateSingleLinePane(NULL, lpDialog->lpDialogControls[i].bReadOnly));
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
		auto lpResults = new _AddInDialogResult;
		if (lpResults)
		{
			lpResults->ulNumControls = lpDialog->ulNumControls;
			lpResults->lpDialogControlResults = new _AddInDialogControlResult[lpDialog->ulNumControls];
			if (lpResults->lpDialogControlResults)
			{
				ZeroMemory(lpResults->lpDialogControlResults, sizeof(_AddInDialogControlResult)*lpDialog->ulNumControls);
				for (ULONG i = 0; i < lpDialog->ulNumControls; i++)
				{
					lpResults->lpDialogControlResults[i].cType = lpDialog->lpDialogControls[i].cType;
					switch (lpDialog->lpDialogControls[i].cType)
					{
					case ADDIN_CTRL_CHECK:
						lpResults->lpDialogControlResults[i].bCheckState = MyComplexDialog.GetCheck(i);
						break;
					case ADDIN_CTRL_EDIT_TEXT:
					{
						auto szText = MyComplexDialog.GetStringW(i);
						if (!szText.empty())
						{
							auto cchText = szText.length();

							cchText++;
							lpResults->lpDialogControlResults[i].szText = new WCHAR[cchText];

							if (lpResults->lpDialogControlResults[i].szText)
							{
								EC_H(StringCchCopyW(
									lpResults->lpDialogControlResults[i].szText,
									cchText,
									szText.c_str()));
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
							reinterpret_cast<LPENTRYID*>(&lpResults->lpDialogControlResults[i].lpBin)));
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
}

__declspec(dllexport) void __cdecl FreeDialogResult(_In_ LPADDINDIALOGRESULT lpDialogResult)
{
	if (lpDialogResult)
	{
		if (lpDialogResult->lpDialogControlResults)
		{
			for (ULONG i = 0; i < lpDialogResult->ulNumControls; i++)
			{
				delete[] lpDialogResult->lpDialogControlResults[i].lpBin;
				delete[] lpDialogResult->lpDialogControlResults[i].szText;
			}
		}

		delete[] lpDialogResult;
	}
}
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
}

wstring AddInStructTypeToString(__ParsingTypeEnum iStructType)
{
	for (const auto& smartViewParserType : SmartViewParserTypeArray)
	{
		if (smartViewParserType.ulValue == static_cast<ULONG>(iStructType))
		{
			return smartViewParserType.lpszName;
		}
	}

	return L"";
}

wstring AddInSmartView(__ParsingTypeEnum iStructType, ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	// Don't let add-ins hijack our built in types
	if (iStructType <= IDS_STEND - 1) return L"";

	auto szStructType = AddInStructTypeToString(iStructType);
	if (szStructType.empty()) return L"";

	for (const auto& addIn : g_lpMyAddins)
	{
		if (addIn.ulSmartViewParserTypes)
		{
			for (ULONG i = 0; i < addIn.ulSmartViewParserTypes; i++)
			{
				if (0 == szStructType.compare(addIn.lpSmartViewParserTypes[i]))
				{
					auto pfnSmartViewParse = GetFunction<LPSMARTVIEWPARSE>(addIn.hMod, szSmartViewParse);
					auto pfnFreeParse = GetFunction<LPFREEPARSE>(addIn.hMod, szFreeParse);

					if (pfnSmartViewParse && pfnFreeParse)
					{
						auto szParse = pfnSmartViewParse(szStructType.c_str(), cbBin, lpBin);
						wstring szRet = szParse;

						pfnFreeParse(szParse);
						return szRet;
					}
				}
			}
		}
	}

	return L"No parser found";
}