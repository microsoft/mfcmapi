#include "stdafx.h"
#include "MrMAPI.h"
#include "MMFolder.h"
#include "MAPIFunctions.h"
#include "ExtraPropTags.h"
#include "InterpretProp2.h"
#include "MMStore.h"
#include "String.h"

// Search folder for entry ID of child folder by name.
HRESULT HrMAPIFindFolderW(
	_In_ LPMAPIFOLDER lpFolder, // pointer to folder
	_In_ wstring lpszName, // name of child folder to find
	_Out_opt_ ULONG* lpcbeid, // pointer to count of bytes in entry ID
	_Deref_out_opt_ LPENTRYID* lppeid) // pointer to entry ID pointer
{
	DebugPrint(DBGGeneric, L"HrMAPIFindFolderW: Locating folder \"%ws\"\n", lpszName.c_str());
	auto hRes = S_OK;
	LPMAPITABLE lpTable = nullptr;
	LPSRowSet lpRow = nullptr;
	LPSPropValue lpRowProp = nullptr;

	enum
	{
		ePR_DISPLAY_NAME_W,
		ePR_ENTRYID,
		NUM_COLS
	};
	static const SizedSPropTagArray(NUM_COLS, rgColProps) =
	{
		NUM_COLS,
		PR_DISPLAY_NAME_W,
		PR_ENTRYID
	};

	if (!lpcbeid || !lppeid) return MAPI_E_INVALID_PARAMETER;
	*lpcbeid = 0;
	*lppeid = nullptr;
	if (!lpFolder) return MAPI_E_INVALID_PARAMETER;

	WC_MAPI(lpFolder->GetHierarchyTable(MAPI_UNICODE | MAPI_DEFERRED_ERRORS, &lpTable));
	if (SUCCEEDED(hRes) && lpTable)
	{
		WC_MAPI(HrQueryAllRows(lpTable, LPSPropTagArray(&rgColProps), NULL, NULL, 0, &lpRow));
	}

	if (SUCCEEDED(hRes) && lpRow)
	{
		for (ULONG i = 0; i < lpRow->cRows; i++)
		{
			lpRowProp = lpRow->aRow[i].lpProps;

			if (PR_DISPLAY_NAME_W == lpRowProp[ePR_DISPLAY_NAME_W].ulPropTag &&
				_wcsicmp(lpRowProp[ePR_DISPLAY_NAME_W].Value.lpszW, lpszName.c_str()) == 0 &&
				PR_ENTRYID == lpRowProp[ePR_ENTRYID].ulPropTag)
			{
				WC_H(MAPIAllocateBuffer(lpRowProp[ePR_ENTRYID].Value.bin.cb, reinterpret_cast<LPVOID*>(lppeid)));
				if (SUCCEEDED(hRes) && lppeid)
				{
					*lpcbeid = lpRowProp[ePR_ENTRYID].Value.bin.cb;
					CopyMemory(*lppeid, lpRowProp[ePR_ENTRYID].Value.bin.lpb, *lpcbeid);
				}
				break;
			}
		}
	}

	if (!*lpcbeid)
	{
		hRes = MAPI_E_NOT_FOUND;
	}

	FreeProws(lpRow);
	if (lpTable) lpTable->Release();

	return hRes;
}

#define wszBackslash L'\\'
#define wszBackslashStr L"\\"
#define wszDoubleBackslash L"\\\\"

// Converts // to /
wstring unescape(_In_ wstring lpsz)
{
	DebugPrint(DBGGeneric, L"unescape: working on path \"%ws\"\n", lpsz.c_str());

	size_t index = 0;
	while (index != string::npos)
	{
		index = lpsz.find(wszDoubleBackslash, index);
		if (index == string::npos) break;
		lpsz.replace(index, 2, wszBackslashStr);
		index++;
	}

	return lpsz;
}

// Splits string lpsz at backslashes and returns components in vector
vector<wstring> SplitPath(_In_ wstring lpsz)
{
	DebugPrint(DBGGeneric, L"SplitPath: splitting path \"%ws\"\n", lpsz.c_str());

	vector<wstring> result;

	auto str = lpsz.c_str();
	do
	{
		auto begin = str;

		while (*str != 0)
		{
			if (*str == wszBackslash && *(str + 1) == wszBackslash)
			{
				str += 2;
			}
			else if (*str != wszBackslash)
			{
				str++;
			}
			else break;
		}

		if (begin != str)
		{
			result.push_back(unescape(wstring(begin, str)));
		}
	} while (0 != *str++);

	return result;
}

// Finds an arbitrarily nested folder in the indicated folder given
// a hierarchical list of subfolders.
HRESULT HrMAPIFindSubfolderExW(
	_In_ LPMAPIFOLDER lpRootFolder, // open root folder
	vector<wstring> FolderList, // hierarchical list of subfolders to navigate
	_Out_opt_ ULONG* lpcbeid, // pointer to count of bytes in entry ID
	_Deref_out_opt_ LPENTRYID* lppeid) // pointer to entry ID pointer
{
	auto hRes = S_OK;
	ULONG cbeid = 0;
	LPENTRYID lpeid = nullptr;

	if (!lpcbeid || !lppeid) return MAPI_E_INVALID_PARAMETER;
	if (FolderList.empty()) return MAPI_E_INVALID_PARAMETER;

	auto lpParentFolder = lpRootFolder;
	if (lpRootFolder) lpRootFolder->AddRef();

	for (ULONG i = 0; i < FolderList.size(); i++)
	{
		LPMAPIFOLDER lpChildFolder = nullptr;
		ULONG ulObjType = 0;

		// Free entryid before re-use.
		MAPIFreeBuffer(lpeid);

		WC_H(HrMAPIFindFolderW(lpParentFolder, FolderList[i].c_str(), &cbeid, &lpeid));
		if (FAILED(hRes)) break;

		// Only OpenEntry if needed for next tier of folder path.
		if (i + 1 < FolderList.size())
		{
			WC_MAPI(lpParentFolder->OpenEntry(
				cbeid,
				lpeid,
				NULL,
				MAPI_DEFERRED_ERRORS,
				&ulObjType,
				reinterpret_cast<LPUNKNOWN*>(&lpChildFolder)));
			if (FAILED(hRes) || ulObjType != MAPI_FOLDER)
			{
				MAPIFreeBuffer(lpeid);
				hRes = MAPI_E_CALL_FAILED;
				break;
			}
		}

		// No longer need the parent folder
		if (lpParentFolder) lpParentFolder->Release();
		lpParentFolder = lpChildFolder;
		lpChildFolder = nullptr;
	}

	// Success!
	*lpcbeid = cbeid;
	*lppeid = lpeid;

	if (lpParentFolder) lpParentFolder->Release();

	return hRes;
}

// Compare folder name to known root folder ENTRYID strings.  Return ENTRYID,
// if matched.
static HRESULT HrLookupRootFolderW(
	_In_ LPMDB lpMDB, // pointer to open message store
	_In_ wstring lpszRootFolder, // root folder name only (no separators)
	_Out_opt_ ULONG* lpcbeid, // size of entryid
	_Deref_out_opt_ LPENTRYID* lppeid) // pointer to entryid
{
	DebugPrint(DBGGeneric, L"HrLookupRootFolderW: Locating root folder \"%ws\"\n", lpszRootFolder.c_str());
	if (!lpcbeid || !lppeid) return MAPI_E_INVALID_PARAMETER;
	auto hRes = S_OK;

	*lpcbeid = 0;
	*lppeid = nullptr;

	if (!lpMDB) return MAPI_E_INVALID_PARAMETER;
	// Implicitly recognize no root folder as THE root folder
	if (lpszRootFolder.empty()) return S_OK;

	auto ulPropTag = LookupPropName(lpszRootFolder);
	if (!ulPropTag)
	{
		// Maybe one of our folder constants was passed.
		// These are base 10.
		auto ulFolder = wstringToUlong(lpszRootFolder, 10);
		if (0 < ulFolder && ulFolder < NUM_DEFAULT_PROPS)
		{
			WC_H(GetDefaultFolderEID(ulFolder, lpMDB, lpcbeid, lppeid));
			return hRes;
		}

		// Still no match?
		// Maybe a prop tag was passed as hex
		ulPropTag = wstringToUlong(lpszRootFolder, 16);
	}

	if (!ulPropTag) return MAPI_E_NOT_FOUND;

	if (ulPropTag)
	{
		SPropTagArray rgPropTag = { 1, ulPropTag };
		LPSPropValue lpPropValue = nullptr;
		ULONG cValues = 0;

		// Get the outbox entry ID property.
		WC_MAPI(lpMDB->GetProps(
			&rgPropTag,
			MAPI_UNICODE,
			&cValues,
			&lpPropValue));
		if (SUCCEEDED(hRes) && lpPropValue && lpPropValue->ulPropTag == ulPropTag)
		{
			WC_H(MAPIAllocateBuffer(lpPropValue->Value.bin.cb, reinterpret_cast<LPVOID*>(lppeid)));
			if (SUCCEEDED(hRes))
			{
				*lpcbeid = lpPropValue->Value.bin.cb;
				CopyMemory(*lppeid, lpPropValue->Value.bin.lpb, *lpcbeid);
			}
		}
		else
		{
			if (SUCCEEDED(hRes))
			{
				hRes = MAPI_E_CALL_FAILED;
			}
		}

		MAPIFreeBuffer(lpPropValue);
	}

	return hRes;
}

// Finds an arbitrarily nested folder in the indicated store given its
// path name.
HRESULT HrMAPIFindFolderExW(
	_In_ LPMDB lpMDB, // Open message store
	_In_ wstring lpszFolderPath, // folder path
	_Out_opt_ ULONG* lpcbeid, // pointer to count of bytes in entry ID
	_Deref_out_opt_ LPENTRYID* lppeid) // pointer to entry ID pointer
{
	DebugPrint(DBGGeneric, L"HrMAPIFindFolderExW: Locating path \"%ws\"\n", lpszFolderPath.c_str());
	auto hRes = S_OK;
	LPMAPIFOLDER lpRootFolder = nullptr;
	ULONG cbeid = 0;
	LPENTRYID lpeid = nullptr;

	if (!lpcbeid || !lppeid) return MAPI_E_INVALID_PARAMETER;

	*lpcbeid = 0;
	*lppeid = nullptr;

	if (!lpMDB) return MAPI_E_INVALID_PARAMETER;

	auto FolderList = SplitPath(lpszFolderPath);

	// Check for literal property name
	if (!FolderList.empty() && FolderList[0][0] == L'@')
	{
		WC_H(HrLookupRootFolderW(lpMDB, FolderList[0].c_str() + 1, &cbeid, &lpeid));
		if (SUCCEEDED(hRes))
		{
			FolderList.erase(FolderList.begin());
		}
		else
		{
			// Can we handle the case where one of our constants like DEFAULT_CALENDAR was passed?
			// We can parse the constant and call OpenDefaultFolder, then later not call OpenEntry for the root
		}
	}

	// If we have any subfolders, chase them
	if (SUCCEEDED(hRes) && !FolderList.empty())
	{
		ULONG ulObjType = 0;

		WC_MAPI(lpMDB->OpenEntry(
			cbeid,
			lpeid,
			NULL,
			MAPI_BEST_ACCESS | MAPI_DEFERRED_ERRORS,
			&ulObjType,
			reinterpret_cast<LPUNKNOWN*>(&lpRootFolder)));
		if (SUCCEEDED(hRes) && MAPI_FOLDER == ulObjType)
		{
			// Free before re-use
			MAPIFreeBuffer(lpeid);

			// Find the subfolder in question
			WC_H(HrMAPIFindSubfolderExW(
				lpRootFolder,
				FolderList,
				&cbeid,
				&lpeid));
		}
		if (lpRootFolder) lpRootFolder->Release();
	}

	if (SUCCEEDED(hRes))
	{
		*lpcbeid = cbeid;
		*lppeid = lpeid;
	}

	return hRes;
}

// Opens an arbitrarily nested folder in the indicated store given its
// path name.
HRESULT HrMAPIOpenFolderExW(
	_In_ LPMDB lpMDB, // Open message store
	_In_ wstring lpszFolderPath, // folder path
	_Deref_out_opt_ LPMAPIFOLDER* lppFolder) // pointer to folder opened
{
	DebugPrint(DBGGeneric, L"HrMAPIOpenFolderExW: Locating path \"%ws\"\n", lpszFolderPath.c_str());
	auto hRes = S_OK;
	LPENTRYID lpeid = nullptr;
	ULONG cbeid = 0;
	ULONG ulObjType = 0;

	if (!lppFolder) return MAPI_E_INVALID_PARAMETER;

	*lppFolder = nullptr;

	WC_H(HrMAPIFindFolderExW(
		lpMDB,
		lpszFolderPath,
		&cbeid,
		&lpeid));

	if (SUCCEEDED(hRes))
	{
		WC_MAPI(lpMDB->OpenEntry(
			cbeid,
			lpeid,
			NULL,
			MAPI_BEST_ACCESS | MAPI_DEFERRED_ERRORS,
			&ulObjType,
			reinterpret_cast<LPUNKNOWN*>(lppFolder)));
	}

	MAPIFreeBuffer(lpeid);

	return hRes;
}

void DumpHierarchyTable(
	_In_ wstring lpszProfile,
	_In_ LPMAPIFOLDER lpFolder,
	_In_ ULONG ulFolder,
	_In_ wstring lpszFolder,
	_In_ ULONG ulDepth)
{
	if (0 == ulDepth)
	{
		DebugPrint(DBGGeneric, L"DumpHierarchyTable: Outputting hierarchy table for folder %u / %ws from profile %ws \n", ulFolder, lpszFolder.c_str(), lpszProfile.c_str());
	}
	auto hRes = S_OK;

	if (lpFolder)
	{
		LPMAPITABLE lpTable = nullptr;
		LPSRowSet lpRow = nullptr;

		enum
		{
			ePR_DISPLAY_NAME_W,
			ePR_ENTRYID,
			NUM_COLS
		};
		static const SizedSPropTagArray(NUM_COLS, rgColProps) =
		{
			NUM_COLS,
			PR_DISPLAY_NAME_W,
			PR_ENTRYID,
		};

		WC_MAPI(lpFolder->GetHierarchyTable(MAPI_DEFERRED_ERRORS, &lpTable));
		if (SUCCEEDED(hRes) && lpTable)
		{
			WC_MAPI(lpTable->SetColumns(LPSPropTagArray(&rgColProps), TBL_ASYNC));

			if (!FAILED(hRes)) for (;;)
			{
				hRes = S_OK;
				if (lpRow) FreeProws(lpRow);
				lpRow = nullptr;
				WC_MAPI(lpTable->QueryRows(
					50,
					NULL,
					&lpRow));
				if (FAILED(hRes) || !lpRow || !lpRow->cRows) break;

				for (ULONG i = 0; i < lpRow->cRows; i++)
				{
					hRes = S_OK;
					if (ulDepth >= 1)
					{
						for (ULONG iTab = 0; iTab < ulDepth; iTab++)
						{
							printf("  ");
						}
					}
					if (PR_DISPLAY_NAME_W == lpRow->aRow[i].lpProps[ePR_DISPLAY_NAME_W].ulPropTag)
					{
						printf("%ws\n", lpRow->aRow[i].lpProps[ePR_DISPLAY_NAME_W].Value.lpszW);
					}

					if (PR_ENTRYID == lpRow->aRow[i].lpProps[ePR_ENTRYID].ulPropTag)
					{
						ULONG ulObjType = NULL;
						LPMAPIFOLDER lpSubfolder = nullptr;

						WC_MAPI(lpFolder->OpenEntry(lpRow->aRow[i].lpProps[ePR_ENTRYID].Value.bin.cb,
							reinterpret_cast<LPENTRYID>(lpRow->aRow[i].lpProps[ePR_ENTRYID].Value.bin.lpb),
							NULL,
							MAPI_BEST_ACCESS,
							&ulObjType,
							reinterpret_cast<LPUNKNOWN *>(&lpSubfolder)));

						if (SUCCEEDED(hRes) && lpSubfolder)
						{
							DumpHierarchyTable(lpszProfile, lpSubfolder, 0, L"", ulDepth + 1);
						}

						if (lpSubfolder) lpSubfolder->Release();
					}
				}
			}

			if (lpRow) FreeProws(lpRow);
		}

		if (lpTable) lpTable->Release();
	}
}

ULONGLONG ComputeSingleFolderSize(
	_In_ LPMAPIFOLDER lpFolder)
{
	auto hRes = S_OK;
	LPMAPITABLE lpTable = nullptr;
	LPSRowSet lpsRowSet = nullptr;
	SizedSPropTagArray(1, sProps) = { 1, { PR_MESSAGE_SIZE } };
	ULONGLONG ullThisFolderSize = 0;

	// Look at each item in this folder
	WC_MAPI(lpFolder->GetContentsTable(0, &lpTable));
	if (lpTable)
	{
		WC_MAPI(HrQueryAllRows(lpTable, reinterpret_cast<LPSPropTagArray>(&sProps), NULL, NULL, 0, &lpsRowSet));

		if (lpsRowSet)
		{
			for (ULONG i = 0; i < lpsRowSet->cRows; i++)
			{
				if (PROP_TYPE(lpsRowSet->aRow[i].lpProps[0].ulPropTag) != PT_ERROR)
					ullThisFolderSize += lpsRowSet->aRow[i].lpProps[0].Value.l;
			}
			MAPIFreeBuffer(lpsRowSet);
			lpsRowSet = nullptr;
		}

		lpTable->Release();
		lpTable = nullptr;
	}
	DebugPrint(DBGGeneric, L"Content size = %I64u\n", ullThisFolderSize);

	WC_MAPI(lpFolder->GetContentsTable(MAPI_ASSOCIATED, &lpTable));
	if (lpTable)
	{
		WC_MAPI(HrQueryAllRows(lpTable, reinterpret_cast<LPSPropTagArray>(&sProps), NULL, NULL, 0, &lpsRowSet));

		if (lpsRowSet)
		{
			for (ULONG i = 0; i < lpsRowSet->cRows; i++)
			{
				if (PROP_TYPE(lpsRowSet->aRow[i].lpProps[0].ulPropTag) != PT_ERROR)
					ullThisFolderSize += lpsRowSet->aRow[i].lpProps[0].Value.l;
			}
			MAPIFreeBuffer(lpsRowSet);
			lpsRowSet = nullptr;
		}

		lpTable->Release();
		lpTable = nullptr;
	}

	DebugPrint(DBGGeneric, L"Total size = %I64u\n", ullThisFolderSize);

	return ullThisFolderSize;
}

ULONGLONG ComputeFolderSize(
	_In_ wstring lpszProfile,
	_In_ LPMAPIFOLDER lpFolder,
	_In_ ULONG ulFolder,
	_In_ wstring lpszFolder)
{
	DebugPrint(DBGGeneric, L"ComputeFolderSize: Calculating size (including subfolders) for folder %u / %ws from profile %ws \n", ulFolder, lpszFolder.c_str(), lpszProfile.c_str());
	auto hRes = S_OK;

	if (lpFolder)
	{
		LPMAPITABLE lpTable = nullptr;
		LPSRowSet lpRow = nullptr;
		ULONGLONG ullSize = 0;

		enum
		{
			ePR_DISPLAY_NAME_W,
			ePR_ENTRYID,
			ePR_FOLDER_TYPE,
			NUM_COLS
		};
		static const SizedSPropTagArray(NUM_COLS, rgColProps) =
		{
			NUM_COLS,
			PR_DISPLAY_NAME_W,
			PR_ENTRYID,
			PR_FOLDER_TYPE,
		};

		// Size of this folder
		ullSize += ComputeSingleFolderSize(lpFolder);

		// Size of children
		WC_MAPI(lpFolder->GetHierarchyTable(MAPI_DEFERRED_ERRORS, &lpTable));
		if (SUCCEEDED(hRes) && lpTable)
		{
			WC_MAPI(lpTable->SetColumns(LPSPropTagArray(&rgColProps), TBL_ASYNC));

			if (!FAILED(hRes)) for (;;)
			{
				hRes = S_OK;
				if (lpRow) FreeProws(lpRow);
				lpRow = nullptr;
				WC_MAPI(lpTable->QueryRows(
					50,
					NULL,
					&lpRow));
				if (FAILED(hRes) || !lpRow || !lpRow->cRows) break;

				for (ULONG i = 0; i < lpRow->cRows; i++)
				{
					hRes = S_OK;
					// Don't look at search folders
					if (PR_FOLDER_TYPE == lpRow->aRow[i].lpProps[ePR_FOLDER_TYPE].ulPropTag && FOLDER_SEARCH == lpRow->aRow[i].lpProps[ePR_FOLDER_TYPE].Value.ul)
					{
						continue;
					}

					if (PR_ENTRYID == lpRow->aRow[i].lpProps[ePR_ENTRYID].ulPropTag)
					{
						ULONG ulObjType = NULL;
						LPMAPIFOLDER lpSubfolder = nullptr;

						WC_MAPI(lpFolder->OpenEntry(lpRow->aRow[i].lpProps[ePR_ENTRYID].Value.bin.cb,
							reinterpret_cast<LPENTRYID>(lpRow->aRow[i].lpProps[ePR_ENTRYID].Value.bin.lpb),
							NULL,
							MAPI_BEST_ACCESS,
							&ulObjType,
							reinterpret_cast<LPUNKNOWN *>(&lpSubfolder)));

						if (SUCCEEDED(hRes) && lpSubfolder)
						{
							wstring szDisplayName;
							if (PR_DISPLAY_NAME_W == lpRow->aRow[i].lpProps[ePR_DISPLAY_NAME_W].ulPropTag)
							{
								szDisplayName = lpRow->aRow[i].lpProps[ePR_DISPLAY_NAME_W].Value.lpszW;
							}

							ullSize += ComputeFolderSize(lpszProfile, lpSubfolder, 0, szDisplayName);
						}

						if (lpSubfolder) lpSubfolder->Release();
					}
				}
			}

			if (lpRow) FreeProws(lpRow);
		}

		if (lpTable) lpTable->Release();

		return ullSize;
	}

	return 0;
}

void DumpSearchState(
	_In_ wstring lpszProfile,
	_In_ LPMAPIFOLDER lpFolder,
	_In_ ULONG ulFolder,
	_In_ wstring lpszFolder)
{
	DebugPrint(DBGGeneric, L"DumpSearchState: Outputting search state for folder %u / %ws from profile %ws \n", ulFolder, lpszFolder.c_str(), lpszProfile.c_str());

	auto hRes = S_OK;

	if (lpFolder)
	{
		LPSRestriction lpRes = nullptr;
		LPENTRYLIST lpEntryList = nullptr;
		ULONG ulSearchState = 0;

		WC_H(lpFolder->GetSearchCriteria(
			fMapiUnicode,
			&lpRes,
			&lpEntryList,
			&ulSearchState));
		if (MAPI_E_NOT_INITIALIZED == hRes)
		{
			printf("No search criteria has been set on this folder.\n");
		}
		else if (MAPI_E_NO_SUPPORT == hRes)
		{
			printf("This does not appear to be a search folder.\n");
		}
		else if (SUCCEEDED(hRes))
		{
			auto szFlags = InterpretFlags(flagSearchState, ulSearchState);
			printf("Search state %ws == 0x%08X\n", szFlags.c_str(), ulSearchState);
			printf("\n");
			printf("Search Scope:\n");
			_OutputEntryList(DBGNoDebug, stdout, lpEntryList);
			printf("\n");
			printf("Search Criteria:\n");
			_OutputRestriction(DBGNoDebug, stdout, lpRes, nullptr);
		}

		MAPIFreeBuffer(lpRes);
		MAPIFreeBuffer(lpEntryList);
	}
}

void DoFolderProps(_In_ MYOPTIONS ProgOpts)
{
	if (ProgOpts.lpFolder)
	{
		PrintObjectProperties(L"folderprops", ProgOpts.lpFolder, PropNameToPropTag(ProgOpts.lpszUnswitchedOption));
	}
}

void DoFolderSize(_In_ MYOPTIONS ProgOpts)
{
	LONGLONG ullSize = ComputeFolderSize(
		ProgOpts.lpszProfile,
		ProgOpts.lpFolder,
		ProgOpts.ulFolder,
		ProgOpts.lpszFolderPath);
	printf("Folder size (including subfolders)\n");
	printf("Bytes: %I64d\n", ullSize);
	printf("KB: %I64d\n", ullSize / 1024);
	printf("MB: %I64d\n", ullSize / (1024 * 1024));
}

void DoChildFolders(_In_ MYOPTIONS ProgOpts)
{
	DumpHierarchyTable(
		!ProgOpts.lpszProfile.empty() ? ProgOpts.lpszProfile.c_str() : L"",
		ProgOpts.lpFolder,
		ProgOpts.ulFolder,
		!ProgOpts.lpszFolderPath.empty() ? ProgOpts.lpszFolderPath.c_str() : L"",
		0);
}

void DoSearchState(_In_ MYOPTIONS ProgOpts)
{
	DumpSearchState(
		!ProgOpts.lpszProfile.empty() ? ProgOpts.lpszProfile.c_str() : L"",
		ProgOpts.lpFolder,
		ProgOpts.ulFolder,
		!ProgOpts.lpszFolderPath.empty() ? ProgOpts.lpszFolderPath.c_str() : L"");
}