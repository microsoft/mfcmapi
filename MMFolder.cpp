#include "stdafx.h"
#include "MrMAPI.h"
#include "MMFolder.h"
#include "MAPIFunctions.h"
#include "ExtraPropTags.h"
#include "InterpretProp2.h"
#include "MMStore.h"

// Search folder for entry ID of child folder by name.
HRESULT HrMAPIFindFolderW(
	_In_ LPMAPIFOLDER lpFolder,        // pointer to folder
	_In_z_ LPCWSTR lpszName,           // name of child folder to find
	_Out_opt_ ULONG* lpcbeid,          // pointer to count of bytes in entry ID
	_Deref_out_opt_ LPENTRYID* lppeid) // pointer to entry ID pointer
{
	DebugPrint(DBGGeneric,"HrMAPIFindFolderW: Locating folder \"%ws\"\n", lpszName);
	HRESULT hRes = S_OK;
	LPMAPITABLE lpTable = NULL;
	LPSRowSet lpRow = NULL;
	LPSPropValue lpRowProp = NULL;
	ULONG i = 0;

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
	*lppeid  = NULL;
	if (!lpFolder) return MAPI_E_INVALID_PARAMETER;

	WC_MAPI(lpFolder->GetHierarchyTable(MAPI_UNICODE | MAPI_DEFERRED_ERRORS, &lpTable));
	if (SUCCEEDED(hRes) && lpTable)
	{
		WC_MAPI(HrQueryAllRows(lpTable, (LPSPropTagArray)&rgColProps, NULL, NULL, 0, &lpRow));
	}

	if (SUCCEEDED(hRes) && lpRow)
	{
		for (i = 0; i < lpRow->cRows; i++)
		{
			lpRowProp = lpRow->aRow[i].lpProps;

			if (PR_DISPLAY_NAME_W == lpRowProp[ePR_DISPLAY_NAME_W].ulPropTag &&
				_wcsicmp(lpRowProp[ePR_DISPLAY_NAME_W].Value.lpszW, lpszName) == 0 &&
				PR_ENTRYID == lpRowProp[ePR_ENTRYID].ulPropTag)
			{
				WC_H(MAPIAllocateBuffer(lpRowProp[ePR_ENTRYID].Value.bin.cb, (LPVOID*)lppeid));
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
} // HrMAPIFindFolderW

#define wszBackslash L'\\'

// Splits string lpsz at backslashes and points elements of array
// *lpppsz to string components.
HRESULT HrSplitPath(
	_In_z_ LPCWSTR lpsz,           // separated string
	_Out_opt_ ULONG* lpcpsz,       // count of string pointers
	_Deref_out_opt_z_ LPWSTR** lpppsz) // pointer to list of strings
{
	DebugPrint(DBGGeneric,"HrSplitPath: splitting path \"%ws\"\n", lpsz);
	if (!lpcpsz || !lpppsz) return MAPI_E_INVALID_PARAMETER;
	HRESULT hRes = S_OK;
	LPCWSTR pch = NULL;
	LPBYTE lpb = NULL;
	LPWSTR *lppsz = NULL;
	LPWSTR pchNew = NULL;
	ULONG cToken = 0;
	ULONG cArrayBytes = 0;
	ULONG cBytes = 0;
	ULONG csz = 0;

	*lpcpsz = 0;
	*lpppsz = NULL;

	// Skip to end if nothing to do
	if (*lpsz == 0) return S_OK;

	// Count the number of tokens.
	pch = lpsz;
	while (*pch)
	{
		// Skip adjacent backslashes
		while (wszBackslash == *pch)
			pch++;

		if (*pch == 0) break;

		cToken++;

		// Skip next token
		while (*pch && !(wszBackslash == *pch))
			pch++;
	}

	// Array will consist of pointers followed by the tokenized string
	cArrayBytes = (cToken + 1) * (sizeof(LPWSTR));
	cBytes = cArrayBytes + (lstrlenW(lpsz) + 1) * sizeof(WCHAR);

	WC_H(MAPIAllocateBuffer(cBytes, (LPVOID*) &lpb));
	if (SUCCEEDED(hRes) && lpb)
	{
		lppsz  = (LPWSTR*) lpb;
		pchNew = (LPWSTR) (lpb + cArrayBytes);

		wcscpy_s(pchNew, (cBytes-cArrayBytes), lpsz);

		while (*pchNew)
		{
			// remove and skip backslashes
			while (wszBackslash == *pchNew)
				*pchNew++ = 0;

			if (*pchNew == 0) break;

			lppsz[csz++] = pchNew;

			// skip next token
			while (*pchNew && !(wszBackslash == *pchNew))
				pchNew++;
		}

		lppsz[csz] = NULL;

		*lpcpsz = csz;
		*lpppsz = lppsz;
	}

	if (FAILED(hRes))
		MAPIFreeBuffer(lpb);

	return hRes;
} // HrSplitPath

// Finds an arbitrarily nested folder in the indicated folder given
// a hierarchical list of subfolders.
HRESULT HrMAPIFindSubfolderExW(
	_In_ LPMAPIFOLDER lpRootFolder,    // open root folder
	ULONG ulFolderCount,               // count of hierarchical list of subfolders to navigate
	LPWSTR* lppszFolderList,           // hierarchical list of subfolders to navigate
	_Out_opt_ ULONG* lpcbeid,          // pointer to count of bytes in entry ID
	_Deref_out_opt_ LPENTRYID* lppeid) // pointer to entry ID pointer
{
	HRESULT hRes = S_OK;
	LPMAPIFOLDER lpParentFolder = lpRootFolder;
	LPMAPIFOLDER lpChildFolder = NULL;
	ULONG cbeid = 0;
	LPENTRYID lpeid = NULL;
	ULONG ulObjType = 0;
	ULONG i = 0;

	if (!lpcbeid || !lppeid) return MAPI_E_INVALID_PARAMETER;
	if (!ulFolderCount || !lppszFolderList) return MAPI_E_INVALID_PARAMETER;

	for (i = 0; i < ulFolderCount; i++)
	{
		// Free entryid before re-use.
		MAPIFreeBuffer(lpeid);

		WC_H(HrMAPIFindFolderW(lpParentFolder, lppszFolderList[i], &cbeid, &lpeid));
		if (FAILED(hRes)) break;

		// Only OpenEntry if needed for next tier of folder path.
		if (lppszFolderList[i+1] != NULL)
		{
			WC_MAPI(lpParentFolder->OpenEntry(
				cbeid,
				lpeid,
				NULL,
				MAPI_DEFERRED_ERRORS,
				&ulObjType,
				(LPUNKNOWN*)&lpChildFolder));
			if (FAILED(hRes) || ulObjType != MAPI_FOLDER)
			{
				MAPIFreeBuffer(lpeid);
				hRes = MAPI_E_CALL_FAILED;
				break;
			}
		}

		// No longer need the parent folder
		// (Don't release the folder that was passed!)
		if (i > 0)
		{
			if (lpParentFolder) lpParentFolder->Release();
		}

		lpParentFolder = lpChildFolder;
		lpChildFolder  = NULL;
	}

	// Success!
	*lpcbeid = cbeid;
	*lppeid  = lpeid;

	return hRes;
} // HrMAPIFindSubfolderExW

// Compare folder name to known root folder ENTRYID strings.  Return ENTRYID,
// if matched.
static HRESULT HrLookupRootFolderW(
	_In_ LPMDB lpMDB,                  // pointer to open message store
	_In_z_ LPCWSTR lpszRootFolder,     // root folder name only (no separators)
	_Out_opt_ ULONG* lpcbeid,          // size of entryid
	_Deref_out_opt_ LPENTRYID* lppeid) // pointer to entryid
{
	DebugPrint(DBGGeneric,"HrLookupRootFolderW: Locating root folder \"%ws\"\n", lpszRootFolder);
	if (!lpcbeid || !lppeid) return MAPI_E_INVALID_PARAMETER;
	HRESULT hRes = S_OK;

	*lpcbeid = 0;
	*lppeid  = NULL;

	if (!lpMDB) return MAPI_E_INVALID_PARAMETER;
	// Implicitly recognize no root folder as THE root folder
	if (!lpszRootFolder || !lpszRootFolder[0]) return S_OK;

	ULONG ulPropTag = NULL;

	WC_H(LookupPropName(lpszRootFolder, &ulPropTag));
	if (!ulPropTag)
	{
		// Maybe one of our folder constants was passed.
		// These are base 10.
		ULONG ulFolder = wcstoul(lpszRootFolder,NULL,10);
		if (0 < ulFolder && ulFolder < NUM_DEFAULT_PROPS)
		{
			WC_H(GetDefaultFolderEID(ulFolder, lpMDB, lpcbeid, lppeid));
			return hRes;
		}

		// Still no match?
		// Maybe a prop tag was passed as hex
		ulPropTag = wcstoul(lpszRootFolder,NULL,16);
	}
	if (!ulPropTag) return MAPI_E_NOT_FOUND;

	if (ulPropTag)
	{
		SPropTagArray rgPropTag = { 1, ulPropTag};
		LPSPropValue lpPropValue = NULL;
		ULONG cValues = 0;

		// Get the outbox entry ID property.
		WC_MAPI(lpMDB->GetProps(
			&rgPropTag,
			MAPI_UNICODE,
			&cValues,
			&lpPropValue));
		if (SUCCEEDED(hRes) && lpPropValue && lpPropValue->ulPropTag == ulPropTag)
		{
			WC_H(MAPIAllocateBuffer(lpPropValue->Value.bin.cb, (LPVOID*)lppeid));
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
} // HrLookupRootFolderW

// Finds an arbitrarily nested folder in the indicated store given its
// path name.
HRESULT HrMAPIFindFolderExW(
	_In_ LPMDB lpMDB,                  // Open message store
	_In_z_ LPCWSTR lpszFolderPath,     // folder path
	_Out_opt_ ULONG* lpcbeid,          // pointer to count of bytes in entry ID
	_Deref_out_opt_ LPENTRYID* lppeid) // pointer to entry ID pointer
{
	DebugPrint(DBGGeneric,"HrMAPIFindFolderExW: Locating path \"%ws\"\n", lpszFolderPath);
	HRESULT hRes = S_OK;
	LPMAPIFOLDER lpRootFolder = NULL;
	ULONG cbeid = 0;
	LPENTRYID lpeid = NULL;

	if (!lpcbeid || !lppeid) return MAPI_E_INVALID_PARAMETER;

	*lpcbeid = 0;
	*lppeid  = NULL;

	if (!lpMDB) return MAPI_E_INVALID_PARAMETER;

	ULONG ulFolderCount = 0;
	LPWSTR* lppszFolderList = NULL;
	WC_H(HrSplitPath(lpszFolderPath, &ulFolderCount, &lppszFolderList));

	// Check for literal property name
	if (ulFolderCount && lppszFolderList && lppszFolderList[0] && lppszFolderList[0][0] == L'@')
	{
		WC_H(HrLookupRootFolderW(lpMDB, lppszFolderList[0]+1, &cbeid, &lpeid));
		if (SUCCEEDED(hRes))
		{
			ulFolderCount--;
			if (ulFolderCount)
			{
				lppszFolderList = &lppszFolderList[1];
			}
		}
		else
		{
			// Can we handle the case where one of our constants like DEFAULT_CALENDAR was passed?
			// We can parse the constant and call OpenDefaultFolder, then later not call OpenEntry for the root
		}
	}

	// If we have any subfolders, chase them
	if (SUCCEEDED(hRes) && ulFolderCount)
	{
		ULONG ulObjType = 0;

		WC_MAPI(lpMDB->OpenEntry(
			cbeid,
			lpeid,
			NULL,
			MAPI_BEST_ACCESS | MAPI_DEFERRED_ERRORS,
			&ulObjType,
			(LPUNKNOWN*)&lpRootFolder));
		if (SUCCEEDED(hRes) && MAPI_FOLDER == ulObjType)
		{
			// Free before re-use
			MAPIFreeBuffer(lpeid);

			// Find the subfolder in question
			WC_H(HrMAPIFindSubfolderExW(
				lpRootFolder,
				ulFolderCount,
				lppszFolderList,
				&cbeid,
				&lpeid));
		}
		if (lpRootFolder) lpRootFolder->Release();
	}

	if (SUCCEEDED(hRes))
	{
		*lpcbeid = cbeid;
		*lppeid  = lpeid;
	}

	MAPIFreeBuffer(lppszFolderList);
	return hRes;
} // HrMAPIFindFolderExW

// Opens an arbitrarily nested folder in the indicated store given its
// path name.
HRESULT HrMAPIOpenFolderExW(
	_In_ LPMDB lpMDB,                        // Open message store
	_In_z_ LPCWSTR lpszFolderPath,           // folder path
	_Deref_out_opt_ LPMAPIFOLDER* lppFolder) // pointer to folder opened
{
	DebugPrint(DBGGeneric,"HrMAPIOpenFolderExW: Locating path \"%ws\"\n", lpszFolderPath);
	HRESULT hRes = S_OK;
	LPENTRYID lpeid = NULL;
	ULONG cbeid = 0;
	ULONG ulObjType = 0;

	if (!lppFolder) return MAPI_E_INVALID_PARAMETER;

	*lppFolder = NULL;

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
			(LPUNKNOWN*)lppFolder));
	}

	MAPIFreeBuffer(lpeid);

	return hRes;
} // HrMAPIOpenFolderExW

void DumpHierarchyTable(
	_In_z_ LPWSTR lpszProfile,
	_In_ LPMAPIFOLDER lpFolder,
	_In_ ULONG ulFolder,
	_In_z_ LPWSTR lpszFolder,
	_In_ ULONG ulDepth)
{
	if (0 == ulDepth)
	{
		DebugPrint(DBGGeneric,"DumpHierarchyTable: Outputting hierarchy table for folder %u / %ws from profile %ws \n", ulFolder, lpszFolder?lpszFolder:L"", lpszProfile);
	}
	HRESULT hRes = S_OK;

	if (lpFolder)
	{
		LPMAPITABLE lpTable = NULL;
		LPSRowSet lpRow = NULL;
		ULONG i = 0;

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
			WC_MAPI(lpTable->SetColumns((LPSPropTagArray)&rgColProps, TBL_ASYNC));

			if (!FAILED(hRes)) for (;;)
			{
				hRes = S_OK;
				if (lpRow) FreeProws(lpRow);
				lpRow = NULL;
				WC_MAPI(lpTable->QueryRows(
					50,
					NULL,
					&lpRow));
				if (FAILED(hRes) || !lpRow || !lpRow->cRows) break;

				for (i = 0; i < lpRow->cRows; i++)
				{
					hRes = S_OK;
					if (ulDepth >= 1)
					{
						ULONG iTab = 0;
						for (iTab = 0; iTab < ulDepth; iTab++)
						{
							printf("  ");
						}
					}
					if (PR_DISPLAY_NAME_W == lpRow->aRow[i].lpProps[ePR_DISPLAY_NAME_W].ulPropTag)
					{
						printf("%ws\n",lpRow->aRow[i].lpProps[ePR_DISPLAY_NAME_W].Value.lpszW);
					}

					if (PR_ENTRYID == lpRow->aRow[i].lpProps[ePR_ENTRYID].ulPropTag)
					{
						ULONG ulObjType = NULL;
						LPMAPIFOLDER lpSubfolder = NULL;

						WC_MAPI(lpFolder->OpenEntry(lpRow->aRow[i].lpProps[ePR_ENTRYID].Value.bin.cb,
							(LPENTRYID)lpRow->aRow[i].lpProps[ePR_ENTRYID].Value.bin.lpb,
							NULL,
							MAPI_BEST_ACCESS,
							&ulObjType,
							(LPUNKNOWN *) &lpSubfolder));

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
} // DumpHierarchyTable

ULONGLONG ComputeSingleFolderSize(
	_In_ LPMAPIFOLDER lpFolder)
{
	HRESULT hRes = S_OK;
	LPMAPITABLE lpTable = NULL;
	LPSRowSet lpsRowSet = NULL; 
	LPMAPIFOLDER lpSubfolder = NULL;
	SizedSPropTagArray (1, sProps) = { 1, {PR_MESSAGE_SIZE} };
	ULONGLONG ullThisFolderSize = 0;

	// Look at each item in this folder
	WC_MAPI(lpFolder->GetContentsTable(0, &lpTable));
	if (lpTable)
	{
		WC_MAPI(HrQueryAllRows(lpTable, (LPSPropTagArray)&sProps, NULL, NULL, 0, &lpsRowSet));

		if (lpsRowSet)
		{
			for(ULONG i = 0; i < lpsRowSet->cRows; i++)
			{
				if (PROP_TYPE(lpsRowSet->aRow[i].lpProps[0].ulPropTag) != PT_ERROR)
					ullThisFolderSize += lpsRowSet->aRow[i].lpProps[0].Value.l;
			}
			MAPIFreeBuffer(lpsRowSet);
			lpsRowSet = NULL;
		}
		lpTable->Release();
		lpTable = NULL;
	}
	DebugPrint(DBGGeneric, "Content size = %I64u\n", ullThisFolderSize);
//	printf("Content size = %I64d\n", ullThisFolderSize);

	WC_MAPI(lpFolder->GetContentsTable(MAPI_ASSOCIATED, &lpTable));
	if (lpTable)
	{
		WC_MAPI(HrQueryAllRows(lpTable, (LPSPropTagArray)&sProps, NULL, NULL, 0, &lpsRowSet));

		if (lpsRowSet)
		{
			for(ULONG i = 0; i < lpsRowSet->cRows; i++)
			{
				if (PROP_TYPE(lpsRowSet->aRow[i].lpProps[0].ulPropTag) != PT_ERROR)
					ullThisFolderSize += lpsRowSet->aRow[i].lpProps[0].Value.l;
			}
			MAPIFreeBuffer(lpsRowSet);
			lpsRowSet = NULL;
		}
		lpTable->Release();
		lpTable = NULL;
	}
	DebugPrint(DBGGeneric, "Total size = %I64u\n", ullThisFolderSize);
//	printf("Total size = %I64d\n", ullThisFolderSize);

	return ullThisFolderSize;
} // ComputeSingleFolderSize

ULONGLONG ComputeFolderSize(
	_In_z_ LPWSTR lpszProfile,
	_In_ LPMAPIFOLDER lpFolder,
	_In_ ULONG ulFolder,
	_In_z_ LPWSTR lpszFolder)
{
	DebugPrint(DBGGeneric,"ComputeFolderSize: Calculating size (including subfolders) for folder %u / %ws from profile %ws \n", ulFolder, lpszFolder?lpszFolder:L"", lpszProfile);
//	printf("ComputeFolderSize: Calculating size (including subfolders) for folder %i / %ws from profile %ws \n", ulFolder, lpszFolder?lpszFolder:L"", lpszProfile);
	HRESULT hRes = S_OK;

	if (lpFolder)
	{
		LPMAPITABLE lpTable = NULL;
		LPSRowSet lpRow = NULL;
		LPSPropValue lpPropRootSize = NULL;
		ULONG i = 0;
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
			WC_MAPI(lpTable->SetColumns((LPSPropTagArray)&rgColProps, TBL_ASYNC));

			if (!FAILED(hRes)) for (;;)
			{
				hRes = S_OK;
				if (lpRow) FreeProws(lpRow);
				lpRow = NULL;
				WC_MAPI(lpTable->QueryRows(
					50,
					NULL,
					&lpRow));
				if (FAILED(hRes) || !lpRow || !lpRow->cRows) break;

				for (i = 0; i < lpRow->cRows; i++)
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
						LPMAPIFOLDER lpSubfolder = NULL;

						WC_MAPI(lpFolder->OpenEntry(lpRow->aRow[i].lpProps[ePR_ENTRYID].Value.bin.cb,
							(LPENTRYID)lpRow->aRow[i].lpProps[ePR_ENTRYID].Value.bin.lpb,
							NULL,
							MAPI_BEST_ACCESS,
							&ulObjType,
							(LPUNKNOWN *) &lpSubfolder));

						if (SUCCEEDED(hRes) && lpSubfolder)
						{
							LPWSTR szDisplayName = L"";
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
} // ComputeFolderSize

void DoFolderProps(_In_ MYOPTIONS ProgOpts)
{
	if (ProgOpts.lpFolder)
	{
		ULONG ulPropTag = NULL;
		(void) PropNameToPropTagW(ProgOpts.lpszUnswitchedOption, &ulPropTag);
		PrintObjectProperties(_T("folderprops"), ProgOpts.lpFolder, ulPropTag);
	}
} // DoFolderProps

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
} // DoFolderSize

void DoChildFolders(_In_ MYOPTIONS ProgOpts)
{
	DumpHierarchyTable(
		ProgOpts.lpszProfile,
		ProgOpts.lpFolder,
		ProgOpts.ulFolder,
		ProgOpts.lpszFolderPath,
		0);
} // DoChildFolders