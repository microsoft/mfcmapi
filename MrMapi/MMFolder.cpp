#include <StdAfx.h>
#include <MrMapi/MMFolder.h>
#include <MrMapi/mmcli.h>
#include <core/mapi/extraPropTags.h>
#include <MrMapi/MMStore.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/interpret/flags.h>
#include <core/mapi/mapiOutput.h>
#include <core/interpret/proptags.h>
#include <core/mapi/mapiFunctions.h>

// Search folder for entry ID of child folder by name.
LPSBinary MAPIFindFolderW(
	_In_opt_ LPMAPIFOLDER lpFolder, // pointer to folder
	_In_ const std::wstring& lpszName) // name of child folder to find
{
	output::DebugPrint(output::dbgLevel::Generic, L"MAPIFindFolderW: Locating folder \"%ws\"\n", lpszName.c_str());
	LPMAPITABLE lpTable = nullptr;
	LPSRowSet lpRow = nullptr;
	LPSPropValue lpRowProp = nullptr;

	enum
	{
		ePR_DISPLAY_NAME_W,
		ePR_ENTRYID,
		NUM_COLS
	};
	static const SizedSPropTagArray(NUM_COLS, rgColProps) = {NUM_COLS, PR_DISPLAY_NAME_W, PR_ENTRYID};

	if (!lpFolder) return {};

	WC_MAPI_S(lpFolder->GetHierarchyTable(MAPI_UNICODE | MAPI_DEFERRED_ERRORS, &lpTable));
	if (lpTable)
	{
		WC_MAPI_S(HrQueryAllRows(lpTable, LPSPropTagArray(&rgColProps), nullptr, nullptr, 0, &lpRow));
	}

	auto eid = LPSBinary{};
	if (lpRow)
	{
		for (ULONG i = 0; i < lpRow->cRows; i++)
		{
			lpRowProp = lpRow->aRow[i].lpProps;

			if (PR_DISPLAY_NAME_W == lpRowProp[ePR_DISPLAY_NAME_W].ulPropTag &&
				_wcsicmp(lpRowProp[ePR_DISPLAY_NAME_W].Value.lpszW, lpszName.c_str()) == 0 &&
				PR_ENTRYID == lpRowProp[ePR_ENTRYID].ulPropTag)
			{
				eid = mapi::CopySBinary(&mapi::getBin(lpRowProp[ePR_ENTRYID]));
				break;
			}
		}
	}

	FreeProws(lpRow);
	if (lpTable) lpTable->Release();

	return eid;
}

#define wszBackslash L'\\'
#define wszBackslashStr L"\\"
#define wszDoubleBackslash L"\\\\"

// Converts // to /
std::wstring unescape(_In_ std::wstring lpsz)
{
	output::DebugPrint(output::dbgLevel::Generic, L"unescape: working on path \"%ws\"\n", lpsz.c_str());

	size_t index = 0;
	while (index != std::string::npos)
	{
		index = lpsz.find(wszDoubleBackslash, index);
		if (index == std::string::npos) break;
		lpsz.replace(index, 2, wszBackslashStr);
		index++;
	}

	return lpsz;
}

// Finds an arbitrarily nested folder in the indicated folder given
// a hierarchical list of subfolders.
LPSBinary MAPIFindSubfolderExW(
	_In_ LPMAPIFOLDER lpRootFolder, // open root folder
	const std::vector<std::wstring>& FolderList) // hierarchical list of subfolders to navigate
{
	if (FolderList.empty()) return {};

	auto lpParentFolder = lpRootFolder;
	if (lpRootFolder) lpRootFolder->AddRef();

	auto eid = LPSBinary{};
	for (ULONG i = 0; i < FolderList.size(); i++)
	{
		LPMAPIFOLDER lpChildFolder = nullptr;
		ULONG ulObjType = 0;

		// Free entryid before re-use.
		MAPIFreeBuffer(eid);

		eid = MAPIFindFolderW(lpParentFolder, FolderList[i].c_str());
		if (!eid) break;

		// Only OpenEntry if needed for next tier of folder path.
		if ((size_t) i + 1 < FolderList.size())
		{
			WC_MAPI_S(lpParentFolder->OpenEntry(
				eid->cb,
				reinterpret_cast<LPENTRYID>(eid->lpb),
				nullptr,
				MAPI_DEFERRED_ERRORS,
				&ulObjType,
				reinterpret_cast<LPUNKNOWN*>(&lpChildFolder)));
			if (ulObjType != MAPI_FOLDER)
			{
				MAPIFreeBuffer(eid);
				eid = {};
				break;
			}
		}

		// No longer need the parent folder
		if (lpParentFolder) lpParentFolder->Release();
		lpParentFolder = lpChildFolder;
		lpChildFolder = nullptr;
	}

	// Success!
	if (lpParentFolder) lpParentFolder->Release();

	return eid;
}

// Compare folder name to known root folder ENTRYID strings.  Return ENTRYID,
// if matched.
static LPSBinary LookupRootFolderW(
	_In_ LPMDB lpMDB, // pointer to open message store
	_In_ const std::wstring& lpszRootFolder) // root folder name only (no separators)
{
	output::DebugPrint(
		output::dbgLevel::Generic, L"LookupRootFolderW: Locating root folder \"%ws\"\n", lpszRootFolder.c_str());

	if (!lpMDB) return {};
	// Implicitly recognize no root folder as THE root folder
	if (lpszRootFolder.empty())
	{
		auto rootEid = SBinary{};
		return mapi::CopySBinary(&rootEid);
	}

	auto eid = LPSBinary{};
	auto ulPropTag = proptags::LookupPropName(lpszRootFolder);
	if (!ulPropTag)
	{
		// Maybe one of our folder constants was passed.
		// These are base 10.
		const auto ulFolder = strings::wstringToUlong(lpszRootFolder, 10);
		if (0 < ulFolder && ulFolder < mapi::NUM_DEFAULT_PROPS)
		{
			return mapi::GetDefaultFolderEID(ulFolder, lpMDB);
		}

		// Still no match?
		// Maybe a prop tag was passed as hex
		ulPropTag = strings::wstringToUlong(lpszRootFolder, 16);
	}

	if (!ulPropTag) return {};

	if (ulPropTag)
	{
		auto rgPropTag = SPropTagArray{1, ulPropTag};
		LPSPropValue lpPropValue = nullptr;
		ULONG cValues = 0;

		// Get the outbox entry ID property.
		WC_MAPI_S(lpMDB->GetProps(&rgPropTag, MAPI_UNICODE, &cValues, &lpPropValue));
		if (lpPropValue && lpPropValue->ulPropTag == ulPropTag)
		{
			eid = mapi::CopySBinary(&mapi::getBin(lpPropValue));
		}

		MAPIFreeBuffer(lpPropValue);
	}

	return eid;
}

// Finds an arbitrarily nested folder in the indicated store given its
// path name.
LPSBinary MAPIFindFolderExW(
	_In_opt_ LPMDB lpMDB, // Open message store
	_In_ const std::wstring& lpszFolderPath) // folder path
{
	output::DebugPrint(
		output::dbgLevel::Generic, L"MAPIFindFolderExW: Locating path \"%ws\"\n", lpszFolderPath.c_str());
	if (!lpMDB) return {};

	LPMAPIFOLDER lpRootFolder = nullptr;
	auto FolderList = strings::split(lpszFolderPath, wszBackslash);

	// Check for literal property name
	auto eid = LPSBinary{};
	if (!FolderList.empty() && FolderList[0][0] == L'@')
	{
		eid = LookupRootFolderW(lpMDB, FolderList[0].c_str() + 1);
		if (eid)
		{
			FolderList.erase(FolderList.begin());
		}
		else
		{
			// Can we handle the case where one of our constants like DEFAULT_CALENDAR was passed?
			// We can parse the constant and call OpenDefaultFolder, then later not call OpenEntry for the root
		}
	}

	// If we don't have an entry ID yet, start from the root
	if (!eid)
	{
		auto rootEid = SBinary{};
		eid = mapi::CopySBinary(&rootEid);
	}

	// If we have any subfolders, chase them
	if (!FolderList.empty() && eid)
	{
		ULONG ulObjType = 0;

		WC_MAPI_S(lpMDB->OpenEntry(
			eid->cb,
			reinterpret_cast<LPENTRYID>(eid->lpb),
			nullptr,
			MAPI_BEST_ACCESS | MAPI_DEFERRED_ERRORS,
			&ulObjType,
			reinterpret_cast<LPUNKNOWN*>(&lpRootFolder)));
		if (MAPI_FOLDER == ulObjType)
		{
			// Free before re-use
			MAPIFreeBuffer(eid);

			// Find the subfolder in question
			eid = MAPIFindSubfolderExW(lpRootFolder, FolderList);
		}

		if (lpRootFolder) lpRootFolder->Release();
	}

	return eid;
}

// Opens an arbitrarily nested folder in the indicated store given its
// path name.
LPMAPIFOLDER MAPIOpenFolderExW(
	_In_opt_ LPMDB lpMDB, // Open message store
	_In_ const std::wstring& lpszFolderPath) // folder path
{
	output::DebugPrint(
		output::dbgLevel::Generic, L"MAPIOpenFolderExW: Locating path \"%ws\"\n", lpszFolderPath.c_str());
	ULONG ulObjType = 0;

	auto eid = MAPIFindFolderExW(lpMDB, lpszFolderPath);

	LPMAPIFOLDER lpFolder = nullptr;
	if (eid)
	{
		WC_MAPI_S(lpMDB->OpenEntry(
			eid->cb,
			reinterpret_cast<LPENTRYID>(eid->lpb),
			nullptr,
			MAPI_BEST_ACCESS | MAPI_DEFERRED_ERRORS,
			&ulObjType,
			reinterpret_cast<LPUNKNOWN*>(&lpFolder)));
	}

	MAPIFreeBuffer(eid);

	return lpFolder;
}

void DumpHierarchyTable(
	_In_ const std::wstring& lpszProfile,
	_In_opt_ LPMAPIFOLDER lpFolder,
	_In_ const std::wstring& lpszFolder,
	_In_ ULONG ulDepth)
{
	if (0 == ulDepth)
	{
		output::DebugPrint(
			output::dbgLevel::Generic,
			L"DumpHierarchyTable: Outputting hierarchy table for folder %ws from profile %ws \n",
			lpszFolder.c_str(),
			lpszProfile.c_str());
	}

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
		static const SizedSPropTagArray(NUM_COLS, rgColProps) = {
			NUM_COLS,
			PR_DISPLAY_NAME_W,
			PR_ENTRYID,
		};

		WC_MAPI_S(lpFolder->GetHierarchyTable(MAPI_DEFERRED_ERRORS, &lpTable));
		if (lpTable)
		{
			WC_MAPI_S(lpTable->SetColumns(LPSPropTagArray(&rgColProps), TBL_ASYNC));

			for (;;)
			{
				if (lpRow) FreeProws(lpRow);
				lpRow = nullptr;
				WC_MAPI_S(lpTable->QueryRows(50, NULL, &lpRow));
				if (!lpRow || !lpRow->cRows) break;

				for (ULONG i = 0; i < lpRow->cRows; i++)
				{
					if (ulDepth >= 1)
					{
						for (ULONG iTab = 0; iTab < ulDepth; iTab++)
						{
							wprintf(L"  ");
						}
					}

					if (PR_DISPLAY_NAME_W == lpRow->aRow[i].lpProps[ePR_DISPLAY_NAME_W].ulPropTag)
					{
						wprintf(L"%ws\n", lpRow->aRow[i].lpProps[ePR_DISPLAY_NAME_W].Value.lpszW);
					}

					if (PR_ENTRYID == lpRow->aRow[i].lpProps[ePR_ENTRYID].ulPropTag)
					{
						ULONG ulObjType = NULL;
						LPMAPIFOLDER lpSubfolder = nullptr;

						WC_MAPI_S(lpFolder->OpenEntry(
							mapi::getBin(lpRow->aRow[i].lpProps[ePR_ENTRYID]).cb,
							reinterpret_cast<LPENTRYID>(mapi::getBin(lpRow->aRow[i].lpProps[ePR_ENTRYID]).lpb),
							nullptr,
							MAPI_BEST_ACCESS,
							&ulObjType,
							reinterpret_cast<LPUNKNOWN*>(&lpSubfolder)));

						if (lpSubfolder)
						{
							DumpHierarchyTable(lpszProfile, lpSubfolder, L"", ulDepth + 1);
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

ULONGLONG ComputeSingleFolderSize(_In_ LPMAPIFOLDER lpFolder)
{
	LPMAPITABLE lpTable = nullptr;
	LPSRowSet lpsRowSet = nullptr;
	SizedSPropTagArray(1, sProps) = {1, {PR_MESSAGE_SIZE}};
	ULONGLONG ullThisFolderSize = 0;

	// Look at each item in this folder
	WC_MAPI_S(lpFolder->GetContentsTable(0, &lpTable));
	if (lpTable)
	{
		WC_MAPI_S(HrQueryAllRows(lpTable, reinterpret_cast<LPSPropTagArray>(&sProps), nullptr, nullptr, 0, &lpsRowSet));

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
	output::DebugPrint(output::dbgLevel::Generic, L"Content size = %I64u\n", ullThisFolderSize);

	WC_MAPI_S(lpFolder->GetContentsTable(MAPI_ASSOCIATED, &lpTable));
	if (lpTable)
	{
		WC_MAPI_S(HrQueryAllRows(lpTable, reinterpret_cast<LPSPropTagArray>(&sProps), nullptr, nullptr, 0, &lpsRowSet));

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

	output::DebugPrint(output::dbgLevel::Generic, L"Total size = %I64u\n", ullThisFolderSize);

	return ullThisFolderSize;
}

ULONGLONG
ComputeFolderSize(
	_In_ const std::wstring& lpszProfile,
	_In_opt_ LPMAPIFOLDER lpFolder,
	_In_ const std::wstring& lpszFolder)
{
	output::DebugPrint(
		output::dbgLevel::Generic,
		L"ComputeFolderSize: Calculating size (including subfolders) for folder %ws from profile %ws \n",
		lpszFolder.c_str(),
		lpszProfile.c_str());
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
		static const SizedSPropTagArray(NUM_COLS, rgColProps) = {
			NUM_COLS,
			PR_DISPLAY_NAME_W,
			PR_ENTRYID,
			PR_FOLDER_TYPE,
		};

		// Size of this folder
		ullSize += ComputeSingleFolderSize(lpFolder);

		// Size of children
		WC_MAPI_S(lpFolder->GetHierarchyTable(MAPI_DEFERRED_ERRORS, &lpTable));
		if (lpTable)
		{
			WC_MAPI_S(lpTable->SetColumns(LPSPropTagArray(&rgColProps), TBL_ASYNC));

			for (;;)
			{
				if (lpRow) FreeProws(lpRow);
				lpRow = nullptr;
				WC_MAPI_S(lpTable->QueryRows(50, NULL, &lpRow));
				if (!lpRow || !lpRow->cRows) break;

				for (ULONG i = 0; i < lpRow->cRows; i++)
				{
					// Don't look at search folders
					if (PR_FOLDER_TYPE == lpRow->aRow[i].lpProps[ePR_FOLDER_TYPE].ulPropTag &&
						FOLDER_SEARCH == lpRow->aRow[i].lpProps[ePR_FOLDER_TYPE].Value.ul)
					{
						continue;
					}

					if (PR_ENTRYID == lpRow->aRow[i].lpProps[ePR_ENTRYID].ulPropTag)
					{
						ULONG ulObjType = NULL;
						LPMAPIFOLDER lpSubfolder = nullptr;

						WC_MAPI_S(lpFolder->OpenEntry(
							mapi::getBin(lpRow->aRow[i].lpProps[ePR_ENTRYID]).cb,
							reinterpret_cast<LPENTRYID>(mapi::getBin(lpRow->aRow[i].lpProps[ePR_ENTRYID]).lpb),
							nullptr,
							MAPI_BEST_ACCESS,
							&ulObjType,
							reinterpret_cast<LPUNKNOWN*>(&lpSubfolder)));

						if (lpSubfolder)
						{
							std::wstring szDisplayName;
							if (PR_DISPLAY_NAME_W == lpRow->aRow[i].lpProps[ePR_DISPLAY_NAME_W].ulPropTag)
							{
								szDisplayName = lpRow->aRow[i].lpProps[ePR_DISPLAY_NAME_W].Value.lpszW;
							}

							ullSize += ComputeFolderSize(lpszProfile, lpSubfolder, szDisplayName);

							lpSubfolder->Release();
						}
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
	_In_ const std::wstring& lpszProfile,
	_In_opt_ LPMAPIFOLDER lpFolder,
	_In_ const std::wstring& lpszFolder)
{
	output::DebugPrint(
		output::dbgLevel::Generic,
		L"DumpSearchState: Outputting search state for folder %ws from profile %ws \n",
		lpszFolder.c_str(),
		lpszProfile.c_str());

	if (lpFolder)
	{
		LPSRestriction lpRes = nullptr;
		LPENTRYLIST lpEntryList = nullptr;
		ULONG ulSearchState = 0;

		const auto hRes = WC_H(lpFolder->GetSearchCriteria(fMapiUnicode, &lpRes, &lpEntryList, &ulSearchState));
		if (MAPI_E_NOT_INITIALIZED == hRes)
		{
			wprintf(L"No search criteria has been set on this folder.\n");
		}
		else if (MAPI_E_NO_SUPPORT == hRes)
		{
			wprintf(L"This does not appear to be a search folder.\n");
		}
		else if (SUCCEEDED(hRes))
		{
			auto szFlags = flags::InterpretFlags(flagSearchState, ulSearchState);
			wprintf(L"Search state %ws == 0x%08X\n", szFlags.c_str(), ulSearchState);
			wprintf(L"\n");
			wprintf(L"Search Scope:\n");
			output::outputEntryList(output::dbgLevel::NoDebug, stdout, lpEntryList);
			wprintf(L"\n");
			wprintf(L"Search Criteria:\n");
			output::outputRestriction(output::dbgLevel::NoDebug, stdout, lpRes, nullptr);
		}

		MAPIFreeBuffer(lpRes);
		MAPIFreeBuffer(lpEntryList);
	}
}

void DoFolderProps(_In_opt_ LPMAPIFOLDER lpFolder)
{
	if (lpFolder)
	{
		PrintObjectProperties(L"folderprops", lpFolder, proptags::PropNameToPropTag(cli::switchUnswitched[0]));
	}
}

void DoFolderSize(_In_opt_ LPMAPIFOLDER lpFolder)
{
	const LONGLONG ullSize = ComputeFolderSize(cli::switchProfile[0], lpFolder, cli::switchFolder[0]);
	wprintf(L"Folder size (including subfolders)\n");
	wprintf(L"Bytes: %I64d\n", ullSize);
	wprintf(L"KB: %I64d\n", ullSize / 1024);
	wprintf(L"MB: %I64d\n", ullSize / (1024 * 1024));
}

void DoChildFolders(_In_opt_ LPMAPIFOLDER lpFolder)
{
	DumpHierarchyTable(cli::switchProfile[0], lpFolder, cli::switchFolder[0], 0);
}

void DoSearchState(_In_opt_ LPMAPIFOLDER lpFolder)
{
	DumpSearchState(cli::switchProfile[0], lpFolder, cli::switchFolder[0]);
}