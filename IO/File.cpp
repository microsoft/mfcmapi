#include "stdafx.h"
#include <IO/File.h>
#include <Interpret/InterpretProp.h>
#include <MAPI/MAPIFunctions.h>
#include <Interpret/String.h>
#include <UI/FileDialogEx.h>
#include <MAPI/MAPIProgress.h>
#include <Interpret/Guids.h>
#include "ImportProcs.h"
#include <UI/MFCUtilityFunctions.h>
#include <shlobj.h>
#include <MAPI/MAPIProcessor/DumpStore.h>
#include <UI/Dialogs/Editors/Editor.h>

wstring ShortenPath(const wstring& path)
{
	if (!path.empty())
	{
		auto hRes = S_OK;
		WCHAR szShortPath[MAX_PATH] = { 0 };
		size_t cchShortPath = NULL;
		// Use the short path to give us as much room as possible
		WC_D(cchShortPath, GetShortPathNameW(path.c_str(), szShortPath, _countof(szShortPath)));
		wstring ret = szShortPath;
		if (ret.back() != L'\\')
		{
			ret += L'\\';
		}

		return ret;
	}

	return path;
}

wstring GetDirectoryPath(HWND hWnd)
{
	WCHAR szPath[MAX_PATH] = { 0 };
	BROWSEINFOW BrowseInfo = { nullptr };
	auto hRes = S_OK;

	LPMALLOC lpMalloc = nullptr;

	EC_H(SHGetMalloc(&lpMalloc));

	if (!lpMalloc) return emptystring;

	auto szInputString = loadstring(IDS_DIRPROMPT);

	BrowseInfo.hwndOwner = hWnd;
	BrowseInfo.lpszTitle = szInputString.c_str();
	BrowseInfo.pszDisplayName = szPath;
	BrowseInfo.ulFlags = BIF_USENEWUI | BIF_RETURNONLYFSDIRS;

	// Note - I don't initialize COM for this call because MAPIInitialize does this
	auto lpItemIdList = SHBrowseForFolderW(&BrowseInfo);
	if (lpItemIdList)
	{
		EC_B(SHGetPathFromIDListW(lpItemIdList, szPath));
		lpMalloc->Free(lpItemIdList);
	}

	lpMalloc->Release();

	auto path = ShortenPath(szPath);
	if (path.length() >= MAXMSGPATH)
	{
		ErrDialog(__FILE__, __LINE__, IDS_EDPATHTOOLONG, path.length(), MAXMSGPATH);
		return emptystring;
	}

	return path;
}

// Opens storage with best access
_Check_return_ HRESULT MyStgOpenStorage(_In_ const wstring& szMessageFile, bool bBestAccess, _Deref_out_ LPSTORAGE* lppStorage)
{
	if (!lppStorage) return MAPI_E_INVALID_PARAMETER;
	DebugPrint(DBGGeneric, L"MyStgOpenStorage: Opening \"%ws\", bBestAccess == %ws\n", szMessageFile.c_str(), bBestAccess ? L"True" : L"False");
	auto hRes = S_OK;
	ULONG ulFlags = STGM_TRANSACTED;

	if (bBestAccess) ulFlags |= STGM_READWRITE;

	WC_H(::StgOpenStorage(
		szMessageFile.c_str(),
		nullptr,
		ulFlags,
		nullptr,
		0,
		lppStorage));

	// If we asked for best access (read/write) and didn't get it, then try it without readwrite
	if (STG_E_ACCESSDENIED == hRes && !*lppStorage && bBestAccess)
	{
		hRes = S_OK;
		EC_H(MyStgOpenStorage(szMessageFile, false, lppStorage));
	}

	return hRes;
}

// Creates an LPMESSAGE on top of the MSG file
_Check_return_ HRESULT LoadMSGToMessage(_In_ const wstring& szMessageFile, _Deref_out_opt_ LPMESSAGE* lppMessage)
{
	if (!lppMessage) return MAPI_E_INVALID_PARAMETER;

	auto hRes = S_OK;
	LPSTORAGE pStorage = nullptr;

	*lppMessage = nullptr;

	// get memory allocation function
	auto lpMalloc = MAPIGetDefaultMalloc();
	if (lpMalloc)
	{
		// Open the compound file
		EC_H(MyStgOpenStorage(szMessageFile, true, &pStorage));

		if (pStorage)
		{
			// Open an IMessage interface on an IStorage object
			EC_MAPI(OpenIMsgOnIStg(nullptr,
				MAPIAllocateBuffer,
				MAPIAllocateMore,
				MAPIFreeBuffer,
				lpMalloc,
				nullptr,
				pStorage,
				nullptr,
				0,
				0,
				lppMessage));
		}
	}

	if (pStorage) pStorage->Release();
	return hRes;
}

// Loads the MSG file into an LPMESSAGE pointer, then copies it into the passed in message
// lpMessage must be created first
_Check_return_ HRESULT LoadFromMSG(_In_ const wstring& szMessageFile, _In_ LPMESSAGE lpMessage, HWND hWnd)
{
	auto hRes = S_OK;
	LPMESSAGE pIMsg = nullptr;

	// Specify properties to exclude in the copy operation. These are
	// the properties that Exchange excludes to save bits and time.
	// Should not be necessary to exclude these, but speeds the process
	// when a lot of messages are being copied.
	static const SizedSPropTagArray(18, excludeTags) =
	{
	18,
	PR_REPLICA_VERSION,
	PR_DISPLAY_BCC,
	PR_DISPLAY_CC,
	PR_DISPLAY_TO,
	PR_ENTRYID,
	PR_MESSAGE_SIZE,
	PR_PARENT_ENTRYID,
	PR_RECORD_KEY,
	PR_STORE_ENTRYID,
	PR_STORE_RECORD_KEY,
	PR_MDB_PROVIDER,
	PR_ACCESS,
	PR_HASATTACH,
	PR_OBJECT_TYPE,
	PR_ACCESS_LEVEL,
	PR_HAS_NAMED_PROPERTIES,
	PR_REPLICA_SERVER,
	PR_HAS_DAMS
	};

	EC_H(LoadMSGToMessage(szMessageFile, &pIMsg));

	if (pIMsg)
	{
		LPSPropProblemArray lpProblems = nullptr;

		LPMAPIPROGRESS lpProgress = GetMAPIProgress(L"IMAPIProp::CopyTo", hWnd); // STRING_OK

		EC_MAPI(pIMsg->CopyTo(
			0,
			nullptr,
			LPSPropTagArray(&excludeTags),
			lpProgress ? reinterpret_cast<ULONG_PTR>(hWnd) : NULL,
			lpProgress,
			&IID_IMessage,
			lpMessage,
			lpProgress ? MAPI_DIALOG : 0,
			&lpProblems));

		if (lpProgress)
			lpProgress->Release();

		EC_PROBLEMARRAY(lpProblems);
		MAPIFreeBuffer(lpProblems);
		if (!FAILED(hRes))
		{
			EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
		}
	}

	if (pIMsg) pIMsg->Release();
	return hRes;
}

// lpMessage must be created first
_Check_return_ HRESULT LoadFromTNEF(_In_ const wstring& szMessageFile, _In_ LPADRBOOK lpAdrBook, _In_ LPMESSAGE lpMessage)
{
	auto hRes = S_OK;
	LPSTREAM lpStream = nullptr;
	LPITNEF lpTNEF = nullptr;
	LPSTnefProblemArray lpError = nullptr;
	LPSTREAM lpBodyStream = nullptr;

	if (szMessageFile.empty() || !lpAdrBook || !lpMessage) return MAPI_E_INVALID_PARAMETER;
	static auto dwKey = static_cast<WORD>(GetTickCount());

	enum
	{
		ulNumTNEFExcludeProps = 1
	};
	static const SizedSPropTagArray(ulNumTNEFExcludeProps, lpPropTnefExcludeArray) =
	{
	ulNumTNEFExcludeProps,
	PR_URL_COMP_NAME
	};

	// Get a Stream interface on the input TNEF file
	EC_H(MyOpenStreamOnFile(
		MAPIAllocateBuffer,
		MAPIFreeBuffer,
		STGM_READ,
		szMessageFile,
		&lpStream));

	// get the key value for OpenTnefStreamEx function
	dwKey++;

#pragma warning(push)
#pragma warning(disable:4616)
#pragma warning(disable:6276)
	EC_H(OpenTnefStreamEx(
		nullptr,
		lpStream,
		LPTSTR("winmail.dat"), // STRING_OK - despite its signature, this function is ANSI only
		TNEF_DECODE,
		lpMessage,
		dwKey,
		lpAdrBook,
		&lpTNEF));
#pragma warning(pop)

	if (lpTNEF)
	{
		// Decode the TNEF stream into our MAPI message.
		EC_MAPI(lpTNEF->ExtractProps(
			TNEF_PROP_EXCLUDE,
			LPSPropTagArray(&lpPropTnefExcludeArray),
			&lpError));

		EC_TNEFERR(lpError);
	}

	EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));

	if (lpBodyStream) lpBodyStream->Release();
	MAPIFreeBuffer(lpError);
	if (lpTNEF) lpTNEF->Release();
	if (lpStream) lpStream->Release();
	return hRes;
}

// Builds a file name out of the passed in message and extension
wstring BuildFileName(
	_In_ const wstring& ext,
	_In_ const wstring& dir,
	_In_ LPMESSAGE lpMessage)
{
	if (!lpMessage) return emptystring;

	enum
	{
		ePR_SUBJECT_W,
		ePR_RECORD_KEY,
		NUM_COLS
	};
	static const SizedSPropTagArray(NUM_COLS, sptaMessageProps) =
	{
	NUM_COLS,
	PR_SUBJECT_W,
	PR_RECORD_KEY
	};

	// Get subject line of message
	// This will be used as the new file name.
	auto hRes = S_OK;
	ULONG ulProps = NULL;
	LPSPropValue lpProps = nullptr;
	WC_H_GETPROPS(lpMessage->GetProps(
		LPSPropTagArray(&sptaMessageProps),
		fMapiUnicode,
		&ulProps,
		&lpProps));

	wstring subj;
	if (CheckStringProp(&lpProps[ePR_SUBJECT_W], PT_UNICODE))
	{
		subj = lpProps[ePR_SUBJECT_W].Value.lpszW;
	}

	LPSBinary lpRecordKey = nullptr;
	if (PR_RECORD_KEY == lpProps[ePR_RECORD_KEY].ulPropTag)
	{
		lpRecordKey = &lpProps[ePR_RECORD_KEY].Value.bin;
	}

	auto szFileOut = BuildFileNameAndPath(
		ext,
		subj,
		dir,
		lpRecordKey);

	MAPIFreeBuffer(lpProps);
	return szFileOut;
}

// The file name we generate should be shorter than MAX_PATH
// This includes our directory name too!
wstring BuildFileNameAndPath(
	_In_ const wstring& szExt,
	_In_ const wstring& szSubj,
	_In_ const wstring& szRootPath,
	_In_opt_ const LPSBinary lpBin)
{
	DebugPrint(DBGGeneric, L"BuildFileNameAndPath ext = \"%ws\"\n", szExt.c_str());
	DebugPrint(DBGGeneric, L"BuildFileNameAndPath subj = \"%ws\"\n", szSubj.c_str());
	DebugPrint(DBGGeneric, L"BuildFileNameAndPath rootPath = \"%ws\"\n", szRootPath.c_str());

	// Set up the path portion of the output.
	auto cleanRoot = ShortenPath(szRootPath);
	DebugPrint(DBGGeneric, L"BuildFileNameAndPath cleanRoot = \"%ws\"\n", cleanRoot.c_str());

	// If we don't have enough space for even the shortest filename, give up.
	if (cleanRoot.length() >= MAXMSGPATH) return emptystring;

	// Work with a max path which allows us to add our extension.
	// Shrink the max path to allow for a -ATTACHxxx extension.
	auto maxFile = MAX_PATH - cleanRoot.length() - szExt.length() - MAXATTACH;

	wstring cleanSubj;
	if (!szSubj.empty())
	{
		cleanSubj = SanitizeFileName(szSubj);
	}
	else
	{
		// We must have failed to get a subject before. Make one up.
		cleanSubj = L"UnknownSubject"; // STRING_OK
	}

	DebugPrint(DBGGeneric, L"BuildFileNameAndPath cleanSubj = \"%ws\"\n", cleanSubj.c_str());

	wstring szBin;
	if (lpBin && lpBin->cb)
	{
		szBin = L"_" + BinToHexString(lpBin, false);
	}

	if (cleanSubj.length() + szBin.length() <= maxFile)
	{
		auto szFile = cleanRoot + cleanSubj + szBin + szExt;
		DebugPrint(DBGGeneric, L"BuildFileNameAndPath fileOut= \"%ws\"\n", szFile.c_str());
		return szFile;
	}

	// We couldn't build the string we wanted, so try something shorter
	// Compute a shorter subject length that should fit.
	auto maxSubj = maxFile - min(MAXBIN, szBin.length()) - 1;
	auto szFile = cleanSubj.substr(0, maxSubj) + szBin.substr(0, MAXBIN);
	DebugPrint(DBGGeneric, L"BuildFileNameAndPath shorter file = \"%ws\"\n", szFile.c_str());
	DebugPrint(DBGGeneric, L"BuildFileNameAndPath new length = %d\n", szFile.length());

	if (szFile.length() >= maxFile)
	{
		szFile = cleanSubj.substr(0, MAXSUBJTIGHT) + szBin.substr(0, MAXBIN);
		DebugPrint(DBGGeneric, L"BuildFileNameAndPath shorter file = \"%ws\"\n", szFile.c_str());
		DebugPrint(DBGGeneric, L"BuildFileNameAndPath new length = %d\n", szFile.length());
	}

	if (szFile.length() >= maxFile)
	{
		DebugPrint(DBGGeneric, L"BuildFileNameAndPath failed to build a string\n");
		return emptystring;
	}

	auto szOut = cleanRoot + szFile + szExt;
	DebugPrint(DBGGeneric, L"BuildFileNameAndPath fileOut= \"%ws\"\n", szOut.c_str());
	return szOut;
}

void SaveFolderContentsToTXT(_In_ LPMDB lpMDB, _In_ LPMAPIFOLDER lpFolder, bool bRegular, bool bAssoc, bool bDescend, HWND hWnd)
{
	auto szDir = GetDirectoryPath(hWnd);

	if (!szDir.empty())
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.
		CDumpStore MyDumpStore;
		MyDumpStore.InitMDB(lpMDB);
		MyDumpStore.InitFolder(lpFolder);
		MyDumpStore.InitFolderPathRoot(szDir);
		MyDumpStore.ProcessFolders(
			bRegular,
			bAssoc,
			bDescend);
	}
}

_Check_return_ HRESULT SaveFolderContentsToMSG(_In_ LPMAPIFOLDER lpFolder, _In_ const wstring& szPathName, bool bAssoc, bool bUnicode, HWND hWnd)
{
	auto hRes = S_OK;
	LPMAPITABLE lpFolderContents = nullptr;
	LPSRowSet pRows = nullptr;

	enum
	{
		fldPR_ENTRYID,
		fldPR_SUBJECT_W,
		fldPR_RECORD_KEY,
		fldNUM_COLS
	};

	static const SizedSPropTagArray(fldNUM_COLS, fldCols) =
	{
	fldNUM_COLS,
	PR_ENTRYID,
	PR_SUBJECT_W,
	PR_RECORD_KEY
	};

	if (!lpFolder || szPathName.empty()) return MAPI_E_INVALID_PARAMETER;
	if (szPathName.length() >= MAXMSGPATH) return MAPI_E_INVALID_PARAMETER;

	DebugPrint(DBGGeneric, L"SaveFolderContentsToMSG: Saving contents of folder to \"%ws\"\n", szPathName.c_str());

	EC_MAPI(lpFolder->GetContentsTable(
		fMapiUnicode | (bAssoc ? MAPI_ASSOCIATED : NULL),
		&lpFolderContents));

	if (lpFolderContents)
	{
		EC_MAPI(lpFolderContents->SetColumns(LPSPropTagArray(&fldCols), TBL_BATCH));

		if (!FAILED(hRes)) for (;;)
		{
			hRes = S_OK;
			if (pRows) FreeProws(pRows);
			pRows = nullptr;
			EC_MAPI(lpFolderContents->QueryRows(
				1,
				NULL,
				&pRows));
			if (FAILED(hRes) || !pRows || pRows && !pRows->cRows) break;

			WC_H(SaveToMSG(
				lpFolder,
				szPathName,
				pRows->aRow->lpProps[fldPR_ENTRYID],
				&pRows->aRow->lpProps[fldPR_RECORD_KEY],
				&pRows->aRow->lpProps[fldPR_SUBJECT_W],
				bUnicode,
				hWnd));
		}
	}

	if (pRows) FreeProws(pRows);
	if (lpFolderContents) lpFolderContents->Release();
	return hRes;
}

#ifndef MRMAPI
void ExportMessages(_In_ const LPMAPIFOLDER lpFolder, HWND hWnd)
{
	auto hRes = S_OK;
	CEditor MyData(
		nullptr,
		IDS_EXPORTTITLE,
		IDS_EXPORTPROMPT,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	MyData.InitPane(0, TextPane::CreateSingleLinePane(IDS_EXPORTSEARCHTERM, false));

	WC_H(MyData.DisplayDialog());
	if (S_OK != hRes) return;

	auto szDir = GetDirectoryPath(hWnd);
	if (szDir.empty()) return;
	auto restrictString = MyData.GetStringW(0);
	if (restrictString.empty()) return;

	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	LPSRestriction lpRes = nullptr;
	// Allocate and create our SRestriction
	EC_H(CreatePropertyStringRestriction(
		PR_SUBJECT_W,
		restrictString,
		FL_SUBSTRING | FL_IGNORECASE,
		nullptr,
		&lpRes));

	LPMAPITABLE lpTable = nullptr;
	WC_MAPI(lpFolder->GetContentsTable(MAPI_DEFERRED_ERRORS | MAPI_UNICODE, &lpTable));
	if (lpTable)
	{
		WC_MAPI(lpTable->Restrict(lpRes, 0));
		if (SUCCEEDED(hRes))
		{
			enum
			{
				fldPR_ENTRYID,
				fldPR_SUBJECT_W,
				fldPR_RECORD_KEY,
				fldNUM_COLS
			};

			static const SizedSPropTagArray(fldNUM_COLS, fldCols) =
			{
				fldNUM_COLS,
				PR_ENTRYID,
				PR_SUBJECT_W,
				PR_RECORD_KEY
			};

			WC_MAPI(lpTable->SetColumns(LPSPropTagArray(&fldCols), TBL_ASYNC));

			// Export messages in the rows
			LPSRowSet lpRows = nullptr;
			if (!FAILED(hRes)) for (;;)
			{
				hRes = S_OK;
				if (lpRows) FreeProws(lpRows);
				lpRows = nullptr;
				WC_MAPI(lpTable->QueryRows(
					50,
					NULL,
					&lpRows));
				if (FAILED(hRes) || !lpRows || !lpRows->cRows) break;

				for (ULONG i = 0; i < lpRows->cRows; i++)
				{
					hRes = S_OK;
					WC_H(SaveToMSG(
						lpFolder,
						szDir,
						lpRows->aRow[i].lpProps[fldPR_ENTRYID],
						&lpRows->aRow[i].lpProps[fldPR_RECORD_KEY],
						&lpRows->aRow[i].lpProps[fldPR_SUBJECT_W],
						true,
						hWnd));
				}
			}
		}

		lpTable->Release();
	}

	if (lpRes) MAPIFreeBuffer(lpRes);
}
#endif

_Check_return_ HRESULT WriteStreamToFile(_In_ LPSTREAM pStrmSrc, _In_ const wstring& szFileName)
{
	if (!pStrmSrc || szFileName.empty()) return MAPI_E_INVALID_PARAMETER;

	auto hRes = S_OK;
	LPSTREAM pStrmDest = nullptr;
	STATSTG StatInfo = { nullptr };

	// Open an IStream interface and create the file at the
	// same time. This code will create the file in the
	// current directory.
	EC_H(MyOpenStreamOnFile(
		MAPIAllocateBuffer,
		MAPIFreeBuffer,
		STGM_CREATE | STGM_READWRITE,
		szFileName,
		&pStrmDest));

	if (pStrmDest)
	{
		pStrmSrc->Stat(&StatInfo, STATFLAG_NONAME);

		DebugPrint(DBGStream, L"WriteStreamToFile: Writing cb = %llu bytes\n", StatInfo.cbSize.QuadPart);

		EC_MAPI(pStrmSrc->CopyTo(pStrmDest,
			StatInfo.cbSize,
			nullptr,
			nullptr));

		// Commit changes to new stream
		EC_MAPI(pStrmDest->Commit(STGC_DEFAULT));

		pStrmDest->Release();
	}

	return hRes;
}

_Check_return_ HRESULT SaveToEML(_In_ LPMESSAGE lpMessage, _In_ const wstring& szFileName)
{
	auto hRes = S_OK;
	LPSTREAM pStrmSrc = nullptr;

	if (!lpMessage || szFileName.empty()) return MAPI_E_INVALID_PARAMETER;
	DebugPrint(DBGGeneric, L"SaveToEML: Saving message to \"%ws\"\n", szFileName.c_str());

	// Open the property of the attachment
	// containing the file data
	EC_MAPI(lpMessage->OpenProperty(
		PR_INTERNET_CONTENT,
		const_cast<LPIID>(&IID_IStream),
		0,
		NULL, // MAPI_MODIFY is not needed
		reinterpret_cast<LPUNKNOWN *>(&pStrmSrc)));
	if (FAILED(hRes))
	{
		if (MAPI_E_NOT_FOUND == hRes)
		{
			DebugPrint(DBGGeneric, L"No internet content found\n");
		}
	}
	else
	{
		if (pStrmSrc)
		{
			WC_H(WriteStreamToFile(pStrmSrc, szFileName));

			pStrmSrc->Release();
		}
	}

	return hRes;
}

_Check_return_ HRESULT STDAPICALLTYPE MyStgCreateStorageEx(_In_ const wstring& pName,
	DWORD grfMode,
	DWORD stgfmt,
	DWORD grfAttrs,
	_In_ STGOPTIONS * pStgOptions,
	_Pre_null_ void * reserved,
	_In_ REFIID riid,
	_Out_ void ** ppObjectOpen)
{
	if (pName.empty()) return MAPI_E_INVALID_PARAMETER;

	if (pfnStgCreateStorageEx)
	{
		return pfnStgCreateStorageEx(
			pName.c_str(),
			grfMode,
			stgfmt,
			grfAttrs,
			pStgOptions,
			reserved,
			riid,
			ppObjectOpen);
	}

	// Fallback for NT4, which doesn't have StgCreateStorageEx
	return StgCreateDocfile(
		pName.c_str(),
		grfMode,
		0,
		reinterpret_cast<LPSTORAGE*>(ppObjectOpen));
}

_Check_return_ HRESULT CreateNewMSG(_In_ const wstring& szFileName, bool bUnicode, _Deref_out_opt_ LPMESSAGE* lppMessage, _Deref_out_opt_ LPSTORAGE* lppStorage)
{
	if (szFileName.empty() || !lppMessage || !lppStorage) return MAPI_E_INVALID_PARAMETER;

	auto hRes = S_OK;
	LPSTORAGE pStorage = nullptr;
	LPMESSAGE pIMsg = nullptr;

	*lppMessage = nullptr;
	*lppStorage = nullptr;

	// get memory allocation function
	auto pMalloc = MAPIGetDefaultMalloc();
	if (pMalloc)
	{
		STGOPTIONS myOpts = { 0 };

		myOpts.usVersion = 1; // STGOPTIONS_VERSION
		myOpts.ulSectorSize = 4096;

		// Open the compound file
		EC_H(MyStgCreateStorageEx(
			szFileName,
			STGM_READWRITE | STGM_TRANSACTED | STGM_CREATE,
			STGFMT_DOCFILE,
			0, // FILE_FLAG_NO_BUFFERING,
			&myOpts,
			nullptr,
			__uuidof(IStorage),
			reinterpret_cast<LPVOID*>(&pStorage)));
		if (SUCCEEDED(hRes) && pStorage)
		{
			// Open an IMessage interface on an IStorage object
			EC_MAPI(OpenIMsgOnIStg(
				nullptr,
				MAPIAllocateBuffer,
				MAPIAllocateMore,
				MAPIFreeBuffer,
				pMalloc,
				nullptr,
				pStorage,
				nullptr,
				0,
				bUnicode ? MAPI_UNICODE : 0,
				&pIMsg));
			if (SUCCEEDED(hRes) && pIMsg)
			{
				// write the CLSID to the IStorage instance - pStorage. This will
				// only work with clients that support this compound document type
				// as the storage medium. If the client does not support
				// CLSID_MailMessage as the compound document, you will have to use
				// the CLSID that it does support.
				EC_MAPI(WriteClassStg(pStorage, CLSID_MailMessage));
				if (SUCCEEDED(hRes))
				{
					*lppStorage = pStorage;
					(*lppStorage)->AddRef();
					*lppMessage = pIMsg;
					(*lppMessage)->AddRef();
				}
			}

			if (pIMsg) pIMsg->Release();
		}

		if (pStorage) pStorage->Release();
	}

	return hRes;
}

_Check_return_ HRESULT SaveToMSG(
	_In_ const LPMAPIFOLDER lpFolder,
	_In_ const wstring& szPathName,
	_In_ const SPropValue& entryID,
	_In_ const LPSPropValue lpRecordKey,
	_In_ const LPSPropValue lpSubject,
	bool bUnicode,
	HWND hWnd)
{
	if (szPathName.empty() || szPathName.length() >= MAXMSGPATH) return MAPI_E_INVALID_PARAMETER;
	if (entryID.ulPropTag != PR_ENTRYID) return MAPI_E_INVALID_PARAMETER;

	auto hRes = S_OK;
	LPMESSAGE lpMessage = nullptr;

	DebugPrint(DBGGeneric, L"SaveToMSG: Saving message to \"%ws\"\n", szPathName.c_str());

	DebugPrint(DBGGeneric, L"Source Message =\n");
	DebugPrintBinary(DBGGeneric, entryID.Value.bin);

	EC_H(CallOpenEntry(
		nullptr,
		nullptr,
		reinterpret_cast<LPMAPICONTAINER>(lpFolder),
		nullptr,
		&entryID.Value.bin,
		nullptr,
		MAPI_BEST_ACCESS,
		nullptr,
		reinterpret_cast<LPUNKNOWN*>(&lpMessage)));
	if (FAILED(hRes) || !lpMessage) return hRes;

	auto szSubj = CheckStringProp(lpSubject, PT_UNICODE) ? lpSubject->Value.lpszW : L"UnknownSubject";
	LPSBinary recordKey = (lpRecordKey && lpRecordKey->ulPropTag == PR_RECORD_KEY) ? &lpRecordKey->Value.bin : nullptr;

	auto szFileName = BuildFileNameAndPath(L".msg", szSubj, szPathName, recordKey); // STRING_OK
	if (!szFileName.empty())
	{
		DebugPrint(DBGGeneric, L"Saving to = \"%ws\"\n", szFileName.c_str());

		EC_H(SaveToMSG(
			lpMessage,
			szFileName,
			bUnicode,
			hWnd,
			false));

		DebugPrint(DBGGeneric, L"Message Saved\n");
	}

	if (lpMessage) lpMessage->Release();

	return hRes;
}

_Check_return_ HRESULT SaveToMSG(_In_ LPMESSAGE lpMessage, _In_ const wstring& szFileName, bool bUnicode, HWND hWnd, bool bAllowUI)
{
	auto hRes = S_OK;
	LPSTORAGE pStorage = nullptr;
	LPMESSAGE pIMsg = nullptr;

	if (!lpMessage || szFileName.empty()) return MAPI_E_INVALID_PARAMETER;

	DebugPrint(DBGGeneric, L"SaveToMSG: Saving message to \"%ws\"\n", szFileName.c_str());

	EC_H(CreateNewMSG(szFileName, bUnicode, &pIMsg, &pStorage));
	if (pIMsg && pStorage)
	{
		// Specify properties to exclude in the copy operation. These are
		// the properties that Exchange excludes to save bits and time.
		// Should not be necessary to exclude these, but speeds the process
		// when a lot of messages are being copied.
		static const SizedSPropTagArray(7, excludeTags) =
		{
		7,
		PR_ACCESS,
		PR_BODY,
		PR_RTF_SYNC_BODY_COUNT,
		PR_RTF_SYNC_BODY_CRC,
		PR_RTF_SYNC_BODY_TAG,
		PR_RTF_SYNC_PREFIX_COUNT,
		PR_RTF_SYNC_TRAILING_COUNT
		};

		EC_H(CopyTo(
			hWnd,
			lpMessage,
			pIMsg,
			&IID_IMessage,
			LPSPropTagArray(&excludeTags),
			false,
			bAllowUI));

		// save changes to IMessage object.
		EC_MAPI(pIMsg->SaveChanges(KEEP_OPEN_READWRITE));

		// save changes in storage of new doc file
		EC_MAPI(pStorage->Commit(STGC_DEFAULT));
	}

	if (pStorage) pStorage->Release();
	if (pIMsg) pIMsg->Release();

	return hRes;
}

_Check_return_ HRESULT SaveToTNEF(_In_ LPMESSAGE lpMessage, _In_ LPADRBOOK lpAdrBook, _In_ const wstring& szFileName)
{
	auto hRes = S_OK;

	enum
	{
		ulNumTNEFIncludeProps = 2
	};
	static const SizedSPropTagArray(ulNumTNEFIncludeProps, lpPropTnefIncludeArray) =
	{
	ulNumTNEFIncludeProps,
	PR_MESSAGE_RECIPIENTS,
	PR_ATTACH_DATA_BIN
	};

	enum
	{
		ulNumTNEFExcludeProps = 1
	};
	static const SizedSPropTagArray(ulNumTNEFExcludeProps, lpPropTnefExcludeArray) =
	{
	ulNumTNEFExcludeProps,
	PR_URL_COMP_NAME
	};

	if (!lpMessage || !lpAdrBook || szFileName.empty()) return MAPI_E_INVALID_PARAMETER;
	DebugPrint(DBGGeneric, L"SaveToTNEF: Saving message to \"%ws\"\n", szFileName.c_str());

	LPSTREAM lpStream = nullptr;
	LPITNEF lpTNEF = nullptr;
	LPSTnefProblemArray lpError = nullptr;

	static auto dwKey = static_cast<WORD>(GetTickCount());

	// Get a Stream interface on the input TNEF file
	EC_H(MyOpenStreamOnFile(
		MAPIAllocateBuffer,
		MAPIFreeBuffer,
		STGM_READWRITE | STGM_CREATE,
		szFileName,
		&lpStream));

	if (lpStream)
	{
		// Open TNEF stream
#pragma warning(push)
#pragma warning(disable:4616)
#pragma warning(disable:6276)
		EC_H(OpenTnefStreamEx(
			nullptr,
			lpStream,
			LPTSTR("winmail.dat"), // STRING_OK - despite its signature, this function is ANSI only
			TNEF_ENCODE,
			lpMessage,
			dwKey,
			lpAdrBook,
			&lpTNEF));
#pragma warning(pop)

		if (lpTNEF)
		{
			// Excludes
			EC_MAPI(lpTNEF->AddProps(
				TNEF_PROP_EXCLUDE,
				0,
				nullptr,
				LPSPropTagArray(&lpPropTnefExcludeArray)
			));
			EC_MAPI(lpTNEF->AddProps(
				TNEF_PROP_EXCLUDE | TNEF_PROP_ATTACHMENTS_ONLY,
				0,
				nullptr,
				LPSPropTagArray(&lpPropTnefExcludeArray)
			));

			EC_MAPI(lpTNEF->AddProps(
				TNEF_PROP_INCLUDE,
				0,
				nullptr,
				LPSPropTagArray(&lpPropTnefIncludeArray)
			));

			EC_MAPI(lpTNEF->Finish(0, &dwKey, &lpError));

			EC_TNEFERR(lpError);

			// Saving stream
			EC_MAPI(lpStream->Commit(STGC_DEFAULT));

			MAPIFreeBuffer(lpError);
			lpTNEF->Release();
		}
		lpStream->Release();
	}

	return hRes;
}

_Check_return_ HRESULT DeleteAttachments(_In_ LPMESSAGE lpMessage, _In_ const wstring& szAttName, HWND hWnd)
{
	LPSPropValue pProps = nullptr;
	auto hRes = S_OK;
	LPMAPITABLE lpAttTbl = nullptr;
	LPSRowSet pRows = nullptr;
	ULONG iRow;

	enum
	{
		ATTACHNUM,
		ATTACHNAME,
		NUM_COLS
	};
	static const SizedSPropTagArray(NUM_COLS, sptAttachTableCols) =
	{
	NUM_COLS,
	PR_ATTACH_NUM,
	PR_ATTACH_LONG_FILENAME_W
	};

	if (!lpMessage) return MAPI_E_INVALID_PARAMETER;

	EC_MAPI(HrGetOneProp(lpMessage, PR_HASATTACH, &pProps));

	if (pProps && PR_HASATTACH == pProps[0].ulPropTag && pProps[0].Value.b)
	{
		EC_MAPI(lpMessage->OpenProperty(
			PR_MESSAGE_ATTACHMENTS,
			&IID_IMAPITable,
			fMapiUnicode,
			0,
			reinterpret_cast<LPUNKNOWN *>(&lpAttTbl)));

		if (lpAttTbl)
		{
			// I would love to just pass a restriction
			// to HrQueryAllRows for the file name. However,
			// we don't support restricting attachment tables (EMSMDB32 that is)
			// So I have to compare the strings myself (see below)
			EC_MAPI(HrQueryAllRows(lpAttTbl,
				LPSPropTagArray(&sptAttachTableCols),
				nullptr,
				nullptr,
				0,
				&pRows));

			if (pRows)
			{
				auto bDirty = false;

				if (!FAILED(hRes)) for (iRow = 0; iRow < pRows->cRows; iRow++)
				{
					hRes = S_OK;

					if (PR_ATTACH_NUM != pRows->aRow[iRow].lpProps[ATTACHNUM].ulPropTag) continue;

					if (!szAttName.empty())
					{
						if (PR_ATTACH_LONG_FILENAME_W != pRows->aRow[iRow].lpProps[ATTACHNAME].ulPropTag ||
							szAttName != pRows->aRow[iRow].lpProps[ATTACHNAME].Value.lpszW)
							continue;
					}

					// Open the attachment
					LPMAPIPROGRESS lpProgress = GetMAPIProgress(L"IMessage::DeleteAttach", hWnd); // STRING_OK

					EC_MAPI(lpMessage->DeleteAttach(
						pRows->aRow[iRow].lpProps[ATTACHNUM].Value.l,
						lpProgress ? reinterpret_cast<ULONG_PTR>(hWnd) : NULL,
						lpProgress,
						lpProgress ? ATTACH_DIALOG : 0));

					if (SUCCEEDED(hRes))
						bDirty = true;

					if (lpProgress)
						lpProgress->Release();
				}

				// Moved this inside the if (pRows) check
				// and also added a flag so we only call this if we
				// got a successful DeleteAttach call
				if (bDirty)
					EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
			}
		}
	}

	if (pRows) FreeProws(pRows);
	if (lpAttTbl) lpAttTbl->Release();
	MAPIFreeBuffer(pProps);

	return hRes;
}

#ifndef MRMAPI
_Check_return_ HRESULT WriteAttachmentsToFile(_In_ LPMESSAGE lpMessage, HWND hWnd)
{
	LPSPropValue pProps = nullptr;
	auto hRes = S_OK;
	LPMAPITABLE lpAttTbl = nullptr;
	LPSRowSet pRows = nullptr;
	ULONG iRow;
	LPATTACH lpAttach = nullptr;

	enum
	{
		ATTACHNUM,
		NUM_COLS
	};
	static const SizedSPropTagArray(NUM_COLS, sptAttachTableCols) =
	{
	NUM_COLS,
	PR_ATTACH_NUM
	};

	if (!lpMessage) return MAPI_E_INVALID_PARAMETER;

	EC_MAPI(HrGetOneProp(lpMessage, PR_HASATTACH, &pProps));

	if (pProps && PR_HASATTACH == pProps[0].ulPropTag && pProps[0].Value.b)
	{
		EC_MAPI(lpMessage->OpenProperty(
			PR_MESSAGE_ATTACHMENTS,
			&IID_IMAPITable,
			fMapiUnicode,
			0,
			reinterpret_cast<LPUNKNOWN *>(&lpAttTbl)));

		if (lpAttTbl)
		{
			EC_MAPI(HrQueryAllRows(lpAttTbl,
				LPSPropTagArray(&sptAttachTableCols),
				nullptr,
				nullptr,
				0,
				&pRows));

			if (pRows)
			{
				if (!FAILED(hRes)) for (iRow = 0; iRow < pRows->cRows; iRow++)
				{
					lpAttach = nullptr;

					if (PR_ATTACH_NUM != pRows->aRow[iRow].lpProps[ATTACHNUM].ulPropTag) continue;

					// Open the attachment
					EC_MAPI(lpMessage->OpenAttach(
						pRows->aRow[iRow].lpProps[ATTACHNUM].Value.l,
						nullptr,
						MAPI_BEST_ACCESS,
						&lpAttach));

					if (lpAttach)
					{
						WC_H(WriteAttachmentToFile(lpAttach, hWnd));
						lpAttach->Release();
						lpAttach = nullptr;
						if (S_OK != hRes && iRow != pRows->cRows - 1)
						{
							if (bShouldCancel(nullptr, hRes)) break;
							hRes = S_OK;
						}
					}
				}
				if (pRows) FreeProws(pRows);
			}
			lpAttTbl->Release();
		}
	}

	MAPIFreeBuffer(pProps);

	return hRes;
}
#endif

_Check_return_ HRESULT WriteEmbeddedMSGToFile(_In_ LPATTACH lpAttach, _In_ const wstring& szFileName, bool bUnicode, HWND hWnd)
{
	auto hRes = S_OK;
	LPMESSAGE lpAttachMsg = nullptr;

	if (!lpAttach || szFileName.empty()) return MAPI_E_INVALID_PARAMETER;

	DebugPrint(DBGGeneric, L"WriteEmbeddedMSGToFile: Saving attachment to \"%ws\"\n", szFileName.c_str());

	EC_MAPI(lpAttach->OpenProperty(
		PR_ATTACH_DATA_OBJ,
		const_cast<LPIID>(&IID_IMessage),
		0,
		NULL, // MAPI_MODIFY is not needed
		reinterpret_cast<LPUNKNOWN *>(&lpAttachMsg)));

	if (lpAttachMsg)
	{
		EC_H(SaveToMSG(lpAttachMsg, szFileName, bUnicode, hWnd, false));
		lpAttachMsg->Release();
	}

	return hRes;
}

_Check_return_ HRESULT WriteAttachStreamToFile(_In_ LPATTACH lpAttach, _In_ const wstring& szFileName)
{
	auto hRes = S_OK;
	LPSTREAM pStrmSrc = nullptr;

	if (!lpAttach || szFileName.empty()) return MAPI_E_INVALID_PARAMETER;

	// Open the property of the attachment
	// containing the file data
	WC_MAPI(lpAttach->OpenProperty(
		PR_ATTACH_DATA_BIN,
		const_cast<LPIID>(&IID_IStream),
		0,
		NULL, // MAPI_MODIFY is not needed
		reinterpret_cast<LPUNKNOWN *>(&pStrmSrc)));
	if (FAILED(hRes))
	{
		if (MAPI_E_NOT_FOUND == hRes)
		{
			DebugPrint(DBGGeneric, L"No attachments found. Maybe the attachment was a message?\n");
		}
		else CHECKHRES(hRes);
	}
	else
	{
		if (pStrmSrc)
		{
			WC_H(WriteStreamToFile(pStrmSrc, szFileName));

			pStrmSrc->Release();
		}
	}

	return hRes;
}

// Pretty sure this covers all OLE attachments - we don't need to look at PR_ATTACH_TAG
_Check_return_ HRESULT WriteOleToFile(_In_ LPATTACH lpAttach, _In_ const wstring& szFileName)
{
	auto hRes = S_OK;
	LPSTORAGE lpStorageSrc = nullptr;
	LPSTORAGE lpStorageDest = nullptr;
	LPSTREAM pStrmSrc = nullptr;

	// Open the property of the attachment containing the OLE data
	// Try to get it as an IStreamDocFile file first as that will be faster
	WC_MAPI(lpAttach->OpenProperty(
		PR_ATTACH_DATA_OBJ,
		const_cast<LPIID>(&IID_IStreamDocfile),
		0,
		NULL,
		reinterpret_cast<LPUNKNOWN *>(&pStrmSrc)));

	// We got IStreamDocFile! Great! We can copy stream to stream into the file
	if (pStrmSrc)
	{
		WC_H(WriteStreamToFile(pStrmSrc, szFileName));

		pStrmSrc->Release();
	}
	// We couldn't get IStreamDocFile! No problem - we'll try IStorage next
	else
	{
		hRes = S_OK;
		EC_MAPI(lpAttach->OpenProperty(
			PR_ATTACH_DATA_OBJ,
			const_cast<LPIID>(&IID_IStorage),
			0,
			NULL,
			reinterpret_cast<LPUNKNOWN *>(&lpStorageSrc)));

		if (lpStorageSrc)
		{
			EC_H(::StgCreateDocfile(
				szFileName.c_str(),
				STGM_READWRITE | STGM_TRANSACTED | STGM_CREATE,
				0,
				&lpStorageDest));
			if (lpStorageDest)
			{
				EC_MAPI(lpStorageSrc->CopyTo(
					NULL,
					nullptr,
					nullptr,
					lpStorageDest));

				EC_MAPI(lpStorageDest->Commit(STGC_DEFAULT));
				lpStorageDest->Release();
			}

			lpStorageSrc->Release();
		}
	}

	return hRes;
}

#ifndef MRMAPI
_Check_return_ HRESULT WriteAttachmentToFile(_In_ LPATTACH lpAttach, HWND hWnd)
{
	auto hRes = S_OK;
	LPSPropValue lpProps = nullptr;
	ULONG ulProps = 0;
	INT_PTR iDlgRet = 0;

	enum
	{
		ATTACH_METHOD,
		ATTACH_LONG_FILENAME_W,
		ATTACH_FILENAME_W,
		DISPLAY_NAME_W,
		NUM_COLS
	};
	static const SizedSPropTagArray(NUM_COLS, sptaAttachProps) =
	{
	NUM_COLS,
	PR_ATTACH_METHOD,
	PR_ATTACH_LONG_FILENAME_W,
	PR_ATTACH_FILENAME_W,
	PR_DISPLAY_NAME_W
	};

	if (!lpAttach) return MAPI_E_INVALID_PARAMETER;

	DebugPrint(DBGGeneric, L"WriteAttachmentToFile: Saving attachment.\n");

	// Get required properties from the message
	EC_H_GETPROPS(lpAttach->GetProps(
		LPSPropTagArray(&sptaAttachProps), // property tag array
		fMapiUnicode, // flags
		&ulProps, // Count of values returned
		&lpProps)); // Values returned

	if (lpProps)
	{
		auto szName = L"Unknown"; // STRING_OK

		// Get a file name to use
		if (CheckStringProp(&lpProps[ATTACH_LONG_FILENAME_W], PT_UNICODE))
		{
			szName = lpProps[ATTACH_LONG_FILENAME_W].Value.lpszW;
		}
		else if (CheckStringProp(&lpProps[ATTACH_FILENAME_W], PT_UNICODE))
		{
			szName = lpProps[ATTACH_FILENAME_W].Value.lpszW;
		}
		else if (CheckStringProp(&lpProps[DISPLAY_NAME_W], PT_UNICODE))
		{
			szName = lpProps[DISPLAY_NAME_W].Value.lpszW;
		}

		auto szFileName = SanitizeFileName(szName);

		// Get File Name
		switch (lpProps[ATTACH_METHOD].Value.l)
		{
		case ATTACH_BY_VALUE:
		case ATTACH_BY_REFERENCE:
		case ATTACH_BY_REF_RESOLVE:
		case ATTACH_BY_REF_ONLY:
		{
			DebugPrint(DBGGeneric, L"WriteAttachmentToFile: Prompting with \"%ws\"\n", szFileName.c_str());

			auto file = CFileDialogExW::SaveAs(
				L"txt", // STRING_OK
				szFileName,
				OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
				loadstring(IDS_ALLFILES));
			if (!file.empty())
			{
				EC_H(WriteAttachStreamToFile(lpAttach, file));
			}
		}
		break;
		case ATTACH_EMBEDDED_MSG:
			// Get File Name
		{
			DebugPrint(DBGGeneric, L"WriteAttachmentToFile: Prompting with \"%ws\"\n", szFileName.c_str());
			auto file = CFileDialogExW::SaveAs(
				L"msg", // STRING_OK
				szFileName,
				OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
				loadstring(IDS_MSGFILES));
			if (!file.empty())
			{
				EC_H(WriteEmbeddedMSGToFile(lpAttach, file, (MAPI_UNICODE == fMapiUnicode) ? true : false, hWnd));
			}
		}
		break;
		case ATTACH_OLE:
		{
			DebugPrint(DBGGeneric, L"WriteAttachmentToFile: Prompting with \"%ws\"\n", szFileName.c_str());
			auto file = CFileDialogExW::SaveAs(
				emptystring,
				szFileName,
				OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
				loadstring(IDS_ALLFILES));
			if (!file.empty())
			{
				EC_H(WriteOleToFile(lpAttach, file));
			}
		}
		break;
		default:
			ErrDialog(__FILE__, __LINE__, IDS_EDUNKNOWNATTMETHOD); break;
		}
	}
	if (iDlgRet == IDCANCEL) hRes = MAPI_E_USER_CANCEL;

	MAPIFreeBuffer(lpProps);
	return hRes;
}
#endif
