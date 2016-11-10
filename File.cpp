#include "stdafx.h"
#include "File.h"
#include "InterpretProp.h"
#include "MAPIFunctions.h"
#include "String.h"
#include "FileDialogEx.h"
#include "MAPIProgress.h"
#include "guids.h"
#include "ImportProcs.h"
#include "MFCUtilityFunctions.h"
#include <shlobj.h>
#include "Dumpstore.h"

wstring GetDirectoryPath(HWND hWnd)
{
	WCHAR szPath[MAX_PATH] = { 0 };
	BROWSEINFOW BrowseInfo = { 0 };
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
	return szPath;
}

// Opens storage with best access
_Check_return_ HRESULT MyStgOpenStorage(_In_z_ LPCWSTR szMessageFile, bool bBestAccess, _Deref_out_ LPSTORAGE* lppStorage)
{
	if (!lppStorage) return MAPI_E_INVALID_PARAMETER;
	DebugPrint(DBGGeneric, L"MyStgOpenStorage: Opening \"%ws\", bBestAccess == %ws\n", szMessageFile, bBestAccess ? L"True" : L"False");
	auto hRes = S_OK;
	ULONG ulFlags = STGM_TRANSACTED;

	if (bBestAccess) ulFlags |= STGM_READWRITE;

	WC_H(::StgOpenStorage(
		szMessageFile,
		NULL,
		ulFlags,
		NULL,
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
_Check_return_ HRESULT LoadMSGToMessage(_In_z_ LPCWSTR szMessageFile, _Deref_out_opt_ LPMESSAGE* lppMessage)
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
			EC_MAPI(OpenIMsgOnIStg(NULL,
				MAPIAllocateBuffer,
				MAPIAllocateMore,
				MAPIFreeBuffer,
				lpMalloc,
				NULL,
				pStorage,
				NULL,
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
_Check_return_ HRESULT LoadFromMSG(_In_z_ LPCWSTR szMessageFile, _In_ LPMESSAGE lpMessage, HWND hWnd)
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
			NULL,
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
_Check_return_ HRESULT LoadFromTNEF(_In_z_ LPCWSTR szMessageFile, _In_ LPADRBOOK lpAdrBook, _In_ LPMESSAGE lpMessage)
{
	auto hRes = S_OK;
	LPSTREAM lpStream = nullptr;
	LPITNEF lpTNEF = nullptr;
	LPSTnefProblemArray lpError = nullptr;
	LPSTREAM lpBodyStream = nullptr;

	if (!szMessageFile || !lpAdrBook || !lpMessage) return MAPI_E_INVALID_PARAMETER;
	static auto dwKey = static_cast<WORD>(::GetTickCount());

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
		NULL,
		&lpStream));

	// get the key value for OpenTnefStreamEx function
	dwKey++;

#pragma warning(push)
#pragma warning(disable:4616)
#pragma warning(disable:6276)
	EC_H(OpenTnefStreamEx(
		NULL,
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
	_In_ const wstring& szExt,
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

	wstring szSubj;
	if (CheckStringProp(&lpProps[ePR_SUBJECT_W], PT_UNICODE))
	{
		szSubj = lpProps[ePR_SUBJECT_W].Value.lpszW;
	}

	LPSBinary lpRecordKey = nullptr;
	if (PR_RECORD_KEY == lpProps[ePR_RECORD_KEY].ulPropTag)
	{
		lpRecordKey = &lpProps[ePR_RECORD_KEY].Value.bin;
	}

	auto szFileOut = BuildFileNameAndPath(
		szExt,
		szSubj,
		emptystring,
		lpRecordKey);

	MAPIFreeBuffer(lpProps);
	return szFileOut;
}

// The file name we generate should be shorter than MAX_PATH
// This includes our directory name too!
#define MAXSUBJ 25
#define MAXBIN 141
wstring BuildFileNameAndPath(
	_In_ const wstring& szExt,
	_In_ wstring szSubj,
	_In_ wstring szRootPath,
	_In_opt_ LPSBinary lpBin)
{
	auto hRes = S_OK;

	// set up the path portion of the output:
	WCHAR szShortPath[MAX_PATH] = { 0 };
	if (!szRootPath.empty())
	{
		size_t cchShortPath = NULL;
		// Use the short path to give us as much room as possible
		EC_D(cchShortPath, GetShortPathNameW(szRootPath.c_str(), szShortPath, _countof(szShortPath)));
		szRootPath = szShortPath;
		if (szRootPath.back() != L'\\')
		{
			szRootPath += L'\\';
		}
	}

	if (!szSubj.empty())
	{
		szSubj = SanitizeFileNameW(szSubj);
	}
	else
	{
		// We must have failed to get a subject before. Make one up.
		szSubj = L"UnknownSubject"; // STRING_OK
	}

	wstring szBin;
	if (lpBin && lpBin->cb)
	{
		szBin = L"_" + BinToHexString(lpBin, false);
	}

	auto szFileOut = szRootPath + szSubj + szBin + szExt;

	if (szFileOut.length() > MAX_PATH)
	{
		szFileOut = szRootPath + szSubj.substr(0, MAXSUBJ) + szBin.substr(0, MAXBIN) + szExt;
	}

	return szFileOut;
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

_Check_return_ HRESULT SaveFolderContentsToMSG(_In_ LPMAPIFOLDER lpFolder, _In_z_ LPCWSTR szPathName, bool bAssoc, bool bUnicode, HWND hWnd)
{
	auto hRes = S_OK;
	LPMAPITABLE lpFolderContents = nullptr;
	LPMESSAGE lpMessage = nullptr;
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

	if (!lpFolder || !szPathName) return MAPI_E_INVALID_PARAMETER;

	DebugPrint(DBGGeneric, L"SaveFolderContentsToMSG: Saving contents of folder to \"%ws\"\n", szPathName);

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

			pRows->aRow->lpProps[fldPR_ENTRYID].ulPropTag;

			if (PT_ERROR != PROP_TYPE(pRows->aRow->lpProps[fldPR_ENTRYID].ulPropTag))
			{
				DebugPrint(DBGGeneric, L"Source Message =\n");
				DebugPrintBinary(DBGGeneric, &pRows->aRow->lpProps[fldPR_ENTRYID].Value.bin);

				if (lpMessage) lpMessage->Release();
				lpMessage = nullptr;
				EC_H(CallOpenEntry(
					NULL,
					NULL,
					lpFolder,
					NULL,
					pRows->aRow->lpProps[fldPR_ENTRYID].Value.bin.cb,
					reinterpret_cast<LPENTRYID>(pRows->aRow->lpProps[fldPR_ENTRYID].Value.bin.lpb),
					NULL,
					MAPI_BEST_ACCESS,
					NULL,
					reinterpret_cast<LPUNKNOWN*>(&lpMessage)));
				if (!lpMessage) continue;

				auto szSubj = L"UnknownSubject"; // STRING_OK

				if (CheckStringProp(&pRows->aRow->lpProps[fldPR_SUBJECT_W], PT_UNICODE))
				{
					szSubj = pRows->aRow->lpProps[fldPR_SUBJECT_W].Value.lpszW;
				}

				auto szFileName = BuildFileNameAndPath(L".msg", szSubj, szPathName, &pRows->aRow->lpProps[fldPR_RECORD_KEY].Value.bin); // STRING_OK

				DebugPrint(DBGGeneric, L"Saving to = \"%ws\"\n", szFileName.c_str());

				EC_H(SaveToMSG(
					lpMessage,
					szFileName.c_str(),
					bUnicode,
					hWnd,
					false));

				DebugPrint(DBGGeneric, L"Message Saved\n");
			}
		}
	}

	if (pRows) FreeProws(pRows);
	if (lpMessage) lpMessage->Release();
	if (lpFolderContents) lpFolderContents->Release();
	return hRes;
}

_Check_return_ HRESULT WriteStreamToFile(_In_ LPSTREAM pStrmSrc, _In_z_ LPCWSTR szFileName)
{
	if (!pStrmSrc || !szFileName) return MAPI_E_INVALID_PARAMETER;

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
		NULL,
		&pStrmDest));

	if (pStrmDest)
	{
		pStrmSrc->Stat(&StatInfo, STATFLAG_NONAME);

		DebugPrint(DBGStream, L"WriteStreamToFile: Writing cb = %llu bytes\n", StatInfo.cbSize.QuadPart);

		EC_MAPI(pStrmSrc->CopyTo(pStrmDest,
			StatInfo.cbSize,
			NULL,
			NULL));

		// Commit changes to new stream
		EC_MAPI(pStrmDest->Commit(STGC_DEFAULT));

		pStrmDest->Release();
	}
	return hRes;
}

_Check_return_ HRESULT SaveToEML(_In_ LPMESSAGE lpMessage, _In_z_ LPCWSTR szFileName)
{
	auto hRes = S_OK;
	LPSTREAM pStrmSrc = nullptr;

	if (!lpMessage || !szFileName) return MAPI_E_INVALID_PARAMETER;
	DebugPrint(DBGGeneric, L"SaveToEML: Saving message to \"%ws\"\n", szFileName);

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

_Check_return_ HRESULT STDAPICALLTYPE MyStgCreateStorageEx(IN const WCHAR* pName,
	DWORD grfMode,
	DWORD stgfmt,
	DWORD grfAttrs,
	_In_ STGOPTIONS * pStgOptions,
	_Pre_null_ void * reserved,
	_In_ REFIID riid,
	_Out_ void ** ppObjectOpen)
{
	auto hRes = S_OK;
	if (!pName) return MAPI_E_INVALID_PARAMETER;

	if (pfnStgCreateStorageEx) hRes = pfnStgCreateStorageEx(
		pName,
		grfMode,
		stgfmt,
		grfAttrs,
		pStgOptions,
		reserved,
		riid,
		ppObjectOpen);
	// Fallback for NT4, which doesn't have StgCreateStorageEx
	else hRes = ::StgCreateDocfile(
		pName,
		grfMode,
		0,
		reinterpret_cast<LPSTORAGE*>(ppObjectOpen));
	return hRes;
}

_Check_return_ HRESULT CreateNewMSG(_In_z_ LPCWSTR szFileName, bool bUnicode, _Deref_out_opt_ LPMESSAGE* lppMessage, _Deref_out_opt_ LPSTORAGE* lppStorage)
{
	if (!szFileName || !lppMessage || !lppStorage) return MAPI_E_INVALID_PARAMETER;

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
				NULL,
				MAPIAllocateBuffer,
				MAPIAllocateMore,
				MAPIFreeBuffer,
				pMalloc,
				NULL,
				pStorage,
				NULL,
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

_Check_return_ HRESULT SaveToMSG(_In_ LPMESSAGE lpMessage, _In_z_ LPCWSTR szFileName, bool bUnicode, HWND hWnd, bool bAllowUI)
{
	auto hRes = S_OK;
	LPSTORAGE pStorage = nullptr;
	LPMESSAGE pIMsg = nullptr;

	if (!lpMessage || !szFileName) return MAPI_E_INVALID_PARAMETER;

	DebugPrint(DBGGeneric, L"SaveToMSG: Saving message to \"%ws\"\n", szFileName);

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

_Check_return_ HRESULT SaveToTNEF(_In_ LPMESSAGE lpMessage, _In_ LPADRBOOK lpAdrBook, _In_z_ LPCWSTR szFileName)
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

	if (!lpMessage || !lpAdrBook || !szFileName) return MAPI_E_INVALID_PARAMETER;
	DebugPrint(DBGGeneric, L"SaveToTNEF: Saving message to \"%ws\"\n", szFileName);

	LPSTREAM lpStream = nullptr;
	LPITNEF lpTNEF = nullptr;
	LPSTnefProblemArray lpError = nullptr;

	static auto dwKey = static_cast<WORD>(::GetTickCount());

	// Get a Stream interface on the input TNEF file
	EC_H(MyOpenStreamOnFile(
		MAPIAllocateBuffer,
		MAPIFreeBuffer,
		STGM_READWRITE | STGM_CREATE,
		szFileName,
		NULL,
		&lpStream));

	if (lpStream)
	{
		// Open TNEF stream
#pragma warning(push)
#pragma warning(disable:4616)
#pragma warning(disable:6276)
		EC_H(OpenTnefStreamEx(
			NULL,
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
				NULL,
				LPSPropTagArray(&lpPropTnefExcludeArray)
			));
			EC_MAPI(lpTNEF->AddProps(
				TNEF_PROP_EXCLUDE | TNEF_PROP_ATTACHMENTS_ONLY,
				0,
				NULL,
				LPSPropTagArray(&lpPropTnefExcludeArray)
			));

			EC_MAPI(lpTNEF->AddProps(
				TNEF_PROP_INCLUDE,
				0,
				NULL,
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

	EC_MAPI(HrGetOneProp(
		lpMessage,
		PR_HASATTACH,
		&pProps));

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
				NULL,
				NULL,
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

	EC_MAPI(HrGetOneProp(
		lpMessage,
		PR_HASATTACH,
		&pProps));

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
				NULL,
				NULL,
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
						NULL,
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

_Check_return_ HRESULT WriteEmbeddedMSGToFile(_In_ LPATTACH lpAttach, _In_z_ LPCWSTR szFileName, bool bUnicode, HWND hWnd)
{
	auto hRes = S_OK;
	LPMESSAGE lpAttachMsg = nullptr;

	if (!lpAttach || !szFileName) return MAPI_E_INVALID_PARAMETER;

	DebugPrint(DBGGeneric, L"WriteEmbeddedMSGToFile: Saving attachment to \"%ws\"\n", szFileName);

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

_Check_return_ HRESULT WriteAttachStreamToFile(_In_ LPATTACH lpAttach, _In_z_ LPCWSTR szFileName)
{
	auto hRes = S_OK;
	LPSTREAM pStrmSrc = nullptr;

	if (!lpAttach || !szFileName) return MAPI_E_INVALID_PARAMETER;

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
_Check_return_ HRESULT WriteOleToFile(_In_ LPATTACH lpAttach, _In_z_ LPCWSTR szFileName)
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
				szFileName,
				STGM_READWRITE | STGM_TRANSACTED | STGM_CREATE,
				0,
				&lpStorageDest));
			if (lpStorageDest)
			{
				EC_MAPI(lpStorageSrc->CopyTo(
					NULL,
					NULL,
					NULL,
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

		auto szFileName = SanitizeFileNameW(szName);

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
				EC_H(WriteAttachStreamToFile(lpAttach, file.c_str()));
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
				EC_H(WriteEmbeddedMSGToFile(lpAttach, file.c_str(), (MAPI_UNICODE == fMapiUnicode) ? true : false, hWnd));
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
				EC_H(WriteOleToFile(lpAttach, file.c_str()));
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
