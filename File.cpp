// file.cpp

#include "stdafx.h"
#include "File.h"
#include "InterpretProp.h"
#include "MAPIFunctions.h"
#include "FileDialogEx.h"
#include "MAPIProgress.h"
#include "guids.h"
#include "ImportProcs.h"
#include "MFCUtilityFunctions.h"
#include <shlobj.h>
#include "Dumpstore.h"

// Add current Entry ID to file name
_Check_return_ HRESULT AppendEntryID(_Inout_z_count_(cchFileName) LPWSTR szFileName, size_t cchFileName, _In_ LPSBinary lpBin, size_t cchMaxAppend)
{
	HRESULT hRes = S_OK;
	LPTSTR szBin = NULL;

	if (!lpBin || !lpBin->cb || !szFileName || cchMaxAppend <= 1) return MAPI_E_INVALID_PARAMETER;

	MyHexFromBin(
		lpBin->lpb,
		lpBin->cb,
		false,
		&szBin);

	if (szBin)
	{
		EC_H(StringCchCatNW(szFileName, cchFileName, L"_",1)); // STRING_OK
#ifdef UNICODE
		EC_H(StringCchCatNW(szFileName, cchFileName, szBin,cchMaxAppend-1));
#else
		LPWSTR szWideBin = NULL;
		EC_H(AnsiToUnicode(
			szBin,
			&szWideBin));
		if (szWideBin)
		{
			EC_H(StringCchCatNW(szFileName, cchFileName, szWideBin,cchMaxAppend-1));
		}
		delete[] szWideBin;
#endif
		delete[] szBin;
	}

	return hRes;
} // AppendEntryID

_Check_return_ HRESULT GetDirectoryPath(HWND hWnd, _Inout_z_ LPWSTR szPath)
{
	BROWSEINFOW BrowseInfo;
	LPITEMIDLIST lpItemIdList = NULL;
	HRESULT hRes = S_OK;

	if (!szPath) return MAPI_E_INVALID_PARAMETER;

	LPMALLOC lpMalloc = NULL;

	EC_H(SHGetMalloc(&lpMalloc));

	if (!lpMalloc) return hRes;

	memset(&BrowseInfo,NULL,sizeof(BROWSEINFO));

	szPath[0] = NULL;

	CStringW szInputString;
	EC_B(szInputString.LoadString(IDS_DIRPROMPT));

	BrowseInfo.hwndOwner = hWnd;
	BrowseInfo.lpszTitle = szInputString;
	BrowseInfo.pszDisplayName = szPath;
	BrowseInfo.ulFlags = BIF_USENEWUI | BIF_RETURNONLYFSDIRS;

	// Note - I don't initialize COM for this call because MAPIInitialize does this
	lpItemIdList = SHBrowseForFolderW(&BrowseInfo);
	if (lpItemIdList)
	{
		EC_B(SHGetPathFromIDListW(lpItemIdList,szPath));
		lpMalloc->Free(lpItemIdList);
	}
	else
	{
		hRes = MAPI_E_USER_CANCEL;
	}

	lpMalloc->Release();
	return hRes;
} // GetDirectoryPath

// Opens storage with best access
_Check_return_ HRESULT MyStgOpenStorage(_In_z_ LPCWSTR szMessageFile, bool bBestAccess, _Deref_out_ LPSTORAGE* lppStorage)
{
	if (!lppStorage) return MAPI_E_INVALID_PARAMETER;
	DebugPrint(DBGGeneric,_T("MyStgOpenStorage: Opening \"%ws\", bBestAccess == %s\n"),szMessageFile,bBestAccess?_T("True"):_T("False"));
	HRESULT		hRes = S_OK;
	ULONG		ulFlags = STGM_TRANSACTED;

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
} // MyStgOpenStorage

// Creates an LPMESSAGE on top of the MSG file
_Check_return_ HRESULT LoadMSGToMessage(_In_z_ LPCWSTR szMessageFile, _Deref_out_opt_ LPMESSAGE* lppMessage)
{
	if (!lppMessage) return MAPI_E_INVALID_PARAMETER;

	HRESULT		hRes = S_OK;
	LPSTORAGE	pStorage = NULL;
	LPMALLOC	lpMalloc = NULL;

	*lppMessage = NULL;

	// get memory allocation function
	lpMalloc = MAPIGetDefaultMalloc();

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
} // LoadMSGToMessage

// Loads the MSG file into an LPMESSAGE pointer, then copies it into the passed in message
// lpMessage must be created first
_Check_return_ HRESULT LoadFromMSG(_In_z_ LPCWSTR szMessageFile, _In_ LPMESSAGE lpMessage, HWND hWnd)
{
	HRESULT				hRes = S_OK;
	LPMESSAGE			pIMsg = NULL;

	// Specify properties to exclude in the copy operation. These are
	// the properties that Exchange excludes to save bits and time.
	// Should not be necessary to exclude these, but speeds the process
	// when a lot of messages are being copied.
	static const SizedSPropTagArray	(18, excludeTags) =
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

	EC_H(LoadMSGToMessage(szMessageFile,&pIMsg));

	if (pIMsg)
	{
		LPSPropProblemArray lpProblems = NULL;

		LPMAPIPROGRESS lpProgress = GetMAPIProgress(_T("IMAPIProp::CopyTo"), hWnd); // STRING_OK

		EC_MAPI(pIMsg->CopyTo(
			0,
			NULL,
			(LPSPropTagArray)&excludeTags,
			lpProgress ? (ULONG_PTR)hWnd : NULL,
			lpProgress,
			&IID_IMessage,
			lpMessage,
			lpProgress ? MAPI_DIALOG : 0,
			&lpProblems));

		if (lpProgress)
			lpProgress->Release();

		lpProgress = NULL;

		EC_PROBLEMARRAY(lpProblems);
		MAPIFreeBuffer(lpProblems);
		if (!FAILED(hRes))
		{
			EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
		}
	}

	if (pIMsg) pIMsg->Release();
	return hRes;
} // LoadFromMSG

// lpMessage must be created first
_Check_return_ HRESULT LoadFromTNEF(_In_z_ LPCWSTR szMessageFile, _In_ LPADRBOOK lpAdrBook, _In_ LPMESSAGE lpMessage)
{
	HRESULT				hRes = S_OK;
	LPSTREAM			lpStream = NULL;
	LPITNEF				lpTNEF = NULL;
	LPSTnefProblemArray	lpError = NULL;
	LPSTREAM			lpBodyStream = NULL;

	if (!szMessageFile || !lpAdrBook || !lpMessage) return MAPI_E_INVALID_PARAMETER;
	static WORD dwKey = (WORD)::GetTickCount();

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
		(LPTSTR) "winmail.dat", // STRING_OK - despite its signature, this function is ANSI only
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
			(LPSPropTagArray) &lpPropTnefExcludeArray,
			&lpError));

		EC_TNEFERR(lpError);
	}

	EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));

	if (lpBodyStream) lpBodyStream->Release();
	MAPIFreeBuffer(lpError);
	if (lpTNEF) lpTNEF->Release();
	if (lpStream) lpStream->Release();
	return hRes;
} // LoadFromTNEF

// Builds a file name out of the passed in message and extension
_Check_return_ HRESULT BuildFileName(_Inout_z_count_(cchFileOut) LPWSTR szFileOut,
									 size_t cchFileOut,
									 _In_z_count_(cchExt) LPCWSTR szExt,
									 size_t cchExt,
									 _In_ LPMESSAGE lpMessage)
{
	HRESULT hRes = S_OK;
	ULONG ulProps = NULL;
	LPSPropValue lpProps = NULL;
	LPWSTR szSubj = NULL;
	LPSBinary lpRecordKey = NULL;

	if (!lpMessage || !szFileOut) return MAPI_E_INVALID_PARAMETER;

	enum
	{
		ePR_SUBJECT_W,
		ePR_RECORD_KEY,
		NUM_COLS
	};
	static const SizedSPropTagArray(NUM_COLS,sptaMessageProps) =
	{
		NUM_COLS,
		PR_SUBJECT_W,
		PR_RECORD_KEY
	};

	// Get subject line of message
	// This will be used as the new file name.
	WC_H_GETPROPS(lpMessage->GetProps(
		(LPSPropTagArray) &sptaMessageProps,
		fMapiUnicode,
		&ulProps,
		&lpProps));
	hRes = S_OK;

	szFileOut[0] = NULL;
	if (CheckStringProp(&lpProps[ePR_SUBJECT_W],PT_UNICODE))
	{
		szSubj = lpProps[ePR_SUBJECT_W].Value.lpszW;
	}
	if (PR_RECORD_KEY == lpProps[ePR_RECORD_KEY].ulPropTag)
	{
		lpRecordKey = &lpProps[ePR_RECORD_KEY].Value.bin;
	}

	EC_H(BuildFileNameAndPath(
		szFileOut,
		cchFileOut,
		szExt,
		cchExt,
		szSubj,
		lpRecordKey,
		NULL));

	MAPIFreeBuffer(lpProps);
	return hRes;
} // BuildFileName

// Problem here is that cchFileOut can't be longer than MAX_PATH
// So the file name we generate must be shorter than MAX_PATH
// This includes our directory name too!
// So directory is part of the input and output now
#define MAXSUBJ 25
#define MAXBIN 141
_Check_return_ HRESULT BuildFileNameAndPath(_Inout_z_count_(cchFileOut) LPWSTR szFileOut,
											size_t cchFileOut,
											_In_z_count_(cchExt) LPCWSTR szExt,
											size_t cchExt,
											_In_opt_z_ LPCWSTR szSubj,
											_In_opt_ LPSBinary lpBin,
											_In_opt_z_ LPCWSTR szRootPath)
{
	HRESULT			hRes = S_OK;

	if (!szFileOut) return MAPI_E_INVALID_PARAMETER;
	if (cchFileOut > MAX_PATH) return MAPI_E_INVALID_PARAMETER;

	szFileOut[0] = L'\0'; // initialize our string to NULL
	size_t cchCharRemaining = cchFileOut;

	size_t cchRootPath = NULL;

	// set up the path portion of the output:
	if (szRootPath)
	{
		// Use the short path to give us as much room as possible
		EC_D(cchRootPath,GetShortPathNameW(szRootPath, szFileOut, (DWORD)cchFileOut));
		// stuff a slash in there if we need one
		if (cchRootPath+1 < cchFileOut && szFileOut[cchRootPath-1] != L'\\')
		{
			szFileOut[cchRootPath] = L'\\';
			szFileOut[cchRootPath+1] = L'\0';
			cchRootPath++;
		}
		cchCharRemaining -= cchRootPath;
	}

	// We now have cchCharRemaining characters in which to work
	// Suppose this is 0? Need at least 12 for an 8.3 name

	size_t cchBin = 0;
	if (lpBin) cchBin = (2*lpBin->cb)+1; // bin + '_'

	size_t cchSubj = 14; // length of 'UnknownSubject'
	if (szSubj)
	{
		EC_H(StringCchLengthW(szSubj,STRSAFE_MAX_CCH,&cchSubj));
	}

	if (cchCharRemaining < cchSubj + cchBin + cchExt + 1)
	{
		// don't have enough space - need to shorten things:
		if (cchSubj > MAXSUBJ) cchSubj = MAXSUBJ;
		if (cchBin  > MAXBIN)  cchBin  = MAXBIN;
	}
	if (cchCharRemaining < cchSubj + cchBin + cchExt + 1)
	{
		// still don't have enough space - need to shorten things:
		// TODO: generate a unique 8.3 name and return it
		return MAPI_E_INVALID_PARAMETER;
	}
	else
	{
		if (szSubj)
		{
			EC_H(SanitizeFileNameW(
				szFileOut + cchRootPath,
				cchCharRemaining,
				szSubj,
				cchSubj));
		}
		else
		{
			// We must have failed to get a subject before. Make one up.
			EC_H(StringCchCopyW(szFileOut + cchRootPath, cchCharRemaining, L"UnknownSubject")); // STRING_OK
		}

		if (lpBin && lpBin->cb)
		{
			EC_H(AppendEntryID(szFileOut,cchFileOut,lpBin,cchBin));
		}

		// Add our extension
		if (szExt && cchExt)
		{
			EC_H(StringCchCatNW(szFileOut, cchFileOut, szExt,cchExt));
		}
	}

	return hRes;
} // BuildFileNameAndPath

// Takes szFileIn and copies it to szFileOut, replacing non file system characters with underscores
// Do NOT call with full path - just file names
// Resulting string will have no more than cchCharsToCopy characters
_Check_return_ HRESULT SanitizeFileNameA(
	_Inout_z_count_(cchFileOut) LPSTR szFileOut, // output buffer
	size_t cchFileOut, // length of output buffer
	_In_z_ LPCSTR szFileIn, // File name in
	size_t cchCharsToCopy)
{
	HRESULT hRes = S_OK;
	LPSTR szCur = NULL;

	EC_H(StringCchCopyNA(szFileOut, cchFileOut, szFileIn, cchCharsToCopy));
	while (NULL != (szCur = strpbrk(szFileOut,"^&*-+=[]\\|;:\",<>/?"))) // STRING_OK
	{
		*szCur = '_';
	}
	return hRes;
} // SanitizeFileNameA

_Check_return_ HRESULT SanitizeFileNameW(
	_Inout_z_count_(cchFileOut) LPWSTR szFileOut, // output buffer
	size_t cchFileOut, // length of output buffer
	_In_z_ LPCWSTR szFileIn, // File name in
	size_t cchCharsToCopy)
{
	HRESULT hRes = S_OK;
	LPWSTR szCur = NULL;

	EC_H(StringCchCopyNW(szFileOut, cchFileOut, szFileIn, cchCharsToCopy));
	while (NULL != (szCur = wcspbrk(szFileOut,L"^&*-+=[]\\|;:\",<>/?\r\n"))) // STRING_OK
	{
		*szCur = L'_';
	}
	return hRes;
} // SanitizeFileNameW

void SaveFolderContentsToTXT(_In_ LPMDB lpMDB, _In_ LPMAPIFOLDER lpFolder, bool bRegular, bool bAssoc, bool bDescend, HWND hWnd)
{
	HRESULT hRes = S_OK;
	WCHAR szDir[MAX_PATH] = {0};

	WC_H(GetDirectoryPath(hWnd, szDir));

	if (S_OK == hRes && szDir[0])
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
	HRESULT			hRes = S_OK;
	LPMAPITABLE		lpFolderContents = NULL;
	LPMESSAGE		lpMessage = NULL;
	LPSRowSet		pRows = NULL;

	enum
	{
		fldPR_ENTRYID,
		fldPR_SUBJECT_W,
		fldPR_RECORD_KEY,
		fldNUM_COLS
	};

	static const SizedSPropTagArray(fldNUM_COLS,fldCols) =
	{
		fldNUM_COLS,
		PR_ENTRYID,
		PR_SUBJECT_W,
		PR_RECORD_KEY
	};

	if (!lpFolder || !szPathName) return MAPI_E_INVALID_PARAMETER;

	DebugPrint(DBGGeneric,_T("SaveFolderContentsToMSG: Saving contents of folder to \"%ws\"\n"),szPathName);

	EC_MAPI(lpFolder->GetContentsTable(
		fMapiUnicode | (bAssoc?MAPI_ASSOCIATED:NULL),
		&lpFolderContents));

	if (lpFolderContents)
	{
		EC_MAPI(lpFolderContents->SetColumns((LPSPropTagArray)&fldCols, TBL_BATCH));

		if (!FAILED(hRes)) for (;;)
		{
			hRes = S_OK;
			if (pRows) FreeProws(pRows);
			pRows = NULL;
			EC_MAPI(lpFolderContents->QueryRows(
				1,
				NULL,
				&pRows));
			if (FAILED(hRes) || !pRows || (pRows && !pRows->cRows)) break;

			pRows->aRow->lpProps[fldPR_ENTRYID].ulPropTag;

			if (PT_ERROR != PROP_TYPE(pRows->aRow->lpProps[fldPR_ENTRYID].ulPropTag))
			{
				DebugPrint(DBGGeneric,_T("Source Message =\n"));
				DebugPrintBinary(DBGGeneric,&pRows->aRow->lpProps[fldPR_ENTRYID].Value.bin);

				if (lpMessage) lpMessage->Release();
				lpMessage = NULL;
				EC_H(CallOpenEntry(
					NULL,
					NULL,
					lpFolder,
					NULL,
					pRows->aRow->lpProps[fldPR_ENTRYID].Value.bin.cb,
					(LPENTRYID) pRows->aRow->lpProps[fldPR_ENTRYID].Value.bin.lpb,
					NULL,
					MAPI_BEST_ACCESS,
					NULL,
					(LPUNKNOWN*)&lpMessage));
				if (!lpMessage) continue;

				WCHAR szFileName[MAX_PATH] = {0};

				LPCWSTR szSubj = L"UnknownSubject"; // STRING_OK

				if (CheckStringProp(&pRows->aRow->lpProps[fldPR_SUBJECT_W],PT_UNICODE))
				{
					szSubj = pRows->aRow->lpProps[fldPR_SUBJECT_W].Value.lpszW;
				}
				EC_H(BuildFileNameAndPath(szFileName,_countof(szFileName),L".msg",4,szSubj,&pRows->aRow->lpProps[fldPR_RECORD_KEY].Value.bin,szPathName)); // STRING_OK

				DebugPrint(DBGGeneric,_T("Saving to = \"%ws\"\n"),szFileName);

				EC_H(SaveToMSG(
					lpMessage,
					szFileName,
					bUnicode,
					hWnd));

				DebugPrint(DBGGeneric,_T("Message Saved\n"));
			}
		}
	}

	if (pRows) FreeProws(pRows);
	if (lpMessage) lpMessage->Release();
	if (lpFolderContents) lpFolderContents->Release();
	return hRes;
} // SaveFolderContentsToMSG

_Check_return_ HRESULT WriteStreamToFile(_In_ LPSTREAM pStrmSrc, _In_z_ LPCWSTR szFileName)
{
	if (!pStrmSrc || !szFileName) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;
	LPSTREAM pStrmDest = NULL;
	STATSTG StatInfo = {0};

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

		DebugPrint(DBGStream,_T("WriteStreamToFile: Writing cb = %llu bytes\n"), StatInfo.cbSize.QuadPart);

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
	HRESULT hRes = S_OK;
	LPSTREAM		pStrmSrc = NULL;

	if (!lpMessage || !szFileName) return MAPI_E_INVALID_PARAMETER;
	DebugPrint(DBGGeneric,_T("SaveToEML: Saving message to \"%ws\"\n"),szFileName);

	// Open the property of the attachment
	// containing the file data
	EC_MAPI(lpMessage->OpenProperty(
		PR_INTERNET_CONTENT,
		(LPIID)&IID_IStream,
		0,
		NULL, // MAPI_MODIFY is not needed
		(LPUNKNOWN *)&pStrmSrc));
	if (FAILED(hRes))
	{
		if (MAPI_E_NOT_FOUND == hRes)
		{
			DebugPrint(DBGGeneric,_T("No internet content found\n"));
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
} // SaveToEML

_Check_return_ HRESULT STDAPICALLTYPE MyStgCreateStorageEx(IN const WCHAR* pName,
														   DWORD grfMode,
														   DWORD stgfmt,
														   DWORD grfAttrs,
														   _In_ STGOPTIONS * pStgOptions,
														   _Pre_null_ void * reserved,
														   _In_ REFIID riid,
														   _Out_ void ** ppObjectOpen)
{
	HRESULT hRes = S_OK;
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
		(LPSTORAGE*) ppObjectOpen);
	return hRes;
} // MyStgCreateStorageEx

_Check_return_ HRESULT CreateNewMSG(_In_z_ LPCWSTR szFileName, bool bUnicode, _Deref_out_opt_ LPMESSAGE* lppMessage, _Deref_out_opt_ LPSTORAGE* lppStorage)
{
	if (!szFileName || !lppMessage || !lppStorage) return MAPI_E_INVALID_PARAMETER;

	HRESULT		hRes = S_OK;
	LPSTORAGE	pStorage = NULL;
	LPMESSAGE	pIMsg = NULL;
	LPMALLOC	pMalloc = NULL;

	*lppMessage = NULL;
	*lppStorage = NULL;

	// get memory allocation function
	pMalloc = MAPIGetDefaultMalloc();
	if (pMalloc)
	{
		STGOPTIONS myOpts = {0};

		myOpts.usVersion = 1; // STGOPTIONS_VERSION
		myOpts.ulSectorSize = 4096;

		// Open the compound file
		EC_H(MyStgCreateStorageEx(
			szFileName,
			STGM_READWRITE | STGM_TRANSACTED | STGM_CREATE,
			STGFMT_DOCFILE,
			0, // FILE_FLAG_NO_BUFFERING,
			&myOpts,
			0,
			__uuidof(IStorage),
			(LPVOID*)&pStorage));
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
				bUnicode?MAPI_UNICODE:0,
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
} // CreateNewMSG

_Check_return_ HRESULT SaveToMSG(_In_ LPMESSAGE lpMessage, _In_z_ LPCWSTR szFileName, bool bUnicode, HWND hWnd)
{
	HRESULT hRes = S_OK;
	LPSTORAGE pStorage = NULL;
	LPMESSAGE pIMsg = NULL;

	if (!lpMessage || !szFileName) return MAPI_E_INVALID_PARAMETER;

	DebugPrint(DBGGeneric,_T("SaveToMSG: Saving message to \"%ws\"\n"),szFileName);

	EC_H(CreateNewMSG(szFileName,bUnicode,&pIMsg,&pStorage));
	if (pIMsg && pStorage)
	{
		// Specify properties to exclude in the copy operation. These are
		// the properties that Exchange excludes to save bits and time.
		// Should not be necessary to exclude these, but speeds the process
		// when a lot of messages are being copied.
		static const SizedSPropTagArray (7, excludeTags) =
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

		LPSPropProblemArray lpProblems = NULL;

		// copy message properties to IMessage object opened on top of IStorage.
		LPMAPIPROGRESS lpProgress = GetMAPIProgress(_T("IMAPIProp::CopyTo"), hWnd); // STRING_OK

		EC_MAPI(lpMessage->CopyTo(0, NULL,
			(LPSPropTagArray)&excludeTags,
			lpProgress ? (ULONG_PTR)hWnd : NULL,
			lpProgress,
			(LPIID)&IID_IMessage,
			pIMsg,
			lpProgress ? MAPI_DIALOG : 0,
			&lpProblems));

		if (lpProgress)
			lpProgress->Release();

		lpProgress = NULL;

		EC_PROBLEMARRAY(lpProblems);
		MAPIFreeBuffer(lpProblems);

		// save changes to IMessage object.
		EC_MAPI(pIMsg->SaveChanges(KEEP_OPEN_READWRITE));

		// save changes in storage of new doc file
		EC_MAPI(pStorage->Commit(STGC_DEFAULT));
	}

	if (pStorage) pStorage->Release();
	if (pIMsg) pIMsg->Release();

	return hRes;
} // SaveToMSG

_Check_return_ HRESULT SaveToTNEF(_In_ LPMESSAGE lpMessage, _In_ LPADRBOOK lpAdrBook, _In_z_ LPCWSTR szFileName)
{
	HRESULT hRes = S_OK;

	enum
	{
		ulNumTNEFIncludeProps = 2
	};
	static const SizedSPropTagArray(ulNumTNEFIncludeProps, lpPropTnefIncludeArray ) =
	{
		ulNumTNEFIncludeProps,
		PR_MESSAGE_RECIPIENTS,
		PR_ATTACH_DATA_BIN
	};

	enum
	{
		ulNumTNEFExcludeProps = 1
	};
	static const SizedSPropTagArray(ulNumTNEFExcludeProps, lpPropTnefExcludeArray ) =
	{
		ulNumTNEFExcludeProps,
		PR_URL_COMP_NAME
	};

	if (!lpMessage || !lpAdrBook || !szFileName) return MAPI_E_INVALID_PARAMETER;
	DebugPrint(DBGGeneric,_T("SaveToTNEF: Saving message to \"%ws\"\n"),szFileName);

	LPSTREAM			lpStream	= NULL;
	LPITNEF				lpTNEF		= NULL;
	LPSTnefProblemArray	lpError		= NULL;

	static WORD dwKey = (WORD)::GetTickCount();

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
			(LPTSTR) "winmail.dat", // STRING_OK - despite its signature, this function is ANSI only
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
				(LPSPropTagArray) &lpPropTnefExcludeArray
				));
			EC_MAPI(lpTNEF->AddProps(
				TNEF_PROP_EXCLUDE | TNEF_PROP_ATTACHMENTS_ONLY,
				0,
				NULL,
				(LPSPropTagArray) &lpPropTnefExcludeArray
				));

			EC_MAPI(lpTNEF->AddProps(
				TNEF_PROP_INCLUDE,
				0,
				NULL,
				(LPSPropTagArray) &lpPropTnefIncludeArray
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
} // SaveToTNEF

_Check_return_ HRESULT DeleteAttachments(_In_ LPMESSAGE lpMessage, _In_opt_z_ LPCTSTR szAttName, HWND hWnd)
{
	LPSPropValue	pProps = NULL;
	HRESULT			hRes = S_OK;
	LPMAPITABLE		lpAttTbl = NULL;
	LPSRowSet		pRows = NULL;
	ULONG			iRow;

	enum
	{
		ATTACHNUM,
		ATTACHNAME,
		NUM_COLS
	};
	static const SizedSPropTagArray(NUM_COLS,sptAttachTableCols) =
	{
		NUM_COLS,
		PR_ATTACH_NUM,
		PR_ATTACH_LONG_FILENAME
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
			(LPUNKNOWN *) &lpAttTbl));

		if (lpAttTbl)
		{
			// I would love to just pass a restriction
			// to HrQueryAllRows for the file name.  However,
			// we don't support restricting attachment tables (EMSMDB32 that is)
			// So I have to compare the strings myself (see below)
			EC_MAPI(HrQueryAllRows(lpAttTbl,
				(LPSPropTagArray) &sptAttachTableCols,
				NULL,
				NULL,
				0,
				&pRows));

			if (pRows)
			{
				bool bDirty = false;

				if (!FAILED(hRes)) for (iRow = 0; iRow < pRows -> cRows; iRow++)
				{
					hRes = S_OK;

					if (PR_ATTACH_NUM != pRows->aRow[iRow].lpProps[ATTACHNUM].ulPropTag) continue;

					if (szAttName)
					{
						if (PR_ATTACH_LONG_FILENAME != pRows->aRow[iRow].lpProps[ATTACHNAME].ulPropTag ||
							lstrcmpi(szAttName, pRows->aRow[iRow].lpProps[ATTACHNAME].Value.LPSZ) != 0)
							continue;
					}

					// Open the attachment
					LPMAPIPROGRESS lpProgress = GetMAPIProgress(_T("IMessage::DeleteAttach"), hWnd); // STRING_OK

					EC_MAPI(lpMessage->DeleteAttach(
						pRows->aRow[iRow].lpProps[ATTACHNUM].Value.l,
						lpProgress ? (ULONG_PTR)hWnd : NULL,
						lpProgress,
						lpProgress ? ATTACH_DIALOG : 0));

					if (SUCCEEDED(hRes))
						bDirty = true;

					if (lpProgress)
						lpProgress->Release();

					lpProgress = NULL;
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
	if (lpAttTbl) lpAttTbl -> Release();
	MAPIFreeBuffer(pProps);

	return hRes;
} // DeleteAllAttachments

#ifndef MRMAPI
_Check_return_ HRESULT WriteAttachmentsToFile(_In_ LPMESSAGE lpMessage, HWND hWnd)
{
	LPSPropValue	pProps = NULL;
	HRESULT			hRes = S_OK;
	LPMAPITABLE		lpAttTbl = NULL;
	LPSRowSet		pRows = NULL;
	ULONG			iRow;
	LPATTACH		lpAttach = NULL;

	enum
	{
		ATTACHNUM,
		NUM_COLS
	};
	static const SizedSPropTagArray(NUM_COLS,sptAttachTableCols) =
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
			(LPUNKNOWN *) &lpAttTbl));

		if (lpAttTbl)
		{
			EC_MAPI(HrQueryAllRows(lpAttTbl,
				(LPSPropTagArray) &sptAttachTableCols,
				NULL,
				NULL,
				0,
				&pRows));

			if (pRows)
			{
				if (!FAILED(hRes)) for (iRow = 0; iRow < pRows -> cRows; iRow++)
				{
					lpAttach = NULL;

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
						lpAttach = NULL;
						if (S_OK != hRes && iRow != pRows->cRows-1)
						{
							if (bShouldCancel(NULL,hRes)) break;
							hRes = S_OK;
						}
					}
				}
				if (pRows) FreeProws(pRows);
			}
			lpAttTbl -> Release();
		}
	}

	MAPIFreeBuffer(pProps);

	return hRes;
} // WriteAttachmentsToFile
#endif

_Check_return_ HRESULT WriteEmbeddedMSGToFile(_In_ LPATTACH lpAttach, _In_z_ LPCWSTR szFileName, bool bUnicode, HWND hWnd)
{
	HRESULT			hRes = S_OK;
	LPMESSAGE		lpAttachMsg = NULL;

	if (!lpAttach || !szFileName) return MAPI_E_INVALID_PARAMETER;

	DebugPrint(DBGGeneric,_T("WriteEmbeddedMSGToFile: Saving attachment to \"%ws\"\n"),szFileName);

	EC_MAPI(lpAttach->OpenProperty(
		PR_ATTACH_DATA_OBJ,
		(LPIID)&IID_IMessage,
		0,
		NULL, // MAPI_MODIFY is not needed
		(LPUNKNOWN *)&lpAttachMsg));

	if (lpAttachMsg)
	{
		EC_H(SaveToMSG(lpAttachMsg,szFileName, bUnicode, hWnd));
		lpAttachMsg->Release();
	}

	return hRes;
} // WriteEmbeddedMSGToFile

_Check_return_ HRESULT WriteAttachStreamToFile(_In_ LPATTACH lpAttach, _In_z_ LPCWSTR szFileName)
{
	HRESULT			hRes = S_OK;
	LPSTREAM		pStrmSrc = NULL;

	if (!lpAttach || !szFileName) return MAPI_E_INVALID_PARAMETER;

	// Open the property of the attachment
	// containing the file data
	WC_MAPI(lpAttach->OpenProperty(
		PR_ATTACH_DATA_BIN,
		(LPIID)&IID_IStream,
		0,
		NULL, // MAPI_MODIFY is not needed
		(LPUNKNOWN *)&pStrmSrc));
	if (FAILED(hRes))
	{
		if (MAPI_E_NOT_FOUND == hRes)
		{
			DebugPrint(DBGGeneric,_T("No attachments found. Maybe the attachment was a message?\n"));
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
} // WriteAttachStreamToFile

// Pretty sure this covers all OLE attachments - we don't need to look at PR_ATTACH_TAG
_Check_return_ HRESULT WriteOleToFile(_In_ LPATTACH lpAttach, _In_z_ LPCWSTR szFileName)
{
	HRESULT			hRes = S_OK;
	LPSTORAGE		lpStorageSrc = NULL;
	LPSTORAGE		lpStorageDest = NULL;
	LPSTREAM		pStrmSrc = NULL;

	// Open the property of the attachment containing the OLE data
	// Try to get it as an IStreamDocFile file first as that will be faster
	WC_MAPI(lpAttach->OpenProperty(
		PR_ATTACH_DATA_OBJ,
		(LPIID)&IID_IStreamDocfile,
		0,
		NULL,
		(LPUNKNOWN *)&pStrmSrc));

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
			(LPIID)&IID_IStorage,
			0,
			NULL,
			(LPUNKNOWN *)&lpStorageSrc));

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
} // WriteOleToFile

#ifndef MRMAPI
_Check_return_ HRESULT WriteAttachmentToFile(_In_ LPATTACH lpAttach, HWND hWnd)
{
	HRESULT			hRes = S_OK;
	LPSPropValue	lpProps = NULL;
	ULONG			ulProps = 0;
	WCHAR			szFileName[MAX_PATH] = {0};
	INT_PTR			iDlgRet = 0;

	enum
	{
		ATTACH_METHOD,
		ATTACH_LONG_FILENAME_W,
		ATTACH_FILENAME_W,
		DISPLAY_NAME_W,
		NUM_COLS
	};
	static const SizedSPropTagArray(NUM_COLS,sptaAttachProps) =
	{
		NUM_COLS,
		PR_ATTACH_METHOD,
		PR_ATTACH_LONG_FILENAME_W,
		PR_ATTACH_FILENAME_W,
		PR_DISPLAY_NAME_W
	};

	if (!lpAttach) return MAPI_E_INVALID_PARAMETER;

	DebugPrint(DBGGeneric,_T("WriteAttachmentToFile: Saving attachment.\n"));

	// Get required properties from the message
	EC_H_GETPROPS(lpAttach->GetProps(
		(LPSPropTagArray) &sptaAttachProps, // property tag array
		fMapiUnicode, // flags
		&ulProps, // Count of values returned
		&lpProps)); // Values returned

	if (lpProps)
	{
		LPCWSTR szName = L"Unknown"; // STRING_OK

		// Get a file name to use
		if (CheckStringProp(&lpProps[ATTACH_LONG_FILENAME_W],PT_UNICODE))
		{
			szName = lpProps[ATTACH_LONG_FILENAME_W].Value.lpszW;
		}
		else if (CheckStringProp(&lpProps[ATTACH_FILENAME_W],PT_UNICODE))
		{
			szName = lpProps[ATTACH_FILENAME_W].Value.lpszW;
		}
		else if (CheckStringProp(&lpProps[DISPLAY_NAME_W],PT_UNICODE))
		{
			szName = lpProps[DISPLAY_NAME_W].Value.lpszW;
		}

		EC_H(SanitizeFileNameW(szFileName,_countof(szFileName),szName,_countof(szFileName)));

		// Get File Name
		switch(lpProps[ATTACH_METHOD].Value.l)
		{
		case ATTACH_BY_VALUE:
		case ATTACH_BY_REFERENCE:
		case ATTACH_BY_REF_RESOLVE:
		case ATTACH_BY_REF_ONLY:
			{
				CStringW szFileSpec;
				EC_B(szFileSpec.LoadString(IDS_ALLFILES));

				CFileDialogExW dlgFilePicker;

				DebugPrint(DBGGeneric,_T("WriteAttachmentToFile: Prompting with \"%ws\"\n"),szFileName);

				EC_D_DIALOG(dlgFilePicker.DisplayDialog(
					false,
					L"txt", // STRING_OK
					szFileName,
					OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
					szFileSpec));
				if (iDlgRet == IDOK)
				{
					EC_H(WriteAttachStreamToFile(lpAttach,dlgFilePicker.GetFileName()));
				}
			}
			break;
		case ATTACH_EMBEDDED_MSG:
			// Get File Name
			{
				CStringW szFileSpec;
				EC_B(szFileSpec.LoadString(IDS_MSGFILES));

				CFileDialogExW dlgFilePicker;

				DebugPrint(DBGGeneric,_T("WriteAttachmentToFile: Prompting with \"%ws\"\n"),szFileName);

				EC_D_DIALOG(dlgFilePicker.DisplayDialog(
					false,
					L"msg", // STRING_OK
					szFileName,
					OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
					szFileSpec));
				if (iDlgRet == IDOK)
				{
					EC_H(WriteEmbeddedMSGToFile(lpAttach,dlgFilePicker.GetFileName(), (MAPI_UNICODE == fMapiUnicode)?true:false, hWnd));
				}
			}
			break;
		case ATTACH_OLE:
			{
				CStringW szFileSpec;
				EC_B(szFileSpec.LoadString(IDS_ALLFILES));

				CFileDialogExW dlgFilePicker;

				DebugPrint(DBGGeneric,_T("WriteAttachmentToFile: Prompting with \"%ws\"\n"),szFileName);
				EC_D_DIALOG(dlgFilePicker.DisplayDialog(
					false,
					NULL,
					szFileName,
					OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
					szFileSpec));
				if (iDlgRet == IDOK)
				{
					EC_H(WriteOleToFile(lpAttach,dlgFilePicker.GetFileName()));
				}
			}
			break;
		default:
			ErrDialog(__FILE__,__LINE__,IDS_EDUNKNOWNATTMETHOD); break;
		}
	}
	if (iDlgRet == IDCANCEL) hRes = MAPI_E_USER_CANCEL;

	MAPIFreeBuffer(lpProps);
	return hRes;
} // WriteAttachmentToFile
#endif
