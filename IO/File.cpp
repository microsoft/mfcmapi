#include <StdAfx.h>
#include <IO/File.h>
#include <Interpret/InterpretProp.h>
#include <MAPI/MAPIFunctions.h>
#include <Interpret/String.h>
#include <MAPI/MAPIProgress.h>
#include <Interpret/Guids.h>
#include <ImportProcs.h>
#ifndef MRMAPI
#include <UI/FileDialogEx.h>
#include <UI/Dialogs/MFCUtilityFunctions.h>
#endif
#include <ShlObj.h>
#include <MAPI/MAPIProcessor/DumpStore.h>
#include <UI/Dialogs/Editors/Editor.h>

namespace file
{
	std::wstring ShortenPath(const std::wstring& path)
	{
		if (!path.empty())
		{
			WCHAR szShortPath[MAX_PATH] = {};
			// Use the short path to give us as much room as possible
			WC_DS(GetShortPathNameW(path.c_str(), szShortPath, _countof(szShortPath)));
			std::wstring ret = szShortPath;
			if (ret.back() != L'\\')
			{
				ret += L'\\';
			}

			return ret;
		}

		return path;
	}

	std::wstring GetDirectoryPath(HWND hWnd)
	{
		WCHAR szPath[MAX_PATH] = {0};
		BROWSEINFOW BrowseInfo = {};
		LPMALLOC lpMalloc = nullptr;

		EC_H2S(SHGetMalloc(&lpMalloc));

		if (!lpMalloc) return strings::emptystring;

		auto szInputString = strings::loadstring(IDS_DIRPROMPT);

		BrowseInfo.hwndOwner = hWnd;
		BrowseInfo.lpszTitle = szInputString.c_str();
		BrowseInfo.pszDisplayName = szPath;
		BrowseInfo.ulFlags = BIF_USENEWUI | BIF_RETURNONLYFSDIRS;

		// Note - I don't initialize COM for this call because MAPIInitialize does this
		const auto lpItemIdList = SHBrowseForFolderW(&BrowseInfo);
		if (lpItemIdList)
		{
			EC_BS(SHGetPathFromIDListW(lpItemIdList, szPath));
			lpMalloc->Free(lpItemIdList);
		}

		lpMalloc->Release();

		auto path = ShortenPath(szPath);
		if (path.length() >= MAXMSGPATH)
		{
			error::ErrDialog(__FILE__, __LINE__, IDS_EDPATHTOOLONG, path.length(), MAXMSGPATH);
			return strings::emptystring;
		}

		return path;
	}

	// Opens storage with best access
	_Check_return_ HRESULT
	MyStgOpenStorage(_In_ const std::wstring& szMessageFile, bool bBestAccess, _Deref_out_ LPSTORAGE* lppStorage)
	{
		if (!lppStorage) return MAPI_E_INVALID_PARAMETER;
		output::DebugPrint(
			DBGGeneric,
			L"MyStgOpenStorage: Opening \"%ws\", bBestAccess == %ws\n",
			szMessageFile.c_str(),
			bBestAccess ? L"True" : L"False");
		ULONG ulFlags = STGM_TRANSACTED;

		if (bBestAccess) ulFlags |= STGM_READWRITE;

		auto hRes = WC_H2(::StgOpenStorage(szMessageFile.c_str(), nullptr, ulFlags, nullptr, 0, lppStorage));

		// If we asked for best access (read/write) and didn't get it, then try it without readwrite
		if (STG_E_ACCESSDENIED == hRes && !*lppStorage && bBestAccess)
		{
			hRes = EC_H2(MyStgOpenStorage(szMessageFile, false, lppStorage));
		}

		return hRes;
	}

	// Creates an LPMESSAGE on top of the MSG file
	_Check_return_ HRESULT
	LoadMSGToMessage(_In_ const std::wstring& szMessageFile, _Deref_out_opt_ LPMESSAGE* lppMessage)
	{
		if (!lppMessage) return MAPI_E_INVALID_PARAMETER;

		auto hRes = S_OK;

		*lppMessage = nullptr;

		// get memory allocation function
		const auto lpMalloc = MAPIGetDefaultMalloc();
		if (lpMalloc)
		{
			LPSTORAGE pStorage = nullptr;
			// Open the compound file
			hRes = EC_H2(MyStgOpenStorage(szMessageFile, true, &pStorage));

			if (pStorage)
			{
				// Open an IMessage interface on an IStorage object
				hRes = EC_MAPI2(OpenIMsgOnIStg(
					nullptr,
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
				pStorage->Release();
			}
		}

		return hRes;
	}

	// Loads the MSG file into an LPMESSAGE pointer, then copies it into the passed in message
	// lpMessage must be created first
	_Check_return_ HRESULT LoadFromMSG(_In_ const std::wstring& szMessageFile, _In_ LPMESSAGE lpMessage, HWND hWnd)
	{
		// Specify properties to exclude in the copy operation. These are
		// the properties that Exchange excludes to save bits and time.
		// Should not be necessary to exclude these, but speeds the process
		// when a lot of messages are being copied.
		static const SizedSPropTagArray(18, excludeTags) = {18,
															{PR_REPLICA_VERSION,
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
															 PR_HAS_DAMS}};

		LPMESSAGE pIMsg = nullptr;
		auto hRes = EC_H2(LoadMSGToMessage(szMessageFile, &pIMsg));

		if (pIMsg)
		{
			LPSPropProblemArray lpProblems = nullptr;

			LPMAPIPROGRESS lpProgress = mapi::mapiui::GetMAPIProgress(L"IMAPIProp::CopyTo", hWnd); // STRING_OK

			hRes = EC_MAPI2(pIMsg->CopyTo(
				0,
				nullptr,
				LPSPropTagArray(&excludeTags),
				lpProgress ? reinterpret_cast<ULONG_PTR>(hWnd) : NULL,
				lpProgress,
				&IID_IMessage,
				lpMessage,
				lpProgress ? MAPI_DIALOG : 0,
				&lpProblems));

			if (lpProgress) lpProgress->Release();

			EC_PROBLEMARRAY(lpProblems);
			MAPIFreeBuffer(lpProblems);
			if (!FAILED(hRes))
			{
				hRes = EC_MAPI2(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
			}

			pIMsg->Release();
		}

		return hRes;
	}

	// lpMessage must be created first
	_Check_return_ HRESULT
	LoadFromTNEF(_In_ const std::wstring& szMessageFile, _In_ LPADRBOOK lpAdrBook, _In_ LPMESSAGE lpMessage)
	{
		if (szMessageFile.empty() || !lpAdrBook || !lpMessage) return MAPI_E_INVALID_PARAMETER;

		LPSTREAM lpStream = nullptr;
		// Get a Stream interface on the input TNEF file
		auto hRes =
			EC_H2(mapi::MyOpenStreamOnFile(MAPIAllocateBuffer, MAPIFreeBuffer, STGM_READ, szMessageFile, &lpStream));

		if (lpStream)
		{
			LPITNEF lpTNEF = nullptr;
			hRes = EC_H2(OpenTnefStreamEx(
				nullptr,
				lpStream,
				reinterpret_cast<LPTSTR>(
					"winmail.dat"), // STRING_OK - despite its signature, this function is ANSI only
				TNEF_DECODE,
				lpMessage,
				static_cast<WORD>(GetTickCount() + 1),
				lpAdrBook,
				&lpTNEF));

			if (lpTNEF)
			{
				auto propTnefExcludeArray = SPropTagArray({1, {PR_URL_COMP_NAME}});
				LPSTnefProblemArray lpError = nullptr;
				// Decode the TNEF stream into our MAPI message.
				hRes = EC_MAPI2(lpTNEF->ExtractProps(TNEF_PROP_EXCLUDE, &propTnefExcludeArray, &lpError));

				EC_TNEFERR(lpError);
				MAPIFreeBuffer(lpError);
				lpTNEF->Release();
			}

			if (SUCCEEDED(hRes))
			{
				hRes = EC_MAPI2(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
			}

			lpStream->Release();
		}

		return hRes;
	}

	// Builds a file name out of the passed in message and extension
	std::wstring BuildFileName(_In_ const std::wstring& ext, _In_ const std::wstring& dir, _In_ LPMESSAGE lpMessage)
	{
		if (!lpMessage) return strings::emptystring;

		enum
		{
			ePR_SUBJECT_W,
			ePR_RECORD_KEY,
			NUM_COLS
		};
		static const SizedSPropTagArray(NUM_COLS, sptaMessageProps) = {NUM_COLS, {PR_SUBJECT_W, PR_RECORD_KEY}};

		// Get subject line of message
		// This will be used as the new file name.
		ULONG ulProps = NULL;
		LPSPropValue lpProps = nullptr;
		WC_H_GETPROPS_S(lpMessage->GetProps(LPSPropTagArray(&sptaMessageProps), fMapiUnicode, &ulProps, &lpProps));

		std::wstring szFileOut;
		if (lpProps)
		{
			std::wstring subj;
			if (mapi::CheckStringProp(&lpProps[ePR_SUBJECT_W], PT_UNICODE))
			{
				subj = lpProps[ePR_SUBJECT_W].Value.lpszW;
			}

			LPSBinary lpRecordKey = nullptr;
			if (PR_RECORD_KEY == lpProps[ePR_RECORD_KEY].ulPropTag)
			{
				lpRecordKey = &lpProps[ePR_RECORD_KEY].Value.bin;
			}

			szFileOut = BuildFileNameAndPath(ext, subj, dir, lpRecordKey);
		}

		MAPIFreeBuffer(lpProps);
		return szFileOut;
	}

	// The file name we generate should be shorter than MAX_PATH
	// This includes our directory name too!
	std::wstring BuildFileNameAndPath(
		_In_ const std::wstring& szExt,
		_In_ const std::wstring& szSubj,
		_In_ const std::wstring& szRootPath,
		_In_opt_ const _SBinary* lpBin)
	{
		output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath ext = \"%ws\"\n", szExt.c_str());
		output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath subj = \"%ws\"\n", szSubj.c_str());
		output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath rootPath = \"%ws\"\n", szRootPath.c_str());

		// Set up the path portion of the output.
		auto cleanRoot = ShortenPath(szRootPath);
		output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath cleanRoot = \"%ws\"\n", cleanRoot.c_str());

		// If we don't have enough space for even the shortest filename, give up.
		if (cleanRoot.length() >= MAXMSGPATH) return strings::emptystring;

		// Work with a max path which allows us to add our extension.
		// Shrink the max path to allow for a -ATTACHxxx extension.
		const auto maxFile = MAX_PATH - cleanRoot.length() - szExt.length() - MAXATTACH;

		std::wstring cleanSubj;
		if (!szSubj.empty())
		{
			cleanSubj = strings::SanitizeFileName(szSubj);
		}
		else
		{
			// We must have failed to get a subject before. Make one up.
			cleanSubj = L"UnknownSubject"; // STRING_OK
		}

		output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath cleanSubj = \"%ws\"\n", cleanSubj.c_str());

		std::wstring szBin;
		if (lpBin && lpBin->cb)
		{
			szBin = L"_" + strings::BinToHexString(lpBin, false);
		}

		if (cleanSubj.length() + szBin.length() <= maxFile)
		{
			auto szFile = cleanRoot + cleanSubj + szBin + szExt;
			output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath fileOut= \"%ws\"\n", szFile.c_str());
			return szFile;
		}

		// We couldn't build the string we wanted, so try something shorter
		// Compute a shorter subject length that should fit.
		const auto maxSubj = maxFile - min(MAXBIN, szBin.length()) - 1;
		auto szFile = cleanSubj.substr(0, maxSubj) + szBin.substr(0, MAXBIN);
		output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath shorter file = \"%ws\"\n", szFile.c_str());
		output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath new length = %d\n", szFile.length());

		if (szFile.length() >= maxFile)
		{
			szFile = cleanSubj.substr(0, MAXSUBJTIGHT) + szBin.substr(0, MAXBIN);
			output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath shorter file = \"%ws\"\n", szFile.c_str());
			output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath new length = %d\n", szFile.length());
		}

		if (szFile.length() >= maxFile)
		{
			output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath failed to build a string\n");
			return strings::emptystring;
		}

		auto szOut = cleanRoot + szFile + szExt;
		output::DebugPrint(DBGGeneric, L"BuildFileNameAndPath fileOut= \"%ws\"\n", szOut.c_str());
		return szOut;
	}

	void SaveFolderContentsToTXT(
		_In_ LPMDB lpMDB,
		_In_ LPMAPIFOLDER lpFolder,
		bool bRegular,
		bool bAssoc,
		bool bDescend,
		HWND hWnd)
	{
		auto szDir = GetDirectoryPath(hWnd);

		if (!szDir.empty())
		{
			CWaitCursor Wait; // Change the mouse to an hourglass while we work.
			mapiprocessor::CDumpStore MyDumpStore;
			MyDumpStore.InitMDB(lpMDB);
			MyDumpStore.InitFolder(lpFolder);
			MyDumpStore.InitFolderPathRoot(szDir);
			MyDumpStore.ProcessFolders(bRegular, bAssoc, bDescend);
		}
	}

	_Check_return_ HRESULT SaveFolderContentsToMSG(
		_In_ LPMAPIFOLDER lpFolder,
		_In_ const std::wstring& szPathName,
		bool bAssoc,
		bool bUnicode,
		HWND hWnd)
	{
		enum
		{
			fldPR_ENTRYID,
			fldPR_SUBJECT_W,
			fldPR_RECORD_KEY,
			fldNUM_COLS
		};

		static const SizedSPropTagArray(fldNUM_COLS, fldCols) = {fldNUM_COLS,
																 {PR_ENTRYID, PR_SUBJECT_W, PR_RECORD_KEY}};

		if (!lpFolder || szPathName.empty()) return MAPI_E_INVALID_PARAMETER;
		if (szPathName.length() >= MAXMSGPATH) return MAPI_E_INVALID_PARAMETER;

		output::DebugPrint(
			DBGGeneric, L"SaveFolderContentsToMSG: Saving contents of folder to \"%ws\"\n", szPathName.c_str());

		LPMAPITABLE lpFolderContents = nullptr;
		auto hRes =
			EC_MAPI2(lpFolder->GetContentsTable(fMapiUnicode | (bAssoc ? MAPI_ASSOCIATED : NULL), &lpFolderContents));
		if (lpFolderContents)
		{
			hRes = EC_MAPI2(lpFolderContents->SetColumns(LPSPropTagArray(&fldCols), TBL_BATCH));

			LPSRowSet pRows = nullptr;
			if (!FAILED(hRes))
			{
				for (;;)
				{
					if (pRows) FreeProws(pRows);
					pRows = nullptr;
					hRes = EC_MAPI2(lpFolderContents->QueryRows(1, NULL, &pRows));
					if (FAILED(hRes) || !pRows || pRows && !pRows->cRows) break;

					hRes = WC_H2(SaveToMSG(
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
			lpFolderContents->Release();
		}

		return hRes;
	}

#ifndef MRMAPI
	void ExportMessages(_In_ LPMAPIFOLDER lpFolder, HWND hWnd)
	{
		dialog::editor::CEditor MyData(
			nullptr, IDS_EXPORTTITLE, IDS_EXPORTPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyData.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_EXPORTSEARCHTERM, false));

		auto hRes = WC_H2(MyData.DisplayDialog());
		if (S_OK != hRes) return;

		auto szDir = GetDirectoryPath(hWnd);
		if (szDir.empty()) return;
		auto restrictString = MyData.GetStringW(0);
		if (restrictString.empty()) return;

		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		LPMAPITABLE lpTable = nullptr;
		WC_MAPI(lpFolder->GetContentsTable(MAPI_DEFERRED_ERRORS | MAPI_UNICODE, &lpTable));
		if (lpTable)
		{
			LPSRestriction lpRes = nullptr;
			// Allocate and create our SRestriction
			hRes = EC_H2(mapi::CreatePropertyStringRestriction(
				PR_SUBJECT_W, restrictString, FL_SUBSTRING | FL_IGNORECASE, nullptr, &lpRes));
			if (SUCCEEDED(hRes))
			{
				hRes = WC_MAPI2(lpTable->Restrict(lpRes, 0));
			}

			if (lpRes) MAPIFreeBuffer(lpRes);

			if (SUCCEEDED(hRes))
			{
				enum
				{
					fldPR_ENTRYID,
					fldPR_SUBJECT_W,
					fldPR_RECORD_KEY,
					fldNUM_COLS
				};

				static const SizedSPropTagArray(fldNUM_COLS, fldCols) = {fldNUM_COLS,
																		 {PR_ENTRYID, PR_SUBJECT_W, PR_RECORD_KEY}};

				hRes = WC_MAPI2(lpTable->SetColumns(LPSPropTagArray(&fldCols), TBL_ASYNC));

				// Export messages in the rows
				LPSRowSet lpRows = nullptr;
				if (!FAILED(hRes))
				{
					for (;;)
					{
						if (lpRows) FreeProws(lpRows);
						lpRows = nullptr;
						hRes = WC_MAPI2(lpTable->QueryRows(50, NULL, &lpRows));
						if (FAILED(hRes) || !lpRows || !lpRows->cRows) break;

						for (ULONG i = 0; i < lpRows->cRows; i++)
						{
							hRes = WC_H2(SaveToMSG(
								lpFolder,
								szDir,
								lpRows->aRow[i].lpProps[fldPR_ENTRYID],
								&lpRows->aRow[i].lpProps[fldPR_RECORD_KEY],
								&lpRows->aRow[i].lpProps[fldPR_SUBJECT_W],
								true,
								hWnd));
						}
					}

					if (lpRows) FreeProws(lpRows);
				}
			}

			lpTable->Release();
		}
	}
#endif

	_Check_return_ HRESULT WriteStreamToFile(_In_ LPSTREAM pStrmSrc, _In_ const std::wstring& szFileName)
	{
		if (!pStrmSrc || szFileName.empty()) return MAPI_E_INVALID_PARAMETER;

		LPSTREAM pStrmDest = nullptr;

		// Open an IStream interface and create the file at the
		// same time. This code will create the file in the
		// current directory.
		auto hRes = EC_H2(mapi::MyOpenStreamOnFile(
			MAPIAllocateBuffer, MAPIFreeBuffer, STGM_CREATE | STGM_READWRITE, szFileName, &pStrmDest));

		if (pStrmDest)
		{
			STATSTG StatInfo = {};
			pStrmSrc->Stat(&StatInfo, STATFLAG_NONAME);

			output::DebugPrint(DBGStream, L"WriteStreamToFile: Writing cb = %llu bytes\n", StatInfo.cbSize.QuadPart);

			hRes = EC_MAPI2(pStrmSrc->CopyTo(pStrmDest, StatInfo.cbSize, nullptr, nullptr));

			if (SUCCEEDED(hRes))
			{
				// Commit changes to new stream
				hRes = EC_MAPI2(pStrmDest->Commit(STGC_DEFAULT));
			}

			pStrmDest->Release();
		}

		return hRes;
	}

	_Check_return_ HRESULT SaveToEML(_In_ LPMESSAGE lpMessage, _In_ const std::wstring& szFileName)
	{
		LPSTREAM pStrmSrc = nullptr;

		if (!lpMessage || szFileName.empty()) return MAPI_E_INVALID_PARAMETER;
		output::DebugPrint(DBGGeneric, L"SaveToEML: Saving message to \"%ws\"\n", szFileName.c_str());

		// Open the property of the attachment
		// containing the file data
		auto hRes = EC_MAPI2(lpMessage->OpenProperty(
			PR_INTERNET_CONTENT, // TODO: There's a modern property for this...
			const_cast<LPIID>(&IID_IStream),
			0,
			NULL, // MAPI_MODIFY is not needed
			reinterpret_cast<LPUNKNOWN*>(&pStrmSrc)));
		if (FAILED(hRes))
		{
			if (MAPI_E_NOT_FOUND == hRes)
			{
				output::DebugPrint(DBGGeneric, L"No internet content found\n");
			}
		}
		else
		{
			if (pStrmSrc)
			{
				hRes = WC_H2(WriteStreamToFile(pStrmSrc, szFileName));

				pStrmSrc->Release();
			}
		}

		return hRes;
	}

	_Check_return_ HRESULT STDAPICALLTYPE MyStgCreateStorageEx(
		_In_ const std::wstring& pName,
		DWORD grfMode,
		DWORD stgfmt,
		DWORD grfAttrs,
		_In_ STGOPTIONS* pStgOptions,
		_Pre_null_ void* reserved,
		_In_ REFIID riid,
		_Out_ void** ppObjectOpen)
	{
		if (pName.empty()) return MAPI_E_INVALID_PARAMETER;

		if (import::pfnStgCreateStorageEx)
		{
			return import::pfnStgCreateStorageEx(
				pName.c_str(), grfMode, stgfmt, grfAttrs, pStgOptions, reserved, riid, ppObjectOpen);
		}

		// Fallback for NT4, which doesn't have StgCreateStorageEx
		return StgCreateDocfile(pName.c_str(), grfMode, 0, reinterpret_cast<LPSTORAGE*>(ppObjectOpen));
	}

	_Check_return_ HRESULT CreateNewMSG(
		_In_ const std::wstring& szFileName,
		bool bUnicode,
		_Deref_out_opt_ LPMESSAGE* lppMessage,
		_Deref_out_opt_ LPSTORAGE* lppStorage)
	{
		if (szFileName.empty() || !lppMessage || !lppStorage) return MAPI_E_INVALID_PARAMETER;

		auto hRes = S_OK;

		*lppMessage = nullptr;
		*lppStorage = nullptr;

		// get memory allocation function
		const auto pMalloc = MAPIGetDefaultMalloc();
		if (pMalloc)
		{
			LPSTORAGE pStorage = nullptr;

			STGOPTIONS myOpts = {0};
			myOpts.usVersion = 1; // STGOPTIONS_VERSION
			myOpts.ulSectorSize = 4096;

			// Open the compound file
			hRes = EC_H2(MyStgCreateStorageEx(
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
				LPMESSAGE pIMsg = nullptr;
				// Open an IMessage interface on an IStorage object
				hRes = EC_MAPI2(OpenIMsgOnIStg(
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
					hRes = EC_MAPI2(WriteClassStg(pStorage, guid::CLSID_MailMessage));
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
		_In_ LPMAPIFOLDER lpFolder,
		_In_ const std::wstring& szPathName,
		_In_ const SPropValue& entryID,
		_In_ const _SPropValue* lpRecordKey,
		_In_ const _SPropValue* lpSubject,
		bool bUnicode,
		HWND hWnd)
	{
		if (szPathName.empty() || szPathName.length() >= MAXMSGPATH) return MAPI_E_INVALID_PARAMETER;
		if (entryID.ulPropTag != PR_ENTRYID) return MAPI_E_INVALID_PARAMETER;

		auto hRes = S_OK;
		LPMESSAGE lpMessage = nullptr;

		output::DebugPrint(DBGGeneric, L"SaveToMSG: Saving message to \"%ws\"\n", szPathName.c_str());

		output::DebugPrint(DBGGeneric, L"Source Message =\n");
		output::DebugPrintBinary(DBGGeneric, entryID.Value.bin);

		auto lpMapiContainer = mapi::safe_cast<LPMAPICONTAINER>(lpFolder);
		if (lpMapiContainer)
		{
			hRes = EC_H2(mapi::CallOpenEntry(
				nullptr,
				nullptr,
				lpMapiContainer,
				nullptr,
				&entryID.Value.bin,
				nullptr,
				MAPI_BEST_ACCESS,
				nullptr,
				reinterpret_cast<LPUNKNOWN*>(&lpMessage)));
		}

		if (SUCCEEDED(hRes) && lpMessage != nullptr)
		{
			const auto szSubj =
				mapi::CheckStringProp(lpSubject, PT_UNICODE) ? lpSubject->Value.lpszW : L"UnknownSubject";
			const auto recordKey =
				lpRecordKey && lpRecordKey->ulPropTag == PR_RECORD_KEY ? &lpRecordKey->Value.bin : nullptr;

			auto szFileName = BuildFileNameAndPath(L".msg", szSubj, szPathName, recordKey); // STRING_OK
			if (!szFileName.empty())
			{
				output::DebugPrint(DBGGeneric, L"Saving to = \"%ws\"\n", szFileName.c_str());

				hRes = EC_H2(SaveToMSG(lpMessage, szFileName, bUnicode, hWnd, false));

				output::DebugPrint(DBGGeneric, L"Message Saved\n");
			}
		}

		if (lpMapiContainer) lpMapiContainer->Release();
		if (lpMessage) lpMessage->Release();

		return hRes;
	}

	_Check_return_ HRESULT
	SaveToMSG(_In_ LPMESSAGE lpMessage, _In_ const std::wstring& szFileName, bool bUnicode, HWND hWnd, bool bAllowUI)
	{
		if (!lpMessage || szFileName.empty()) return MAPI_E_INVALID_PARAMETER;

		output::DebugPrint(DBGGeneric, L"SaveToMSG: Saving message to \"%ws\"\n", szFileName.c_str());

		LPSTORAGE pStorage = nullptr;
		LPMESSAGE pIMsg = nullptr;
		auto hRes = EC_H2(CreateNewMSG(szFileName, bUnicode, &pIMsg, &pStorage));
		if (pIMsg && pStorage)
		{
			// Specify properties to exclude in the copy operation. These are
			// the properties that Exchange excludes to save bits and time.
			// Should not be necessary to exclude these, but speeds the process
			// when a lot of messages are being copied.
			static const SizedSPropTagArray(7, excludeTags) = {7,
															   {PR_ACCESS,
																PR_BODY,
																PR_RTF_SYNC_BODY_COUNT,
																PR_RTF_SYNC_BODY_CRC,
																PR_RTF_SYNC_BODY_TAG,
																PR_RTF_SYNC_PREFIX_COUNT,
																PR_RTF_SYNC_TRAILING_COUNT}};

			hRes = EC_H2(
				mapi::CopyTo(hWnd, lpMessage, pIMsg, &IID_IMessage, LPSPropTagArray(&excludeTags), false, bAllowUI));

			if (SUCCEEDED(hRes))
			{
				// save changes to IMessage object.
				hRes = EC_MAPI2(pIMsg->SaveChanges(KEEP_OPEN_READWRITE));
			}

			if (SUCCEEDED(hRes))
			{
				// save changes in storage of new doc file
				hRes = EC_MAPI2(pStorage->Commit(STGC_DEFAULT));
			}
		}

		if (pStorage) pStorage->Release();
		if (pIMsg) pIMsg->Release();

		return hRes;
	}

	_Check_return_ HRESULT
	SaveToTNEF(_In_ LPMESSAGE lpMessage, _In_ LPADRBOOK lpAdrBook, _In_ const std::wstring& szFileName)
	{
		enum
		{
			ulNumTNEFIncludeProps = 2
		};
		static const SizedSPropTagArray(ulNumTNEFIncludeProps, lpPropTnefIncludeArray) = {
			ulNumTNEFIncludeProps, {PR_MESSAGE_RECIPIENTS, PR_ATTACH_DATA_BIN}};

		enum
		{
			ulNumTNEFExcludeProps = 1
		};
		static const SizedSPropTagArray(ulNumTNEFExcludeProps, lpPropTnefExcludeArray) = {ulNumTNEFExcludeProps,
																						  {PR_URL_COMP_NAME}};

		if (!lpMessage || !lpAdrBook || szFileName.empty()) return MAPI_E_INVALID_PARAMETER;
		output::DebugPrint(DBGGeneric, L"SaveToTNEF: Saving message to \"%ws\"\n", szFileName.c_str());

		LPSTREAM lpStream = nullptr;
		LPITNEF lpTNEF = nullptr;

		static auto dwKey = static_cast<WORD>(GetTickCount());

		// Get a Stream interface on the input TNEF file
		auto hRes = EC_H2(mapi::MyOpenStreamOnFile(
			MAPIAllocateBuffer, MAPIFreeBuffer, STGM_READWRITE | STGM_CREATE, szFileName, &lpStream));

		if (lpStream)
		{
		// Open TNEF stream
#pragma warning(push)
#pragma warning(disable : 4616)
#pragma warning(disable : 6276)
			hRes = EC_H2(OpenTnefStreamEx(
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
				hRes =
					EC_MAPI2(lpTNEF->AddProps(TNEF_PROP_EXCLUDE, 0, nullptr, LPSPropTagArray(&lpPropTnefExcludeArray)));

				if (SUCCEEDED(hRes))
				{
					hRes = EC_MAPI2(lpTNEF->AddProps(
						TNEF_PROP_EXCLUDE | TNEF_PROP_ATTACHMENTS_ONLY,
						0,
						nullptr,
						LPSPropTagArray(&lpPropTnefExcludeArray)));
				}

				if (SUCCEEDED(hRes))
				{
					hRes = EC_MAPI2(
						lpTNEF->AddProps(TNEF_PROP_INCLUDE, 0, nullptr, LPSPropTagArray(&lpPropTnefIncludeArray)));
				}

				if (SUCCEEDED(hRes))
				{
					LPSTnefProblemArray lpError = nullptr;
					hRes = EC_MAPI2(lpTNEF->Finish(0, &dwKey, &lpError));
					EC_TNEFERR(lpError);
					MAPIFreeBuffer(lpError);
				}

				if (SUCCEEDED(hRes))
				{
					// Saving stream
					hRes = EC_MAPI2(lpStream->Commit(STGC_DEFAULT));
				}

				lpTNEF->Release();
			}

			lpStream->Release();
		}

		return hRes;
	}

	_Check_return_ HRESULT DeleteAttachments(_In_ LPMESSAGE lpMessage, _In_ const std::wstring& szAttName, HWND hWnd)
	{
		if (!lpMessage) return MAPI_E_INVALID_PARAMETER;

		enum
		{
			ATTACHNUM,
			ATTACHNAME,
			NUM_COLS
		};
		static const SizedSPropTagArray(NUM_COLS, sptAttachTableCols) = {NUM_COLS,
																		 {PR_ATTACH_NUM, PR_ATTACH_LONG_FILENAME_W}};

		LPSPropValue pProps = nullptr;
		auto hRes = EC_MAPI2(HrGetOneProp(lpMessage, PR_HASATTACH, &pProps));

		if (pProps && PR_HASATTACH == pProps[0].ulPropTag && pProps[0].Value.b)
		{
			LPMAPITABLE lpAttTbl = nullptr;
			hRes = EC_MAPI2(lpMessage->OpenProperty(
				PR_MESSAGE_ATTACHMENTS, &IID_IMAPITable, fMapiUnicode, 0, reinterpret_cast<LPUNKNOWN*>(&lpAttTbl)));
			if (lpAttTbl)
			{
				LPSRowSet pRows = nullptr;
				// I would love to just pass a restriction
				// to HrQueryAllRows for the file name. However,
				// we don't support restricting attachment tables (EMSMDB32 that is)
				// So I have to compare the strings myself (see below)
				hRes = EC_MAPI2(
					HrQueryAllRows(lpAttTbl, LPSPropTagArray(&sptAttachTableCols), nullptr, nullptr, 0, &pRows));

				auto bDirty = false;
				if (pRows)
				{
					for (ULONG iRow = 0; iRow < pRows->cRows; iRow++)
					{
						if (PR_ATTACH_NUM != pRows->aRow[iRow].lpProps[ATTACHNUM].ulPropTag) continue;

						if (!szAttName.empty())
						{
							if (PR_ATTACH_LONG_FILENAME_W != pRows->aRow[iRow].lpProps[ATTACHNAME].ulPropTag ||
								szAttName != pRows->aRow[iRow].lpProps[ATTACHNAME].Value.lpszW)
								continue;
						}

						// Open the attachment
						LPMAPIPROGRESS lpProgress =
							mapi::mapiui::GetMAPIProgress(L"IMessage::DeleteAttach", hWnd); // STRING_OK

						hRes = EC_MAPI2(lpMessage->DeleteAttach(
							pRows->aRow[iRow].lpProps[ATTACHNUM].Value.l,
							lpProgress ? reinterpret_cast<ULONG_PTR>(hWnd) : NULL,
							lpProgress,
							lpProgress ? ATTACH_DIALOG : 0));

						if (SUCCEEDED(hRes)) bDirty = true;

						if (lpProgress) lpProgress->Release();
					}

					FreeProws(pRows);
				}

				// Only call this if we got a successful DeleteAttach call
				if (bDirty)
				{
					hRes = EC_MAPI2(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
				}

				lpAttTbl->Release();
			}
		}

		MAPIFreeBuffer(pProps);

		return hRes;
	}

#ifndef MRMAPI
	_Check_return_ HRESULT WriteAttachmentsToFile(_In_ LPMESSAGE lpMessage, HWND hWnd)
	{
		if (!lpMessage) return MAPI_E_INVALID_PARAMETER;

		enum
		{
			ATTACHNUM,
			NUM_COLS
		};
		static const SizedSPropTagArray(NUM_COLS, sptAttachTableCols) = {NUM_COLS, {PR_ATTACH_NUM}};

		LPSPropValue pProps = nullptr;
		auto hRes = EC_MAPI2(HrGetOneProp(lpMessage, PR_HASATTACH, &pProps));

		if (pProps && PR_HASATTACH == pProps[0].ulPropTag && pProps[0].Value.b)
		{
			LPMAPITABLE lpAttTbl = nullptr;
			hRes = EC_MAPI2(lpMessage->OpenProperty(
				PR_MESSAGE_ATTACHMENTS, &IID_IMAPITable, fMapiUnicode, 0, reinterpret_cast<LPUNKNOWN*>(&lpAttTbl)));

			if (lpAttTbl)
			{
				LPSRowSet pRows = nullptr;
				hRes = EC_MAPI2(
					HrQueryAllRows(lpAttTbl, LPSPropTagArray(&sptAttachTableCols), nullptr, nullptr, 0, &pRows));

				if (pRows)
				{
					if (!FAILED(hRes))
					{
						for (ULONG iRow = 0; iRow < pRows->cRows; iRow++)
						{
							if (PR_ATTACH_NUM != pRows->aRow[iRow].lpProps[ATTACHNUM].ulPropTag) continue;

							LPATTACH lpAttach = nullptr;
							// Open the attachment
							hRes = EC_MAPI2(lpMessage->OpenAttach(
								pRows->aRow[iRow].lpProps[ATTACHNUM].Value.l, nullptr, MAPI_BEST_ACCESS, &lpAttach));

							if (lpAttach)
							{
								hRes = WC_H2(WriteAttachmentToFile(lpAttach, hWnd));
								lpAttach->Release();
								if (S_OK != hRes && iRow != pRows->cRows - 1)
								{
									if (dialog::bShouldCancel(nullptr, hRes)) break;
									hRes = S_OK;
								}
							}
						}
					}

					FreeProws(pRows);
				}

				lpAttTbl->Release();
			}
		}

		MAPIFreeBuffer(pProps);
		return hRes;
	}
#endif

	_Check_return_ HRESULT
	WriteEmbeddedMSGToFile(_In_ LPATTACH lpAttach, _In_ const std::wstring& szFileName, bool bUnicode, HWND hWnd)
	{
		if (!lpAttach || szFileName.empty()) return MAPI_E_INVALID_PARAMETER;

		output::DebugPrint(DBGGeneric, L"WriteEmbeddedMSGToFile: Saving attachment to \"%ws\"\n", szFileName.c_str());

		LPMESSAGE lpAttachMsg = nullptr;
		auto hRes = EC_MAPI2(lpAttach->OpenProperty(
			PR_ATTACH_DATA_OBJ,
			const_cast<LPIID>(&IID_IMessage),
			0,
			NULL, // MAPI_MODIFY is not needed
			reinterpret_cast<LPUNKNOWN*>(&lpAttachMsg)));

		if (lpAttachMsg)
		{
			hRes = EC_H2(SaveToMSG(lpAttachMsg, szFileName, bUnicode, hWnd, false));
			lpAttachMsg->Release();
		}

		return hRes;
	}

	_Check_return_ HRESULT WriteAttachStreamToFile(_In_ LPATTACH lpAttach, _In_ const std::wstring& szFileName)
	{
		if (!lpAttach || szFileName.empty()) return MAPI_E_INVALID_PARAMETER;

		LPSTREAM pStrmSrc = nullptr;

		// Open the property of the attachment
		// containing the file data
		auto hRes = WC_MAPI2(lpAttach->OpenProperty(
			PR_ATTACH_DATA_BIN,
			const_cast<LPIID>(&IID_IStream),
			0,
			NULL, // MAPI_MODIFY is not needed
			reinterpret_cast<LPUNKNOWN*>(&pStrmSrc)));
		if (FAILED(hRes))
		{
			if (MAPI_E_NOT_FOUND == hRes)
			{
				output::DebugPrint(DBGGeneric, L"No attachments found. Maybe the attachment was a message?\n");
			}
			else
				CHECKHRES(hRes);
		}
		else
		{
			if (pStrmSrc)
			{
				hRes = WC_H2(WriteStreamToFile(pStrmSrc, szFileName));

				pStrmSrc->Release();
			}
		}

		return hRes;
	}

	// Pretty sure this covers all OLE attachments - we don't need to look at PR_ATTACH_TAG
	_Check_return_ HRESULT WriteOleToFile(_In_ LPATTACH lpAttach, _In_ const std::wstring& szFileName)
	{
		LPSTREAM pStrmSrc = nullptr;

		// Open the property of the attachment containing the OLE data
		// Try to get it as an IStreamDocFile file first as that will be faster
		auto hRes = WC_MAPI2(lpAttach->OpenProperty(
			PR_ATTACH_DATA_OBJ,
			const_cast<LPIID>(&IID_IStreamDocfile),
			0,
			NULL,
			reinterpret_cast<LPUNKNOWN*>(&pStrmSrc)));

		// We got IStreamDocFile! Great! We can copy stream to stream into the file
		if (pStrmSrc)
		{
			hRes = WC_H2(WriteStreamToFile(pStrmSrc, szFileName));

			pStrmSrc->Release();
		}
		// We couldn't get IStreamDocFile! No problem - we'll try IStorage next
		else
		{
			LPSTORAGE lpStorageSrc = nullptr;
			hRes = EC_MAPI2(lpAttach->OpenProperty(
				PR_ATTACH_DATA_OBJ,
				const_cast<LPIID>(&IID_IStorage),
				0,
				NULL,
				reinterpret_cast<LPUNKNOWN*>(&lpStorageSrc)));

			if (lpStorageSrc)
			{
				LPSTORAGE lpStorageDest = nullptr;
				hRes = EC_H2(::StgCreateDocfile(
					szFileName.c_str(), STGM_READWRITE | STGM_TRANSACTED | STGM_CREATE, 0, &lpStorageDest));
				if (lpStorageDest)
				{
					hRes = EC_MAPI2(lpStorageSrc->CopyTo(NULL, nullptr, nullptr, lpStorageDest));

					if (SUCCEEDED(hRes))
					{
						hRes = EC_MAPI2(lpStorageDest->Commit(STGC_DEFAULT));
					}

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
		if (!lpAttach) return MAPI_E_INVALID_PARAMETER;

		enum
		{
			ATTACH_METHOD,
			ATTACH_LONG_FILENAME_W,
			ATTACH_FILENAME_W,
			DISPLAY_NAME_W,
			NUM_COLS
		};
		static const SizedSPropTagArray(NUM_COLS, sptaAttachProps) = {
			NUM_COLS, {PR_ATTACH_METHOD, PR_ATTACH_LONG_FILENAME_W, PR_ATTACH_FILENAME_W, PR_DISPLAY_NAME_W}};

		output::DebugPrint(DBGGeneric, L"WriteAttachmentToFile: Saving attachment.\n");

		ULONG ulProps = 0;
		LPSPropValue lpProps = nullptr;

		// Get required properties from the message
		auto hRes = EC_H_GETPROPS(lpAttach->GetProps(
			LPSPropTagArray(&sptaAttachProps), // property tag array
			fMapiUnicode, // flags
			&ulProps, // Count of values returned
			&lpProps)); // Values returned
		if (lpProps)
		{
			auto szName = L"Unknown"; // STRING_OK

			// Get a file name to use
			if (mapi::CheckStringProp(&lpProps[ATTACH_LONG_FILENAME_W], PT_UNICODE))
			{
				szName = lpProps[ATTACH_LONG_FILENAME_W].Value.lpszW;
			}
			else if (mapi::CheckStringProp(&lpProps[ATTACH_FILENAME_W], PT_UNICODE))
			{
				szName = lpProps[ATTACH_FILENAME_W].Value.lpszW;
			}
			else if (mapi::CheckStringProp(&lpProps[DISPLAY_NAME_W], PT_UNICODE))
			{
				szName = lpProps[DISPLAY_NAME_W].Value.lpszW;
			}

			auto szFileName = strings::SanitizeFileName(szName);

			// Get File Name
			switch (lpProps[ATTACH_METHOD].Value.l)
			{
			case ATTACH_BY_VALUE:
			case ATTACH_BY_REFERENCE:
			case ATTACH_BY_REF_RESOLVE:
			case ATTACH_BY_REF_ONLY:
			{
				output::DebugPrint(DBGGeneric, L"WriteAttachmentToFile: Prompting with \"%ws\"\n", szFileName.c_str());

				auto file = CFileDialogExW::SaveAs(
					L"txt", // STRING_OK
					szFileName,
					OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
					strings::loadstring(IDS_ALLFILES));
				if (!file.empty())
				{
					hRes = EC_H2(WriteAttachStreamToFile(lpAttach, file));
				}
			}
			break;
			case ATTACH_EMBEDDED_MSG:
				// Get File Name
				{
					output::DebugPrint(
						DBGGeneric, L"WriteAttachmentToFile: Prompting with \"%ws\"\n", szFileName.c_str());
					auto file = CFileDialogExW::SaveAs(
						L"msg", // STRING_OK
						szFileName,
						OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
						strings::loadstring(IDS_MSGFILES));
					if (!file.empty())
					{
						hRes = EC_H2(WriteEmbeddedMSGToFile(lpAttach, file, MAPI_UNICODE == fMapiUnicode, hWnd));
					}
				}
				break;
			case ATTACH_OLE:
			{
				output::DebugPrint(DBGGeneric, L"WriteAttachmentToFile: Prompting with \"%ws\"\n", szFileName.c_str());
				auto file = CFileDialogExW::SaveAs(
					strings::emptystring,
					szFileName,
					OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
					strings::loadstring(IDS_ALLFILES));
				if (!file.empty())
				{
					hRes = EC_H2(WriteOleToFile(lpAttach, file));
				}
			}
			break;
			default:
				error::ErrDialog(__FILE__, __LINE__, IDS_EDUNKNOWNATTMETHOD);
				break;
			}
		}

		MAPIFreeBuffer(lpProps);
		return hRes;
	}
#endif

	std::wstring GetModuleFileName(_In_opt_ HMODULE hModule)
	{
		auto buf = std::vector<wchar_t>();
		auto copied = DWORD();
		do
		{
			buf.resize(buf.size() + MAX_PATH);
			copied = EC_D(DWORD, ::GetModuleFileNameW(hModule, &buf.at(0), static_cast<DWORD>(buf.size())));
		} while (copied >= buf.size());

		buf.resize(copied);

		const auto path = std::wstring(buf.begin(), buf.end());

		return path;
	}

	std::wstring GetSystemDirectory()
	{
		auto buf = std::vector<wchar_t>();
		auto copied = DWORD();
		do
		{
			buf.resize(buf.size() + MAX_PATH);
			copied = EC_D(DWORD, ::GetSystemDirectoryW(&buf.at(0), static_cast<UINT>(buf.size())));
		} while (copied >= buf.size());

		buf.resize(copied);

		const auto path = std::wstring(buf.begin(), buf.end());

		return path;
	}

	std::map<std::wstring, std::wstring> GetFileVersionInfo(_In_opt_ HMODULE hModule)
	{
		auto versionStrings = std::map<std::wstring, std::wstring>();
		const auto szFullPath = file::GetModuleFileName(hModule);
		if (!szFullPath.empty())
		{
			auto dwVerInfoSize = EC_D(DWORD, GetFileVersionInfoSizeW(szFullPath.c_str(), nullptr));
			if (dwVerInfoSize)
			{
				// If we were able to get the information, process it.
				const auto pbData = new BYTE[dwVerInfoSize];
				if (pbData == nullptr) return {};

				auto hRes = EC_B(GetFileVersionInfoW(szFullPath.c_str(), NULL, dwVerInfoSize, static_cast<void*>(pbData)));

				if (SUCCEEDED(hRes))
				{
					struct LANGANDCODEPAGE
					{
						WORD wLanguage;
						WORD wCodePage;
					}* lpTranslate = {nullptr};

					UINT cbTranslate = 0;

					// Read the list of languages and code pages.
					hRes = EC_B(VerQueryValueW(
						pbData,
						L"\\VarFileInfo\\Translation", // STRING_OK
						reinterpret_cast<LPVOID*>(&lpTranslate),
						&cbTranslate));

					// Read the file description for each language and code page.

					if (hRes == S_OK && lpTranslate)
					{
						for (UINT iCodePages = 0; iCodePages < cbTranslate / sizeof(LANGANDCODEPAGE); iCodePages++)
						{
							const auto szSubBlock = strings::format(
								L"\\StringFileInfo\\%04x%04x\\", // STRING_OK
								lpTranslate[iCodePages].wLanguage,
								lpTranslate[iCodePages].wCodePage);

							// Load all our strings
							for (auto iVerString = IDS_VER_FIRST; iVerString <= IDS_VER_LAST; iVerString++)
							{
								UINT cchVer = 0;
								wchar_t* lpszVer = nullptr;
								auto szVerString = strings::loadstring(iVerString);
								auto szQueryString = szSubBlock + szVerString;

								hRes = EC_B(VerQueryValueW(
									static_cast<void*>(pbData),
									szQueryString.c_str(),
									reinterpret_cast<void**>(&lpszVer),
									&cchVer));

								if (hRes == S_OK && cchVer && lpszVer)
								{
									versionStrings[szVerString] = lpszVer;
								}
							}
						}
					}
				}

				delete[] pbData;
			}
		}

		return versionStrings;
	}
}