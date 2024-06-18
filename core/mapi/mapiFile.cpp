#include <core/stdafx.h>
#include <core/mapi/mapiFile.h>
#include <core/utility/file.h>
#include <core/utility/strings.h>
#include <core/mapi/mapiProgress.h>
#include <core/utility/import.h>
#include <core/utility/output.h>
#include <core/interpret/guid.h>
#include <core/utility/error.h>
#include <core/mapi/mapiFunctions.h>
#include <core/mapi/mapiOutput.h>

namespace file
{
	// Opens storage with best access
	_Check_return_ HRESULT
	MyStgOpenStorage(_In_ const std::wstring& szMessageFile, bool bBestAccess, _Deref_out_ LPSTORAGE* lppStorage)
	{
		if (!lppStorage) return MAPI_E_INVALID_PARAMETER;
		output::DebugPrint(
			output::dbgLevel::Generic,
			L"MyStgOpenStorage: Opening \"%ws\", bBestAccess == %ws\n",
			szMessageFile.c_str(),
			bBestAccess ? L"True" : L"False");
		ULONG ulFlags = STGM_TRANSACTED;

		if (bBestAccess) ulFlags |= STGM_READWRITE;

		auto hRes = WC_H(::StgOpenStorage(szMessageFile.c_str(), nullptr, ulFlags, nullptr, 0, lppStorage));

		// If we asked for best access (read/write) and didn't get it, then try it without readwrite
		if (hRes == STG_E_ACCESSDENIED && !*lppStorage && bBestAccess)
		{
			hRes = EC_H(MyStgOpenStorage(szMessageFile, false, lppStorage));
		}

		return hRes;
	}

	// Creates an LPMESSAGE on top of the MSG file
	_Check_return_ LPMESSAGE LoadMSGToMessage(_In_ const std::wstring& szMessageFile)
	{
		LPMESSAGE lpMessage = nullptr;
		// get memory allocation function
		const auto lpMalloc = MAPIGetDefaultMalloc();
		if (lpMalloc)
		{
			LPSTORAGE pStorage = nullptr;
			// Open the compound file
			EC_H_S(MyStgOpenStorage(szMessageFile, true, &pStorage));
			if (pStorage)
			{
				// Open an IMessage interface on an IStorage object
				EC_MAPI_S(OpenIMsgOnIStg(
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
					&lpMessage));

				pStorage->Release();
			}
		}

		return lpMessage;
	}

	// Loads the MSG file into an LPMESSAGE pointer, then copies it into the passed in message
	// lpMessage must be created first
	_Check_return_ HRESULT LoadFromMSG(_In_ const std::wstring& szMessageFile, _In_ LPMESSAGE lpMessage, HWND hWnd)
	{
		// Specify properties to exclude in the copy operation. These are
		// the properties that Exchange excludes to save bits and time.
		// Should not be necessary to exclude these, but speeds the process
		// when a lot of messages are being copied.
		static const SizedSPropTagArray(18, excludeTags) = {
			18,
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

		auto hRes = S_OK;
		auto pIMsg = LoadMSGToMessage(szMessageFile);
		if (pIMsg)
		{
			LPSPropProblemArray lpProblems = nullptr;

			auto lpProgress = mapi::mapiui::GetMAPIProgress(L"IMAPIProp::CopyTo", hWnd); // STRING_OK

			hRes = EC_MAPI(pIMsg->CopyTo(
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
			if (SUCCEEDED(hRes))
			{
				hRes = EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
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
			EC_H(file::MyOpenStreamOnFile(MAPIAllocateBuffer, MAPIFreeBuffer, STGM_READ, szMessageFile, &lpStream));

		if (lpStream)
		{
			LPITNEF lpTNEF = nullptr;
#pragma warning(push)
#pragma warning( \
	disable : 28159) // Warning C28159 Consider using *function_name_1* instead of *function_name_2*. Reason: reason
			//GetTickCount is OK for what we're doing
			static auto wKeyVal = static_cast<WORD>(GetTickCount() + 1);
#pragma warning(pop)

			hRes = EC_H(OpenTnefStreamEx(
				nullptr,
				lpStream,
				strings::LPCSTRToLPTSTR("winmail.dat"), // STRING_OK - despite its signature, this function is ANSI only
				TNEF_DECODE,
				lpMessage,
				wKeyVal,
				lpAdrBook,
				&lpTNEF));

			if (lpTNEF)
			{
				auto lpPropTnefExcludeArray = SPropTagArray({1, {PR_URL_COMP_NAME}});
				LPSTnefProblemArray lpError = nullptr;
				// Decode the TNEF stream into our MAPI message.
				hRes = EC_MAPI(
					lpTNEF->ExtractProps(TNEF_PROP_EXCLUDE, LPSPropTagArray(&lpPropTnefExcludeArray), &lpError));

				EC_TNEFERR(lpError);
				MAPIFreeBuffer(lpError);
				lpTNEF->Release();
			}

			if (SUCCEEDED(hRes))
			{
				hRes = EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
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
			if (strings::CheckStringProp(&lpProps[ePR_SUBJECT_W], PT_UNICODE))
			{
				subj = lpProps[ePR_SUBJECT_W].Value.lpszW;
			}

			SBinary recordKey = {};
			if (PR_RECORD_KEY == lpProps[ePR_RECORD_KEY].ulPropTag)
			{
				recordKey = mapi::getBin(lpProps[ePR_RECORD_KEY]);
			}

			szFileOut = BuildFileNameAndPath(ext, subj, dir, &recordKey);
		}

		MAPIFreeBuffer(lpProps);
		return szFileOut;
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

		static const SizedSPropTagArray(fldNUM_COLS, fldCols) = {
			fldNUM_COLS, {PR_ENTRYID, PR_SUBJECT_W, PR_RECORD_KEY}};

		if (!lpFolder || szPathName.empty()) return MAPI_E_INVALID_PARAMETER;
		if (szPathName.length() >= MAXMSGPATH) return MAPI_E_INVALID_PARAMETER;

		output::DebugPrint(
			output::dbgLevel::Generic,
			L"SaveFolderContentsToMSG: Saving contents of folder to \"%ws\"\n",
			szPathName.c_str());

		LPMAPITABLE lpFolderContents = nullptr;
		auto hRes =
			EC_MAPI(lpFolder->GetContentsTable(fMapiUnicode | (bAssoc ? MAPI_ASSOCIATED : NULL), &lpFolderContents));

		if (lpFolderContents)
		{
			hRes = EC_MAPI(lpFolderContents->SetColumns(LPSPropTagArray(&fldCols), TBL_BATCH));

			LPSRowSet pRows = nullptr;
			if (SUCCEEDED(hRes))
			{
				for (;;)
				{
					if (pRows) FreeProws(pRows);
					pRows = nullptr;
					hRes = EC_MAPI(lpFolderContents->QueryRows(1, NULL, &pRows));
					if (FAILED(hRes) || !pRows || pRows && !pRows->cRows) break;

					WC_H_S(SaveToMSG(
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

	_Check_return_ HRESULT WriteStreamToFile(_In_ LPSTREAM pStrmSrc, _In_ const std::wstring& szFileName)
	{
		if (!pStrmSrc || szFileName.empty()) return MAPI_E_INVALID_PARAMETER;

		LPSTREAM pStrmDest = nullptr;

		// Open an IStream interface and create the file at the
		// same time. This code will create the file in the
		// current directory.
		auto hRes = EC_H(file::MyOpenStreamOnFile(
			MAPIAllocateBuffer, MAPIFreeBuffer, STGM_CREATE | STGM_READWRITE, szFileName, &pStrmDest));

		if (pStrmDest)
		{
			STATSTG StatInfo = {};
			pStrmSrc->Stat(&StatInfo, STATFLAG_NONAME);

			output::DebugPrint(
				output::dbgLevel::Stream, L"WriteStreamToFile: Writing cb = %llu bytes\n", StatInfo.cbSize.QuadPart);

			hRes = EC_MAPI(pStrmSrc->CopyTo(pStrmDest, StatInfo.cbSize, nullptr, nullptr));

			if (SUCCEEDED(hRes))
			{
				// Commit changes to new stream
				hRes = EC_MAPI(pStrmDest->Commit(STGC_DEFAULT));
			}

			pStrmDest->Release();
		}

		return hRes;
	}

	_Check_return_ HRESULT SaveToEML(_In_ LPMESSAGE lpMessage, _In_ const std::wstring& szFileName)
	{
		LPSTREAM pStrmSrc = nullptr;

		if (!lpMessage || szFileName.empty()) return MAPI_E_INVALID_PARAMETER;
		output::DebugPrint(output::dbgLevel::Generic, L"SaveToEML: Saving message to \"%ws\"\n", szFileName.c_str());

		// Open the property of the attachment
		// containing the file data
		auto hRes = EC_MAPI(lpMessage->OpenProperty(
			PR_INTERNET_CONTENT, // TODO: There's a modern property for this...
			const_cast<LPIID>(&IID_IStream),
			0,
			NULL, // MAPI_MODIFY is not needed
			reinterpret_cast<LPUNKNOWN*>(&pStrmSrc)));
		if (FAILED(hRes))
		{
			if (hRes == MAPI_E_NOT_FOUND)
			{
				output::DebugPrint(output::dbgLevel::Generic, L"No internet content found\n");
			}
		}
		else
		{
			if (pStrmSrc)
			{
				hRes = WC_H(WriteStreamToFile(pStrmSrc, szFileName));

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
		_Out_ void** ppObjectOpen) noexcept
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

			STGOPTIONS myOpts = {};
			myOpts.usVersion = 1; // STGOPTIONS_VERSION
			myOpts.ulSectorSize = 4096;

			// Open the compound file
			hRes = EC_H(MyStgCreateStorageEx(
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
				hRes = EC_MAPI(OpenIMsgOnIStg(
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
					hRes = EC_MAPI(WriteClassStg(pStorage, guid::CLSID_MailMessage));
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

		LPMESSAGE lpMessage = nullptr;

		output::DebugPrint(output::dbgLevel::Generic, L"SaveToMSG: Saving message to \"%ws\"\n", szPathName.c_str());

		output::DebugPrint(output::dbgLevel::Generic, L"Source Message =\n");
		const auto bin = mapi::getBin(entryID);
		output::outputBinary(output::dbgLevel::Generic, nullptr, bin);

		auto lpMapiContainer = mapi::safe_cast<LPMAPICONTAINER>(lpFolder);
		if (lpMapiContainer)
		{
			lpMessage = mapi::CallOpenEntry<LPMESSAGE>(
				nullptr, nullptr, lpMapiContainer, nullptr, &bin, nullptr, MAPI_BEST_ACCESS, nullptr);
		}

		auto hRes = S_OK;
		if (lpMessage != nullptr)
		{
			const auto szSubj =
				strings::CheckStringProp(lpSubject, PT_UNICODE) ? lpSubject->Value.lpszW : L"UnknownSubject";
			const auto recordKey =
				lpRecordKey && lpRecordKey->ulPropTag == PR_RECORD_KEY ? &mapi::getBin(lpRecordKey) : nullptr;

			auto szFileName = BuildFileNameAndPath(L".msg", szSubj, szPathName, recordKey); // STRING_OK
			if (!szFileName.empty())
			{
				output::DebugPrint(output::dbgLevel::Generic, L"Saving to = \"%ws\"\n", szFileName.c_str());

				hRes = EC_H(SaveToMSG(lpMessage, szFileName, bUnicode, hWnd, false));

				output::DebugPrint(output::dbgLevel::Generic, L"Message Saved\n");
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

		output::DebugPrint(output::dbgLevel::Generic, L"SaveToMSG: Saving message to \"%ws\"\n", szFileName.c_str());

		LPSTORAGE pStorage = nullptr;
		LPMESSAGE pIMsg = nullptr;
		auto hRes = EC_H(CreateNewMSG(szFileName, bUnicode, &pIMsg, &pStorage));
		if (pIMsg && pStorage)
		{
			// Specify properties to exclude in the copy operation. These are
			// the properties that Exchange excludes to save bits and time.
			// Should not be necessary to exclude these, but speeds the process
			// when a lot of messages are being copied.
			static const SizedSPropTagArray(7, excludeTags) = {
				7,
				{PR_ACCESS,
				 PR_BODY,
				 PR_RTF_SYNC_BODY_COUNT,
				 PR_RTF_SYNC_BODY_CRC,
				 PR_RTF_SYNC_BODY_TAG,
				 PR_RTF_SYNC_PREFIX_COUNT,
				 PR_RTF_SYNC_TRAILING_COUNT}};

			hRes = EC_H(
				mapi::CopyTo(hWnd, lpMessage, pIMsg, &IID_IMessage, LPSPropTagArray(&excludeTags), false, bAllowUI));

			if (SUCCEEDED(hRes))
			{
				// save changes to IMessage object.
				hRes = EC_MAPI(pIMsg->SaveChanges(KEEP_OPEN_READWRITE));
			}

			if (SUCCEEDED(hRes))
			{
				// save changes in storage of new doc file
				hRes = EC_MAPI(pStorage->Commit(STGC_DEFAULT));
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
		static const SizedSPropTagArray(ulNumTNEFExcludeProps, lpPropTnefExcludeArray) = {
			ulNumTNEFExcludeProps, {PR_URL_COMP_NAME}};

		if (!lpMessage || !lpAdrBook || szFileName.empty()) return MAPI_E_INVALID_PARAMETER;
		output::DebugPrint(output::dbgLevel::Generic, L"SaveToTNEF: Saving message to \"%ws\"\n", szFileName.c_str());

		LPSTREAM lpStream = nullptr;
		LPITNEF lpTNEF = nullptr;

#pragma warning(push)
#pragma warning( \
	disable : 28159) // Warning C28159 Consider using *function_name_1* instead of *function_name_2*. Reason: reason
		//GetTickCount is OK for what we're doing
		static auto wKeyVal = static_cast<WORD>(GetTickCount() + 1);
#pragma warning(pop)

		// Get a Stream interface on the input TNEF file
		auto hRes = EC_H(file::MyOpenStreamOnFile(
			MAPIAllocateBuffer, MAPIFreeBuffer, STGM_READWRITE | STGM_CREATE, szFileName, &lpStream));

		if (lpStream)
		{
			// Open TNEF stream
			hRes = EC_H(OpenTnefStreamEx(
				nullptr,
				lpStream,
				strings::LPCSTRToLPTSTR("winmail.dat"), // STRING_OK - despite its signature, this function is ANSI only
				TNEF_ENCODE,
				lpMessage,
				wKeyVal,
				lpAdrBook,
				&lpTNEF));

			if (lpTNEF)
			{
				// Excludes
				hRes =
					EC_MAPI(lpTNEF->AddProps(TNEF_PROP_EXCLUDE, 0, nullptr, LPSPropTagArray(&lpPropTnefExcludeArray)));

				if (SUCCEEDED(hRes))
				{
					hRes = EC_MAPI(lpTNEF->AddProps(
						TNEF_PROP_EXCLUDE | TNEF_PROP_ATTACHMENTS_ONLY,
						0,
						nullptr,
						LPSPropTagArray(&lpPropTnefExcludeArray)));
				}

				if (SUCCEEDED(hRes))
				{
					hRes = EC_MAPI(
						lpTNEF->AddProps(TNEF_PROP_INCLUDE, 0, nullptr, LPSPropTagArray(&lpPropTnefIncludeArray)));
				}

				if (SUCCEEDED(hRes))
				{
					LPSTnefProblemArray lpError = nullptr;
					hRes = EC_MAPI(lpTNEF->Finish(0, &wKeyVal, &lpError));
					EC_TNEFERR(lpError);
					MAPIFreeBuffer(lpError);
				}

				if (SUCCEEDED(hRes))
				{
					// Saving stream
					hRes = EC_MAPI(lpStream->Commit(STGC_DEFAULT));
				}

				lpTNEF->Release();
			}

			lpStream->Release();
		}

		return hRes;
	}

	_Check_return_ HRESULT IterateAttachments(
		_In_ LPMESSAGE lpMessage,
		_In_ LPSPropTagArray lpSPropTagArray,
		const std::function<HRESULT(LPSPropValue)>& operation,
		const std::function<bool(HRESULT)>& shouldCancel)
	{
		if (!lpMessage) return MAPI_E_INVALID_PARAMETER;

		LPSPropValue pProps = nullptr;
		auto hRes = EC_MAPI(HrGetOneProp(lpMessage, PR_HASATTACH, &pProps));

		if (pProps && PR_HASATTACH == pProps[0].ulPropTag && pProps[0].Value.b)
		{
			LPMAPITABLE lpAttTbl = nullptr;
			hRes = EC_MAPI(lpMessage->OpenProperty(
				PR_MESSAGE_ATTACHMENTS, &IID_IMAPITable, fMapiUnicode, 0, reinterpret_cast<LPUNKNOWN*>(&lpAttTbl)));
			if (lpAttTbl)
			{
				LPSRowSet pRows = nullptr;
				hRes = EC_MAPI(HrQueryAllRows(lpAttTbl, lpSPropTagArray, nullptr, nullptr, 0, &pRows));

				if (pRows)
				{
					for (ULONG iRow = 0; iRow < pRows->cRows; iRow++)
					{
						hRes = operation(pRows->aRow[iRow].lpProps);
						if (S_OK != hRes && iRow != pRows->cRows - 1)
						{
							if (shouldCancel(hRes)) break;
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

	_Check_return_ HRESULT DeleteAttachments(_In_ LPMESSAGE lpMessage, _In_ const std::wstring& szAttName, HWND hWnd)
	{
		enum
		{
			ATTACHNUM,
			ATTACHNAME,
			NUM_COLS
		};
		static SizedSPropTagArray(NUM_COLS, sptAttachTableCols) = {
			NUM_COLS, {PR_ATTACH_NUM, PR_ATTACH_LONG_FILENAME_W}};

		auto bDirty = false;
		auto hRes = IterateAttachments(
			lpMessage,
			reinterpret_cast<LPSPropTagArray>(&sptAttachTableCols),
			[&](auto props) {
				if (PR_ATTACH_NUM != props[ATTACHNUM].ulPropTag) return S_OK;

				// If we were given an attachment name to look for, skip attachments without that name
				if (!szAttName.empty())
				{
					if (PR_ATTACH_LONG_FILENAME_W != props[ATTACHNAME].ulPropTag ||
						szAttName != props[ATTACHNAME].Value.lpszW)
						return S_OK;
				}

				// Delete the attachment
				auto lpProgress = mapi::mapiui::GetMAPIProgress(L"IMessage::DeleteAttach", hWnd); // STRING_OK

				const auto hResRet = EC_MAPI(lpMessage->DeleteAttach(
					props[ATTACHNUM].Value.l,
					lpProgress ? reinterpret_cast<ULONG_PTR>(hWnd) : NULL,
					lpProgress,
					lpProgress ? ATTACH_DIALOG : 0));

				if (SUCCEEDED(hResRet)) bDirty = true;

				if (lpProgress) lpProgress->Release();
				return hResRet;
			},
			[](auto) { return false; });

		// Only call this if we got a successful DeleteAttach call
		if (bDirty)
		{
			hRes = EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
		}

		return hRes;
	}

	_Check_return_ HRESULT
	WriteEmbeddedMSGToFile(_In_ LPATTACH lpAttach, _In_ const std::wstring& szFileName, bool bUnicode, HWND hWnd)
	{
		if (!lpAttach || szFileName.empty()) return MAPI_E_INVALID_PARAMETER;

		output::DebugPrint(
			output::dbgLevel::Generic, L"WriteEmbeddedMSGToFile: Saving attachment to \"%ws\"\n", szFileName.c_str());

		LPMESSAGE lpAttachMsg = nullptr;
		auto hRes = EC_MAPI(lpAttach->OpenProperty(
			PR_ATTACH_DATA_OBJ,
			const_cast<LPIID>(&IID_IMessage),
			0,
			NULL, // MAPI_MODIFY is not needed
			reinterpret_cast<LPUNKNOWN*>(&lpAttachMsg)));

		if (lpAttachMsg)
		{
			hRes = EC_H(SaveToMSG(lpAttachMsg, szFileName, bUnicode, hWnd, false));
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
		auto hRes = WC_MAPI(lpAttach->OpenProperty(
			PR_ATTACH_DATA_BIN,
			const_cast<LPIID>(&IID_IStream),
			0,
			NULL, // MAPI_MODIFY is not needed
			reinterpret_cast<LPUNKNOWN*>(&pStrmSrc)));
		if (FAILED(hRes))
		{
			if (hRes == MAPI_E_NOT_FOUND)
			{
				output::DebugPrint(
					output::dbgLevel::Generic, L"No attachments found. Maybe the attachment was a message?\n");
			}
			else
				CHECKHRES(hRes);
		}
		else
		{
			if (pStrmSrc)
			{
				hRes = WC_H(WriteStreamToFile(pStrmSrc, szFileName));

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
		auto hRes = WC_MAPI(lpAttach->OpenProperty(
			PR_ATTACH_DATA_OBJ,
			const_cast<LPIID>(&IID_IStreamDocfile),
			0,
			NULL,
			reinterpret_cast<LPUNKNOWN*>(&pStrmSrc)));

		// We got IStreamDocFile! Great! We can copy stream to stream into the file
		if (pStrmSrc)
		{
			hRes = WC_H(WriteStreamToFile(pStrmSrc, szFileName));

			pStrmSrc->Release();
		}
		// We couldn't get IStreamDocFile! No problem - we'll try IStorage next
		else
		{
			LPSTORAGE lpStorageSrc = nullptr;
			hRes = EC_MAPI(lpAttach->OpenProperty(
				PR_ATTACH_DATA_OBJ,
				const_cast<LPIID>(&IID_IStorage),
				0,
				NULL,
				reinterpret_cast<LPUNKNOWN*>(&lpStorageSrc)));

			if (lpStorageSrc)
			{
				LPSTORAGE lpStorageDest = nullptr;
				hRes = EC_H(::StgCreateDocfile(
					szFileName.c_str(), STGM_READWRITE | STGM_TRANSACTED | STGM_CREATE, 0, &lpStorageDest));
				if (lpStorageDest)
				{
					hRes = EC_MAPI(lpStorageSrc->CopyTo(NULL, nullptr, nullptr, lpStorageDest));

					if (SUCCEEDED(hRes))
					{
						hRes = EC_MAPI(lpStorageDest->Commit(STGC_DEFAULT));
					}

					lpStorageDest->Release();
				}

				lpStorageSrc->Release();
			}
		}

		return hRes;
	}

	// Declaration missing from MAPI headers
	_Check_return_ STDAPI OpenStreamOnFileW(
		_In_ LPALLOCATEBUFFER lpAllocateBuffer,
		_In_ LPFREEBUFFER lpFreeBuffer,
		ULONG ulFlags,
		_In_z_ LPCWSTR lpszFileName,
		_In_opt_z_ LPCWSTR lpszPrefix,
		_Out_ LPSTREAM FAR* lppStream);

	// Since I never use lpszPrefix, I don't convert it
	// To make certain of that, I pass NULL for it
	// If I ever do need this param, I'll have to fix this
	_Check_return_ STDMETHODIMP MyOpenStreamOnFile(
		_In_ LPALLOCATEBUFFER lpAllocateBuffer,
		_In_ LPFREEBUFFER lpFreeBuffer,
		ULONG ulFlags,
		_In_ const std::wstring& lpszFileName,
		_Out_ LPSTREAM FAR* lppStream)
	{
		auto hRes =
			OpenStreamOnFileW(lpAllocateBuffer, lpFreeBuffer, ulFlags, lpszFileName.c_str(), nullptr, lppStream);
		if (hRes == MAPI_E_CALL_FAILED)
		{
			hRes = OpenStreamOnFile(
				lpAllocateBuffer,
				lpFreeBuffer,
				ulFlags,
				strings::wstringTotstring(lpszFileName).c_str(),
				nullptr,
				lppStream);
		}

		return hRes;
	}
} // namespace file
