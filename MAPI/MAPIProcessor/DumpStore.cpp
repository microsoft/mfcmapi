
// Routines used in dumping the contents of the Exchange store
// in to a log file

#include <StdAfx.h>
#include <MAPI/MAPIProcessor/DumpStore.h>
#include <MAPI/MAPIFunctions.h>
#include <Interpret/String.h>
#include <IO/File.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>

namespace mapiprocessor
{
	CDumpStore::CDumpStore()
	{
		m_fFolderProps = nullptr;
		m_fFolderContents = nullptr;
		m_fMailboxTable = nullptr;

		m_bRetryStreamProps = true;
		m_bOutputAttachments = true;
		m_bOutputMSG = false;
		m_bOutputList = false;
	}

	CDumpStore::~CDumpStore()
	{
		if (m_fFolderProps) output::CloseFile(m_fFolderProps);
		if (m_fFolderContents) output::CloseFile(m_fFolderContents);
		if (m_fMailboxTable) output::CloseFile(m_fMailboxTable);
	}

	void CDumpStore::InitMessagePath(_In_ const std::wstring& szMessageFileName)
	{
		m_szMessageFileName = szMessageFileName;
	}

	void CDumpStore::InitFolderPathRoot(_In_ const std::wstring& szFolderPathRoot)
	{
		m_szFolderPathRoot = szFolderPathRoot;
		if (m_szFolderPathRoot.length() >= MAXMSGPATH)
		{
			output::DebugPrint(DBGGeneric, L"InitFolderPathRoot: \"%ws\" length (%d) greater than max length (%d)\n", m_szFolderPathRoot.c_str(), m_szFolderPathRoot.length(), MAXMSGPATH);
		}
	}

	void CDumpStore::InitMailboxTablePathRoot(_In_ const std::wstring& szMailboxTablePathRoot)
	{
		m_szMailboxTablePathRoot = szMailboxTablePathRoot;
	}

	void CDumpStore::EnableMSG()
	{
		m_bOutputMSG = true;
	}

	void CDumpStore::EnableList()
	{
		m_bOutputList = true;
	}

	void CDumpStore::DisableStreamRetry()
	{
		m_bRetryStreamProps = false;
	}

	void CDumpStore::DisableEmbeddedAttachments()
	{
		m_bOutputAttachments = false;
	}

	void CDumpStore::BeginMailboxTableWork(_In_ const std::wstring& szExchangeServerName)
	{
		if (m_bOutputList) return;
		const auto szTableContentsFile = strings::format(
			L"%ws\\MAILBOX_TABLE.xml", // STRING_OK
			m_szMailboxTablePathRoot.c_str());
		m_fMailboxTable = output::MyOpenFile(szTableContentsFile, true);
		if (m_fMailboxTable)
		{
			output::OutputToFile(m_fMailboxTable, output::g_szXMLHeader);
			output::OutputToFilef(m_fMailboxTable, L"<mailboxtable server=\"%ws\">\n", szExchangeServerName.c_str());
		}
	}

	void CDumpStore::DoMailboxTablePerRowWork(_In_ LPMDB lpMDB, _In_ const _SRow* lpSRow, ULONG /*ulCurRow*/)
	{
		if (!lpSRow || !m_fMailboxTable) return;
		if (m_bOutputList) return;
		auto hRes = S_OK;

		const auto lpEmailAddress = PpropFindProp(
			lpSRow->lpProps,
			lpSRow->cValues,
			PR_EMAIL_ADDRESS);

		const auto lpDisplayName = PpropFindProp(
			lpSRow->lpProps,
			lpSRow->cValues,
			PR_DISPLAY_NAME);

		output::OutputToFile(m_fMailboxTable, L"<mailbox prdisplayname=\"");
		if (mapi::CheckStringProp(lpDisplayName, PT_TSTRING))
		{
			output::OutputToFile(m_fMailboxTable, strings::LPCTSTRToWstring(lpDisplayName->Value.LPSZ));
		}
		output::OutputToFile(m_fMailboxTable, L"\" premailaddress=\"");
		if (!mapi::CheckStringProp(lpEmailAddress, PT_TSTRING))
		{
			output::OutputToFile(m_fMailboxTable, strings::LPCTSTRToWstring(lpEmailAddress->Value.LPSZ));
		}
		output::OutputToFile(m_fMailboxTable, L"\">\n");

		output::OutputSRowToFile(m_fMailboxTable, lpSRow, lpMDB);

		// build a path for our store's folder output:
		if (mapi::CheckStringProp(lpEmailAddress, PT_TSTRING) && mapi::CheckStringProp(lpDisplayName, PT_TSTRING))
		{
			const auto szTemp = strings::SanitizeFileName(strings::LPCTSTRToWstring(lpDisplayName->Value.LPSZ));

			m_szFolderPathRoot = m_szMailboxTablePathRoot + L"\\" + szTemp;

			// suppress any error here since the folder may already exist
			WC_B(CreateDirectoryW(m_szFolderPathRoot.c_str(), nullptr));
		}

		output::OutputToFile(m_fMailboxTable, L"</mailbox>\n");
	}

	void CDumpStore::EndMailboxTableWork()
	{
		if (m_bOutputList) return;
		if (m_fMailboxTable)
		{
			output::OutputToFile(m_fMailboxTable, L"</mailboxtable>\n");
			output::CloseFile(m_fMailboxTable);
		}

		m_fMailboxTable = nullptr;
	}

	void CDumpStore::BeginStoreWork()
	{
	}

	void CDumpStore::EndStoreWork()
	{
	}

	void CDumpStore::BeginFolderWork()
	{
		auto hRes = S_OK;
		m_szFolderPath = m_szFolderPathRoot + m_szFolderOffset;

		// We've done all the setup we need. If we're just outputting a list, we don't need to do the rest
		if (m_bOutputList) return;

		WC_B(CreateDirectoryW(m_szFolderPath.c_str(), nullptr));
		hRes = S_OK; // ignore the error - the directory may exist already

		// Dump the folder props to a file
		// Holds file/path name for folder props
		const auto szFolderPropsFile = m_szFolderPath + L"FOLDER_PROPS.xml"; // STRING_OK
		m_fFolderProps = output::MyOpenFile(szFolderPropsFile, true);
		if (!m_fFolderProps) return;

		output::OutputToFile(m_fFolderProps, output::g_szXMLHeader);
		output::OutputToFile(m_fFolderProps, L"<folderprops>\n");

		LPSPropValue lpAllProps = nullptr;
		ULONG cValues = 0L;

		WC_H_GETPROPS(mapi::GetPropsNULL(m_lpFolder,
			fMapiUnicode,
			&cValues,
			&lpAllProps));
		if (FAILED(hRes))
		{
			output::OutputToFilef(m_fFolderProps, L"<properties error=\"0x%08X\" />\n", hRes);
		}
		else if (lpAllProps)
		{
			output::OutputToFile(m_fFolderProps, L"<properties listtype=\"summary\">\n");

			output::OutputPropertiesToFile(m_fFolderProps, cValues, lpAllProps, m_lpFolder, m_bRetryStreamProps);

			output::OutputToFile(m_fFolderProps, L"</properties>\n");

			MAPIFreeBuffer(lpAllProps);
		}

		output::OutputToFile(m_fFolderProps, L"<HierarchyTable>\n");
	}

	void CDumpStore::DoFolderPerHierarchyTableRowWork(_In_ const _SRow* lpSRow)
	{
		if (m_bOutputList) return;
		if (!m_fFolderProps || !lpSRow) return;
		output::OutputToFile(m_fFolderProps, L"<row>\n");
		output::OutputSRowToFile(m_fFolderProps, lpSRow, m_lpMDB);
		output::OutputToFile(m_fFolderProps, L"</row>\n");
	}

	void CDumpStore::EndFolderWork()
	{
		if (m_bOutputList) return;
		if (m_fFolderProps)
		{
			output::OutputToFile(m_fFolderProps, L"</HierarchyTable>\n");
			output::OutputToFile(m_fFolderProps, L"</folderprops>\n");
			output::CloseFile(m_fFolderProps);
		}

		m_fFolderProps = nullptr;
	}

	void CDumpStore::BeginContentsTableWork(ULONG ulFlags, ULONG ulCountRows)
	{
		if (m_szFolderPath.empty()) return;
		if (m_bOutputList)
		{
			_tprintf(_T("Subject, Message Class, Filename\n"));
			return;
		}

		// Holds file/path name for contents table output
		const auto szContentsTableFile =
			ulFlags & MAPI_ASSOCIATED ? m_szFolderPath + L"ASSOCIATED_CONTENTS_TABLE.xml" : m_szFolderPath + L"CONTENTS_TABLE.xml"; // STRING_OK
		m_fFolderContents = output::MyOpenFile(szContentsTableFile, true);
		if (m_fFolderContents)
		{
			output::OutputToFile(m_fFolderContents, output::g_szXMLHeader);
			output::OutputToFilef(m_fFolderContents, L"<ContentsTable TableType=\"%ws\" messagecount=\"0x%08X\">\n",
				ulFlags & MAPI_ASSOCIATED ? L"Associated Contents" : L"Contents", // STRING_OK
				ulCountRows);
		}
	}

	// Outputs a single message's details to the screen, so as to produce a list of messages
	void OutputMessageList(
		_In_ const _SRow* lpSRow,
		_In_ const std::wstring& szFolderPath,
		bool bOutputMSG)
	{
		if (!lpSRow || szFolderPath.empty()) return;
		if (szFolderPath.length() >= MAXMSGPATH) return;

		// Get required properties from the message
		auto lpTemp = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_SUBJECT_W);
		std::wstring szSubj;
		if (lpTemp && mapi::CheckStringProp(lpTemp, PT_UNICODE))
		{
			szSubj = lpTemp->Value.lpszW;
		}
		else
		{
			lpTemp = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_SUBJECT_A);
			if (lpTemp && mapi::CheckStringProp(lpTemp, PT_STRING8))
			{
				szSubj = strings::stringTowstring(lpTemp->Value.lpszA);
			}
		}

		lpTemp = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_RECORD_KEY);
		LPSBinary lpRecordKey = nullptr;
		if (lpTemp && PR_RECORD_KEY == lpTemp->ulPropTag)
		{
			lpRecordKey = &lpTemp->Value.bin;
		}

		const auto lpMessageClass = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, CHANGE_PROP_TYPE(PR_MESSAGE_CLASS, PT_UNSPECIFIED));

		wprintf(L"\"%ls\"", szSubj.c_str());
		if (lpMessageClass)
		{
			if (PT_STRING8 == PROP_TYPE(lpMessageClass->ulPropTag))
			{
				wprintf(L",\"%hs\"", lpMessageClass->Value.lpszA ? lpMessageClass->Value.lpszA : "");
			}
			else if (PT_UNICODE == PROP_TYPE(lpMessageClass->ulPropTag))
			{
				wprintf(L",\"%ls\"", lpMessageClass->Value.lpszW ? lpMessageClass->Value.lpszW : L"");
			}
		}

		auto szExt = L".xml"; // STRING_OK
		if (bOutputMSG) szExt = L".msg"; // STRING_OK

		auto szFileName = file::BuildFileNameAndPath(szExt, szSubj, szFolderPath, lpRecordKey);
		wprintf(L",\"%ls\"\n", szFileName.c_str());
	}

	bool CDumpStore::DoContentsTablePerRowWork(_In_ const _SRow* lpSRow, ULONG ulCurRow)
	{
		if (m_bOutputList)
		{
			OutputMessageList(lpSRow, m_szFolderPath, m_bOutputMSG);
			return false;
		}
		if (!m_fFolderContents || !m_lpFolder) return true;

		output::OutputToFilef(m_fFolderContents, L"<message num=\"0x%08X\">\n", ulCurRow);

		output::OutputSRowToFile(m_fFolderContents, lpSRow, m_lpFolder);

		output::OutputToFile(m_fFolderContents, L"</message>\n");
		return true;
	}

	void CDumpStore::EndContentsTableWork()
	{
		if (m_bOutputList) return;
		if (m_fFolderContents)
		{
			output::OutputToFile(m_fFolderContents, L"</ContentsTable>\n");
			output::CloseFile(m_fFolderContents);
		}

		m_fFolderContents = nullptr;
	}

	// TODO: This fails in unicode builds since PR_RTF_COMPRESSED is always ascii.
	void OutputBody(_In_ FILE* fMessageProps, _In_ LPMESSAGE lpMessage, ULONG ulBodyTag, _In_ const std::wstring& szBodyName, bool bWrapEx, ULONG ulCPID)
	{
		auto hRes = S_OK;
		LPSTREAM lpStream = nullptr;
		LPSTREAM lpRTFUncompressed = nullptr;
		LPSTREAM lpOutputStream = nullptr;

		WC_MAPI(lpMessage->OpenProperty(
			ulBodyTag,
			&IID_IStream,
			STGM_READ,
			NULL,
			reinterpret_cast<LPUNKNOWN *>(&lpStream)));
		// The only error we suppress is MAPI_E_NOT_FOUND, so if a body type isn't in the output, it wasn't on the message
		if (MAPI_E_NOT_FOUND != hRes)
		{
			output::OutputToFilef(fMessageProps, L"<body property=\"%ws\"", szBodyName.c_str());
			if (!lpStream)
			{
				output::OutputToFilef(fMessageProps, L" error=\"0x%08X\">\n", hRes);
			}
			else
			{
				if (PR_RTF_COMPRESSED != ulBodyTag)
				{
					lpOutputStream = lpStream;
				}
				else // If we're outputting RTF, we need to wrap it first
				{
					if (bWrapEx)
					{
						ULONG ulStreamFlags = NULL;
						WC_H(mapi::WrapStreamForRTF(
							lpStream,
							true,
							MAPI_NATIVE_BODY,
							ulCPID,
							CP_ACP, // requesting ANSI code page - check if this will be valid in UNICODE builds
							&lpRTFUncompressed,
							&ulStreamFlags));
						auto szFlags = interpretprop::InterpretFlags(flagStreamFlag, ulStreamFlags);
						output::OutputToFilef(fMessageProps, L" ulStreamFlags = \"0x%08X\" szStreamFlags= \"%ws\"", ulStreamFlags, szFlags.c_str());
						output::OutputToFilef(fMessageProps, L" CodePageIn = \"%u\" CodePageOut = \"%d\"", ulCPID, CP_ACP);
					}
					else
					{
						WC_H(mapi::WrapStreamForRTF(
							lpStream,
							false,
							NULL,
							NULL,
							NULL,
							&lpRTFUncompressed,
							nullptr));
					}
					if (!lpRTFUncompressed || FAILED(hRes))
					{
						output::OutputToFilef(fMessageProps, L" rtfWrapError=\"0x%08X\"", hRes);
					}
					else
					{
						lpOutputStream = lpRTFUncompressed;
					}
				}

				output::OutputToFile(fMessageProps, L">\n");
				if (lpOutputStream)
				{
					output::OutputCDataOpen(DBGNoDebug, fMessageProps);
					output::OutputStreamToFile(fMessageProps, lpOutputStream);
					output::OutputCDataClose(DBGNoDebug, fMessageProps);
				}
			}

			output::OutputToFile(fMessageProps, L"</body>\n");
		}

		if (lpRTFUncompressed) lpRTFUncompressed->Release();
		if (lpStream) lpStream->Release();
	}

	void OutputMessageXML(
		_In_ LPMESSAGE lpMessage,
		_In_ LPVOID lpParentMessageData,
		_In_ const std::wstring& szMessageFileName,
		_In_ const std::wstring& szFolderPath,
		bool bRetryStreamProps,
		_Deref_out_ LPVOID* lpData)
	{
		if (!lpMessage || !lpData) return;

		auto hRes = S_OK;

		*lpData = static_cast<LPVOID>(new(MessageData));
		if (!*lpData) return;

		auto lpMsgData = static_cast<LPMESSAGEDATA>(*lpData);

		LPSPropValue lpAllProps = nullptr;
		ULONG cValues = 0L;

		// Get all props, asking for UNICODE string properties
		WC_H_GETPROPS(mapi::GetPropsNULL(lpMessage,
			MAPI_UNICODE,
			&cValues,
			&lpAllProps));
		if (hRes == MAPI_E_BAD_CHARWIDTH)
		{
			// Didn't like MAPI_UNICODE - fall back
			hRes = S_OK;

			WC_H_GETPROPS(mapi::GetPropsNULL(lpMessage,
				NULL,
				&cValues,
				&lpAllProps));
		}

		// If we've got a parent message, we're an attachment - use attachment filename logic
		if (lpParentMessageData)
		{
			// Take the parent filename/path, remove any extension, and append -Attach.xml
			// Should we append something for attachment number?

			// Copy the source string over
			lpMsgData->szFilePath = (static_cast<LPMESSAGEDATA>(lpParentMessageData))->szFilePath;

			// Remove any extension
			lpMsgData->szFilePath = lpMsgData->szFilePath.substr(0, lpMsgData->szFilePath.find_last_of(L'.'));

			// Update file name and add extension
			lpMsgData->szFilePath += strings::format(L"-Attach%u.xml", (static_cast<LPMESSAGEDATA>(lpParentMessageData))->ulCurAttNum); // STRING_OK

			output::OutputToFilef(static_cast<LPMESSAGEDATA>(lpParentMessageData)->fMessageProps, L"<embeddedmessage path=\"%ws\"/>\n", lpMsgData->szFilePath.c_str());
		}
		else if (!szMessageFileName.empty()) // if we've got a file name, use it
		{
			lpMsgData->szFilePath = szMessageFileName;
		}
		else
		{
			std::wstring szSubj; // BuildFileNameAndPath will substitute a subject if we don't find one
			LPSBinary lpRecordKey = nullptr;

			auto lpTemp = PpropFindProp(lpAllProps, cValues, PR_SUBJECT_W);
			if (lpTemp && mapi::CheckStringProp(lpTemp, PT_UNICODE))
			{
				szSubj = lpTemp->Value.lpszW;
			}
			else
			{
				lpTemp = PpropFindProp(lpAllProps, cValues, PR_SUBJECT_A);
				if (lpTemp && mapi::CheckStringProp(lpTemp, PT_STRING8))
				{
					szSubj = strings::stringTowstring(lpTemp->Value.lpszA);
				}
			}

			lpTemp = PpropFindProp(lpAllProps, cValues, PR_RECORD_KEY);
			if (lpTemp && PR_RECORD_KEY == lpTemp->ulPropTag)
			{
				lpRecordKey = &lpTemp->Value.bin;
			}

			lpMsgData->szFilePath = file::BuildFileNameAndPath(L".xml", szSubj, szFolderPath, lpRecordKey); // STRING_OK
		}

		if (!lpMsgData->szFilePath.empty())
		{
			output::DebugPrint(DBGGeneric, L"OutputMessagePropertiesToFile: Saving to \"%ws\"\n", lpMsgData->szFilePath.c_str());
			lpMsgData->fMessageProps = output::MyOpenFile(lpMsgData->szFilePath, true);

			if (lpMsgData->fMessageProps)
			{
				output::OutputToFile(lpMsgData->fMessageProps, output::g_szXMLHeader);
				output::OutputToFile(lpMsgData->fMessageProps, L"<message>\n");
				if (lpAllProps)
				{
					output::OutputToFile(lpMsgData->fMessageProps, L"<properties listtype=\"summary\">\n");
#define NUMPROPS 9
					static const SizedSPropTagArray(NUMPROPS, sptCols) =
					{
					NUMPROPS,
						{
							PR_MESSAGE_CLASS_W,
							PR_SUBJECT_W,
							PR_SENDER_ADDRTYPE_W,
							PR_SENDER_EMAIL_ADDRESS_W,
							PR_MESSAGE_DELIVERY_TIME,
							PR_ENTRYID,
							PR_SEARCH_KEY,
							PR_RECORD_KEY,
							PR_INTERNET_CPID
						},
					};

					for (auto column : sptCols.aulPropTag)
					{
						const auto lpTemp = PpropFindProp(lpAllProps, cValues, column);
						if (lpTemp)
						{
							output::OutputPropertyToFile(lpMsgData->fMessageProps, lpTemp, lpMessage, bRetryStreamProps);
						}
					}

					output::OutputToFile(lpMsgData->fMessageProps, L"</properties>\n");
				}

				// Log Body
				OutputBody(lpMsgData->fMessageProps, lpMessage, PR_BODY, L"PR_BODY", false, NULL);
				OutputBody(lpMsgData->fMessageProps, lpMessage, PR_BODY_HTML, L"PR_BODY_HTML", false, NULL);
				OutputBody(lpMsgData->fMessageProps, lpMessage, PR_RTF_COMPRESSED, L"PR_RTF_COMPRESSED", false, NULL);

				ULONG ulInCodePage = CP_ACP; // picking CP_ACP as our default
				const auto lpTemp = PpropFindProp(lpAllProps, cValues, PR_INTERNET_CPID);
				if (lpTemp && PR_INTERNET_CPID == lpTemp->ulPropTag)
				{
					ulInCodePage = lpTemp->Value.l;
				}

				OutputBody(lpMsgData->fMessageProps, lpMessage, PR_RTF_COMPRESSED, L"WrapCompressedRTFEx best body", true, ulInCodePage);

				if (lpAllProps)
				{
					output::OutputToFile(lpMsgData->fMessageProps, L"<properties listtype=\"FullPropList\">\n");

					output::OutputPropertiesToFile(lpMsgData->fMessageProps, cValues, lpAllProps, lpMessage, bRetryStreamProps);

					output::OutputToFile(lpMsgData->fMessageProps, L"</properties>\n");
				}
			}
		}

		MAPIFreeBuffer(lpAllProps);
	}

	void OutputMessageMSG(
		_In_ LPMESSAGE lpMessage,
		_In_ const std::wstring& szFolderPath)
	{
		auto hRes = S_OK;

		enum
		{
			msgPR_SUBJECT_W,
			msgPR_RECORD_KEY,
			msgNUM_COLS
		};

		static const SizedSPropTagArray(msgNUM_COLS, msgProps) =
		{
		msgNUM_COLS,
			{
				PR_SUBJECT_W,
				PR_RECORD_KEY
			}
		};

		if (!lpMessage || szFolderPath.empty()) return;

		output::DebugPrint(DBGGeneric, L"OutputMessageMSG: Saving message to \"%ws\"\n", szFolderPath.c_str());

		std::wstring szSubj;

		ULONG cProps = 0;
		LPSPropValue lpsProps = nullptr;
		LPSBinary lpRecordKey = nullptr;

		// Get required properties from the message
		EC_H_GETPROPS(lpMessage->GetProps(
			LPSPropTagArray(&msgProps),
			fMapiUnicode,
			&cProps,
			&lpsProps));

		if (cProps == 2 && lpsProps)
		{
			if (mapi::CheckStringProp(&lpsProps[msgPR_SUBJECT_W], PT_UNICODE))
			{
				szSubj = lpsProps[msgPR_SUBJECT_W].Value.lpszW;
			}
			if (PR_RECORD_KEY == lpsProps[msgPR_RECORD_KEY].ulPropTag)
			{
				lpRecordKey = &lpsProps[msgPR_RECORD_KEY].Value.bin;
			}
		}

		auto szFileName = file::BuildFileNameAndPath(L".msg", szSubj, szFolderPath, lpRecordKey); // STRING_OK
		if (!szFileName.empty())
		{
			output::DebugPrint(DBGGeneric, L"Saving to = \"%ws\"\n", szFileName.c_str());

			WC_H(file::SaveToMSG(
				lpMessage,
				szFileName,
				fMapiUnicode != 0,
				nullptr,
				false));
		}
	}

	bool CDumpStore::BeginMessageWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpParentMessageData, _Deref_out_opt_ LPVOID* lpData)
	{
		if (lpData) *lpData = nullptr;
		if (lpParentMessageData && !m_bOutputAttachments) return false;
		if (m_bOutputList) return false;

		if (m_bOutputMSG)
		{
			OutputMessageMSG(lpMessage, m_szFolderPath);
			return false; // no more work necessary
		}

		OutputMessageXML(lpMessage, lpParentMessageData, m_szMessageFileName, m_szFolderPath, m_bRetryStreamProps, lpData);
		return true;
	}

	bool CDumpStore::BeginRecipientWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
	{
		if (!lpData) return false;
		if (m_bOutputMSG) return false; // When outputting message files, no recipient work is needed
		if (m_bOutputList) return false;
		output::OutputToFile((static_cast<LPMESSAGEDATA>(lpData))->fMessageProps, L"<recipients>\n");
		return true;
	}

	void CDumpStore::DoMessagePerRecipientWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData, _In_ const _SRow* lpSRow, ULONG ulCurRow)
	{
		if (!lpMessage || !lpData || !lpSRow) return;
		if (m_bOutputMSG) return; // When outputting message files, no recipient work is needed
		if (m_bOutputList) return;

		const auto lpMsgData = static_cast<LPMESSAGEDATA>(lpData);

		output::OutputToFilef(lpMsgData->fMessageProps, L"<recipient num=\"0x%08X\">\n", ulCurRow);

		output::OutputSRowToFile(lpMsgData->fMessageProps, lpSRow, lpMessage);

		output::OutputToFile(lpMsgData->fMessageProps, L"</recipient>\n");
	}

	void CDumpStore::EndRecipientWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
	{
		if (!lpData) return;
		if (m_bOutputMSG) return; // When outputting message files, no recipient work is needed
		if (m_bOutputList) return;
		output::OutputToFile((static_cast<LPMESSAGEDATA>(lpData))->fMessageProps, L"</recipients>\n");
	}

	bool CDumpStore::BeginAttachmentWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
	{
		if (!lpData) return false;
		if (m_bOutputMSG) return false; // When outputting message files, no attachment work is needed
		if (m_bOutputList) return false;
		output::OutputToFile((static_cast<LPMESSAGEDATA>(lpData))->fMessageProps, L"<attachments>\n");
		return true;
	}

	void CDumpStore::DoMessagePerAttachmentWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData, _In_ const _SRow* lpSRow, _In_ LPATTACH lpAttach, ULONG ulCurRow)
	{
		if (!lpMessage || !lpData || !lpSRow) return;
		if (m_bOutputMSG) return; // When outputting message files, no attachment work is needed
		if (m_bOutputList) return;

		const auto lpMsgData = static_cast<LPMESSAGEDATA>(lpData);

		lpMsgData->ulCurAttNum = ulCurRow; // set this so we can pull it if this is an embedded message

		auto hRes = S_OK;

		output::OutputToFilef(lpMsgData->fMessageProps, L"<attachment num=\"0x%08X\" filename=\"", ulCurRow);

		auto lpAttachName = PpropFindProp(
			lpSRow->lpProps,
			lpSRow->cValues,
			PR_ATTACH_FILENAME);

		if (!lpAttachName || !mapi::CheckStringProp(lpAttachName, PT_TSTRING))
		{
			lpAttachName = PpropFindProp(
				lpSRow->lpProps,
				lpSRow->cValues,
				PR_DISPLAY_NAME);
		}

		if (lpAttachName && mapi::CheckStringProp(lpAttachName, PT_TSTRING))
			output::OutputToFile(lpMsgData->fMessageProps, strings::LPCTSTRToWstring(lpAttachName->Value.LPSZ));
		else
			output::OutputToFile(lpMsgData->fMessageProps, L"PR_ATTACH_FILENAME not found");

		output::OutputToFile(lpMsgData->fMessageProps, L"\">\n");

		output::OutputToFile(lpMsgData->fMessageProps, L"\t<tableprops>\n");
		output::OutputSRowToFile(lpMsgData->fMessageProps, lpSRow, lpMessage);

		output::OutputToFile(lpMsgData->fMessageProps, L"\t</tableprops>\n");

		if (lpAttach)
		{
			ULONG ulAllProps = 0;
			LPSPropValue lpAllProps = nullptr;
			// Let's get all props from the message and dump them.
			WC_H_GETPROPS(mapi::GetPropsNULL(lpAttach,
				fMapiUnicode,
				&ulAllProps,
				&lpAllProps));
			if (lpAllProps)
			{
				output::OutputToFile(lpMsgData->fMessageProps, L"\t<getprops>\n");
				output::OutputPropertiesToFile(lpMsgData->fMessageProps, ulAllProps, lpAllProps, lpAttach, m_bRetryStreamProps);
				output::OutputToFile(lpMsgData->fMessageProps, L"\t</getprops>\n");
			}

			MAPIFreeBuffer(lpAllProps);
			lpAllProps = nullptr;
		}

		output::OutputToFile(lpMsgData->fMessageProps, L"</attachment>\n");
	}

	void CDumpStore::EndAttachmentWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
	{
		if (!lpData) return;
		if (m_bOutputMSG) return; // When outputting message files, no attachment work is needed
		if (m_bOutputList) return;
		output::OutputToFile((static_cast<LPMESSAGEDATA>(lpData))->fMessageProps, L"</attachments>\n");
	}

	void CDumpStore::EndMessageWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
	{
		if (m_bOutputMSG) return; // When outputting message files, no end message work is needed
		if (m_bOutputList) return;
		const auto lpMsgData = static_cast<LPMESSAGEDATA>(lpData);

		if (lpMsgData->fMessageProps)
		{
			output::OutputToFile(lpMsgData->fMessageProps, L"</message>\n");
			output::CloseFile(lpMsgData->fMessageProps);
		}
		delete lpMsgData;
	}
}