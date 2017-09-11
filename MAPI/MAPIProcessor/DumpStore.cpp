// Routines used in dumping the contents of the Exchange store
// in to a log file

#include "stdafx.h"
#include <MAPI/MAPIProcessor/DumpStore.h>
#include <MAPI/MAPIFunctions.h>
#include <Interpret/String.h>
#include <IO/File.h>
#include <Interpret/InterpretProp2.h>
#include <Interpret/ExtraPropTags.h>

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
	if (m_fFolderProps) CloseFile(m_fFolderProps);
	if (m_fFolderContents) CloseFile(m_fFolderContents);
	if (m_fMailboxTable) CloseFile(m_fMailboxTable);
}

void CDumpStore::InitMessagePath(_In_ const wstring& szMessageFileName)
{
	m_szMessageFileName = szMessageFileName;
}

void CDumpStore::InitFolderPathRoot(_In_ const wstring& szFolderPathRoot)
{
	m_szFolderPathRoot = szFolderPathRoot;
	if (m_szFolderPathRoot.length() >= MAXMSGPATH)
	{
		DebugPrint(DBGGeneric, L"InitFolderPathRoot: \"%ws\" length (%d) greater than max length (%d)\n", m_szFolderPathRoot.c_str(), m_szFolderPathRoot.length(), MAXMSGPATH);
	}
}

void CDumpStore::InitMailboxTablePathRoot(_In_ const wstring& szMailboxTablePathRoot)
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

void CDumpStore::BeginMailboxTableWork(_In_ const wstring& szExchangeServerName)
{
	if (m_bOutputList) return;
	auto szTableContentsFile = format(
		L"%ws\\MAILBOX_TABLE.xml", // STRING_OK
		m_szMailboxTablePathRoot.c_str());
	m_fMailboxTable = MyOpenFile(szTableContentsFile, true);
	if (m_fMailboxTable)
	{
		OutputToFile(m_fMailboxTable, g_szXMLHeader);
		OutputToFilef(m_fMailboxTable, L"<mailboxtable server=\"%ws\">\n", szExchangeServerName.c_str());
	}
}

void CDumpStore::DoMailboxTablePerRowWork(_In_ LPMDB lpMDB, _In_ const LPSRow lpSRow, ULONG /*ulCurRow*/)
{
	if (!lpSRow || !m_fMailboxTable) return;
	if (m_bOutputList) return;
	auto hRes = S_OK;

	auto lpEmailAddress = PpropFindProp(
		lpSRow->lpProps,
		lpSRow->cValues,
		PR_EMAIL_ADDRESS);

	auto lpDisplayName = PpropFindProp(
		lpSRow->lpProps,
		lpSRow->cValues,
		PR_DISPLAY_NAME);

	OutputToFile(m_fMailboxTable, L"<mailbox prdisplayname=\"");
	if (CheckStringProp(lpDisplayName, PT_TSTRING))
	{
		OutputToFile(m_fMailboxTable, LPCTSTRToWstring(lpDisplayName->Value.LPSZ));
	}
	OutputToFile(m_fMailboxTable, L"\" premailaddress=\"");
	if (!CheckStringProp(lpEmailAddress, PT_TSTRING))
	{
		OutputToFile(m_fMailboxTable, LPCTSTRToWstring(lpEmailAddress->Value.LPSZ));
	}
	OutputToFile(m_fMailboxTable, L"\">\n");

	OutputSRowToFile(m_fMailboxTable, lpSRow, lpMDB);

	// build a path for our store's folder output:
	if (CheckStringProp(lpEmailAddress, PT_TSTRING) && CheckStringProp(lpDisplayName, PT_TSTRING))
	{
		auto szTemp = SanitizeFileName(LPCTSTRToWstring(lpDisplayName->Value.LPSZ));

		m_szFolderPathRoot = m_szMailboxTablePathRoot + L"\\" + szTemp;

		// suppress any error here since the folder may already exist
		WC_B(CreateDirectoryW(m_szFolderPathRoot.c_str(), nullptr));
	}

	OutputToFile(m_fMailboxTable, L"</mailbox>\n");
}

void CDumpStore::EndMailboxTableWork()
{
	if (m_bOutputList) return;
	if (m_fMailboxTable)
	{
		OutputToFile(m_fMailboxTable, L"</mailboxtable>\n");
		CloseFile(m_fMailboxTable);
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
	auto szFolderPropsFile = m_szFolderPath + L"FOLDER_PROPS.xml"; // STRING_OK
	m_fFolderProps = MyOpenFile(szFolderPropsFile, true);
	if (!m_fFolderProps) return;

	OutputToFile(m_fFolderProps, g_szXMLHeader);
	OutputToFile(m_fFolderProps, L"<folderprops>\n");

	LPSPropValue lpAllProps = nullptr;
	ULONG cValues = 0L;

	WC_H_GETPROPS(GetPropsNULL(m_lpFolder,
		fMapiUnicode,
		&cValues,
		&lpAllProps));
	if (FAILED(hRes))
	{
		OutputToFilef(m_fFolderProps, L"<properties error=\"0x%08X\" />\n", hRes);
	}
	else if (lpAllProps)
	{
		OutputToFile(m_fFolderProps, L"<properties listtype=\"summary\">\n");

		OutputPropertiesToFile(m_fFolderProps, cValues, lpAllProps, m_lpFolder, m_bRetryStreamProps);

		OutputToFile(m_fFolderProps, L"</properties>\n");

		MAPIFreeBuffer(lpAllProps);
	}

	OutputToFile(m_fFolderProps, L"<HierarchyTable>\n");
}

void CDumpStore::DoFolderPerHierarchyTableRowWork(_In_ const LPSRow lpSRow)
{
	if (m_bOutputList) return;
	if (!m_fFolderProps || !lpSRow) return;
	OutputToFile(m_fFolderProps, L"<row>\n");
	OutputSRowToFile(m_fFolderProps, lpSRow, m_lpMDB);
	OutputToFile(m_fFolderProps, L"</row>\n");
}

void CDumpStore::EndFolderWork()
{
	if (m_bOutputList) return;
	if (m_fFolderProps)
	{
		OutputToFile(m_fFolderProps, L"</HierarchyTable>\n");
		OutputToFile(m_fFolderProps, L"</folderprops>\n");
		CloseFile(m_fFolderProps);
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
	auto szContentsTableFile =
		ulFlags & MAPI_ASSOCIATED ? m_szFolderPath + L"ASSOCIATED_CONTENTS_TABLE.xml": m_szFolderPath + L"CONTENTS_TABLE.xml"; // STRING_OK
	m_fFolderContents = MyOpenFile(szContentsTableFile, true);
	if (m_fFolderContents)
	{
		OutputToFile(m_fFolderContents, g_szXMLHeader);
		OutputToFilef(m_fFolderContents, L"<ContentsTable TableType=\"%ws\" messagecount=\"0x%08X\">\n",
			ulFlags & MAPI_ASSOCIATED ? L"Associated Contents" : L"Contents", // STRING_OK
			ulCountRows);
	}
}

// Outputs a single message's details to the screen, so as to produce a list of messages
void OutputMessageList(
	_In_ const LPSRow lpSRow,
	_In_ const wstring& szFolderPath,
	bool bOutputMSG)
{
	if (!lpSRow || szFolderPath.empty()) return;
	if (szFolderPath.length() >= MAXMSGPATH) return;

	// Get required properties from the message
	auto lpTemp = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_SUBJECT_W);
	wstring szSubj;
	if (lpTemp && CheckStringProp(lpTemp, PT_UNICODE))
	{
		szSubj = lpTemp->Value.lpszW;
	}
	else
	{
		lpTemp = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_SUBJECT_A);
		if (lpTemp && CheckStringProp(lpTemp, PT_STRING8))
		{
			szSubj = stringTowstring(lpTemp->Value.lpszA);
		}
	}

	lpTemp = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_RECORD_KEY);
	LPSBinary lpRecordKey = nullptr;
	if (lpTemp && PR_RECORD_KEY == lpTemp->ulPropTag)
	{
		lpRecordKey = &lpTemp->Value.bin;
	}

	auto lpMessageClass = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, CHANGE_PROP_TYPE(PR_MESSAGE_CLASS, PT_UNSPECIFIED));

	wprintf(L"\"%ws\"", szSubj.c_str());
	if (lpMessageClass)
	{
		if (PT_STRING8 == PROP_TYPE(lpMessageClass->ulPropTag))
		{
			wprintf(L",\"%hs\"", lpMessageClass->Value.lpszA ? lpMessageClass->Value.lpszA : "");
		}
		else if (PT_UNICODE == PROP_TYPE(lpMessageClass->ulPropTag))
		{
			wprintf(L",\"%ws\"", lpMessageClass->Value.lpszW ? lpMessageClass->Value.lpszW : L"");
		}
	}

	auto szExt = L".xml"; // STRING_OK
	if (bOutputMSG) szExt = L".msg"; // STRING_OK

	auto szFileName = BuildFileNameAndPath(szExt, szSubj, szFolderPath, lpRecordKey);
	wprintf(L",\"%ws\"\n", szFileName.c_str());
}

bool CDumpStore::DoContentsTablePerRowWork(_In_ const LPSRow lpSRow, ULONG ulCurRow)
{
	if (m_bOutputList)
	{
		OutputMessageList(lpSRow, m_szFolderPath, m_bOutputMSG);
		return false;
	}
	if (!m_fFolderContents || !m_lpFolder) return true;

	OutputToFilef(m_fFolderContents, L"<message num=\"0x%08X\">\n", ulCurRow);

	OutputSRowToFile(m_fFolderContents, lpSRow, m_lpFolder);

	OutputToFile(m_fFolderContents, L"</message>\n");
	return true;
}

void CDumpStore::EndContentsTableWork()
{
	if (m_bOutputList) return;
	if (m_fFolderContents)
	{
		OutputToFile(m_fFolderContents, L"</ContentsTable>\n");
		CloseFile(m_fFolderContents);
	}

	m_fFolderContents = nullptr;
}

// TODO: This fails in unicode builds since PR_RTF_COMPRESSED is always ascii.
void OutputBody(_In_ FILE* fMessageProps, _In_ LPMESSAGE lpMessage, ULONG ulBodyTag, _In_ const wstring& szBodyName, bool bWrapEx, ULONG ulCPID)
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
		OutputToFilef(fMessageProps, L"<body property=\"%ws\"", szBodyName.c_str());
		if (!lpStream)
		{
			OutputToFilef(fMessageProps, L" error=\"0x%08X\">\n", hRes);
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
					WC_H(WrapStreamForRTF(
						lpStream,
						true,
						MAPI_NATIVE_BODY,
						ulCPID,
						CP_ACP, // requesting ANSI code page - check if this will be valid in UNICODE builds
						&lpRTFUncompressed,
						&ulStreamFlags));
					auto szFlags = InterpretFlags(flagStreamFlag, ulStreamFlags);
					OutputToFilef(fMessageProps, L" ulStreamFlags = \"0x%08X\" szStreamFlags= \"%ws\"", ulStreamFlags, szFlags.c_str());
					OutputToFilef(fMessageProps, L" CodePageIn = \"%u\" CodePageOut = \"%d\"", ulCPID, CP_ACP);
				}
				else
				{
					WC_H(WrapStreamForRTF(
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
					OutputToFilef(fMessageProps, L" rtfWrapError=\"0x%08X\"", hRes);
				}
				else
				{
					lpOutputStream = lpRTFUncompressed;
				}
			}

			OutputToFile(fMessageProps, L">\n");
			if (lpOutputStream)
			{
				OutputCDataOpen(DBGNoDebug, fMessageProps);
				OutputStreamToFile(fMessageProps, lpOutputStream);
				OutputCDataClose(DBGNoDebug, fMessageProps);
			}
		}

		OutputToFile(fMessageProps, L"</body>\n");
	}

	if (lpRTFUncompressed) lpRTFUncompressed->Release();
	if (lpStream) lpStream->Release();
}

void OutputMessageXML(
	_In_ LPMESSAGE lpMessage,
	_In_ LPVOID lpParentMessageData,
	_In_ const wstring& szMessageFileName,
	_In_ const wstring& szFolderPath,
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
	WC_H_GETPROPS(GetPropsNULL(lpMessage,
		MAPI_UNICODE,
		&cValues,
		&lpAllProps));
	if (hRes == MAPI_E_BAD_CHARWIDTH)
	{
		// Didn't like MAPI_UNICODE - fall back
		hRes = S_OK;

		WC_H_GETPROPS(GetPropsNULL(lpMessage,
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
		lpMsgData->szFilePath += format(L"-Attach%u.xml", (static_cast<LPMESSAGEDATA>(lpParentMessageData))->ulCurAttNum); // STRING_OK

		OutputToFilef(static_cast<LPMESSAGEDATA>(lpParentMessageData)->fMessageProps, L"<embeddedmessage path=\"%ws\"/>\n", lpMsgData->szFilePath.c_str());
	}
	else if (!szMessageFileName.empty()) // if we've got a file name, use it
	{
		lpMsgData->szFilePath = szMessageFileName;
	}
	else
	{
		wstring szSubj; // BuildFileNameAndPath will substitute a subject if we don't find one
		LPSBinary lpRecordKey = nullptr;

		auto lpTemp = PpropFindProp(lpAllProps, cValues, PR_SUBJECT_W);
		if (lpTemp && CheckStringProp(lpTemp, PT_UNICODE))
		{
			szSubj = lpTemp->Value.lpszW;
		}
		else
		{
			lpTemp = PpropFindProp(lpAllProps, cValues, PR_SUBJECT_A);
			if (lpTemp && CheckStringProp(lpTemp, PT_STRING8))
			{
				szSubj = stringTowstring(lpTemp->Value.lpszA);
			}
		}

		lpTemp = PpropFindProp(lpAllProps, cValues, PR_RECORD_KEY);
		if (lpTemp && PR_RECORD_KEY == lpTemp->ulPropTag)
		{
			lpRecordKey = &lpTemp->Value.bin;
		}

		lpMsgData->szFilePath = BuildFileNameAndPath(L".xml", szSubj, szFolderPath, lpRecordKey); // STRING_OK
	}

	if (!lpMsgData->szFilePath.empty())
	{
		DebugPrint(DBGGeneric, L"OutputMessagePropertiesToFile: Saving to \"%ws\"\n", lpMsgData->szFilePath.c_str());
		lpMsgData->fMessageProps = MyOpenFile(lpMsgData->szFilePath, true);

		if (lpMsgData->fMessageProps)
		{
			OutputToFile(lpMsgData->fMessageProps, g_szXMLHeader);
			OutputToFile(lpMsgData->fMessageProps, L"<message>\n");
			if (lpAllProps)
			{
				OutputToFile(lpMsgData->fMessageProps, L"<properties listtype=\"summary\">\n");
#define NUMPROPS 9
				static const SizedSPropTagArray(NUMPROPS, sptCols) =
				{
				NUMPROPS,
				PR_MESSAGE_CLASS_W,
				PR_SUBJECT_W,
				PR_SENDER_ADDRTYPE_W,
				PR_SENDER_EMAIL_ADDRESS_W,
				PR_MESSAGE_DELIVERY_TIME,
				PR_ENTRYID,
				PR_SEARCH_KEY,
				PR_RECORD_KEY,
				PR_INTERNET_CPID,
				};

				for (auto i = 0; i < NUMPROPS; i++)
				{
					auto lpTemp = PpropFindProp(lpAllProps, cValues, sptCols.aulPropTag[i]);
					if (lpTemp)
					{
						OutputPropertyToFile(lpMsgData->fMessageProps, lpTemp, lpMessage, bRetryStreamProps);
					}
				}

				OutputToFile(lpMsgData->fMessageProps, L"</properties>\n");
			}

			// Log Body
			OutputBody(lpMsgData->fMessageProps, lpMessage, PR_BODY, L"PR_BODY", false, NULL);
			OutputBody(lpMsgData->fMessageProps, lpMessage, PR_BODY_HTML, L"PR_BODY_HTML", false, NULL);
			OutputBody(lpMsgData->fMessageProps, lpMessage, PR_RTF_COMPRESSED, L"PR_RTF_COMPRESSED", false, NULL);

			ULONG ulInCodePage = CP_ACP; // picking CP_ACP as our default
			auto lpTemp = PpropFindProp(lpAllProps, cValues, PR_INTERNET_CPID);
			if (lpTemp && PR_INTERNET_CPID == lpTemp->ulPropTag)
			{
				ulInCodePage = lpTemp->Value.l;
			}

			OutputBody(lpMsgData->fMessageProps, lpMessage, PR_RTF_COMPRESSED, L"WrapCompressedRTFEx best body", true, ulInCodePage);

			if (lpAllProps)
			{
				OutputToFile(lpMsgData->fMessageProps, L"<properties listtype=\"FullPropList\">\n");

				OutputPropertiesToFile(lpMsgData->fMessageProps, cValues, lpAllProps, lpMessage, bRetryStreamProps);

				OutputToFile(lpMsgData->fMessageProps, L"</properties>\n");
			}
		}
	}

	MAPIFreeBuffer(lpAllProps);
}

void OutputMessageMSG(
	_In_ LPMESSAGE lpMessage,
	_In_ const wstring& szFolderPath)
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
	PR_SUBJECT_W,
	PR_RECORD_KEY
	};

	if (!lpMessage || szFolderPath.empty()) return;

	DebugPrint(DBGGeneric, L"OutputMessageMSG: Saving message to \"%ws\"\n", szFolderPath.c_str());

	wstring szSubj;

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
		if (CheckStringProp(&lpsProps[msgPR_SUBJECT_W], PT_UNICODE))
		{
			szSubj = lpsProps[msgPR_SUBJECT_W].Value.lpszW;
		}
		if (PR_RECORD_KEY == lpsProps[msgPR_RECORD_KEY].ulPropTag)
		{
			lpRecordKey = &lpsProps[msgPR_RECORD_KEY].Value.bin;
		}
	}

	auto szFileName = BuildFileNameAndPath(L".msg", szSubj, szFolderPath, lpRecordKey); // STRING_OK
	if (!szFileName.empty())
	{
		DebugPrint(DBGGeneric, L"Saving to = \"%ws\"\n", szFileName.c_str());

		WC_H(SaveToMSG(
			lpMessage,
			szFileName,
			fMapiUnicode ? true : false,
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
	OutputToFile((static_cast<LPMESSAGEDATA>(lpData))->fMessageProps, L"<recipients>\n");
	return true;
}

void CDumpStore::DoMessagePerRecipientWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData, _In_ const LPSRow lpSRow, ULONG ulCurRow)
{
	if (!lpMessage || !lpData || !lpSRow) return;
	if (m_bOutputMSG) return; // When outputting message files, no recipient work is needed
	if (m_bOutputList) return;

	auto lpMsgData = static_cast<LPMESSAGEDATA>(lpData);

	OutputToFilef(lpMsgData->fMessageProps, L"<recipient num=\"0x%08X\">\n", ulCurRow);

	OutputSRowToFile(lpMsgData->fMessageProps, lpSRow, lpMessage);

	OutputToFile(lpMsgData->fMessageProps, L"</recipient>\n");
}

void CDumpStore::EndRecipientWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
{
	if (!lpData) return;
	if (m_bOutputMSG) return; // When outputting message files, no recipient work is needed
	if (m_bOutputList) return;
	OutputToFile((static_cast<LPMESSAGEDATA>(lpData))->fMessageProps, L"</recipients>\n");
}

bool CDumpStore::BeginAttachmentWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
{
	if (!lpData) return false;
	if (m_bOutputMSG) return false; // When outputting message files, no attachment work is needed
	if (m_bOutputList) return false;
	OutputToFile((static_cast<LPMESSAGEDATA>(lpData))->fMessageProps, L"<attachments>\n");
	return true;
}

void CDumpStore::DoMessagePerAttachmentWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData, _In_ const LPSRow lpSRow, _In_ LPATTACH lpAttach, ULONG ulCurRow)
{
	if (!lpMessage || !lpData || !lpSRow) return;
	if (m_bOutputMSG) return; // When outputting message files, no attachment work is needed
	if (m_bOutputList) return;

	auto lpMsgData = static_cast<LPMESSAGEDATA>(lpData);

	lpMsgData->ulCurAttNum = ulCurRow; // set this so we can pull it if this is an embedded message

	auto hRes = S_OK;

	OutputToFilef(lpMsgData->fMessageProps, L"<attachment num=\"0x%08X\" filename=\"", ulCurRow);

	auto lpAttachName = PpropFindProp(
		lpSRow->lpProps,
		lpSRow->cValues,
		PR_ATTACH_FILENAME);

	if (!lpAttachName || !CheckStringProp(lpAttachName, PT_TSTRING))
	{
		lpAttachName = PpropFindProp(
			lpSRow->lpProps,
			lpSRow->cValues,
			PR_DISPLAY_NAME);
	}

	if (lpAttachName && CheckStringProp(lpAttachName, PT_TSTRING))
		OutputToFile(lpMsgData->fMessageProps, LPCTSTRToWstring(lpAttachName->Value.LPSZ));
	else
		OutputToFile(lpMsgData->fMessageProps, L"PR_ATTACH_FILENAME not found");

	OutputToFile(lpMsgData->fMessageProps, L"\">\n");

	OutputToFile(lpMsgData->fMessageProps, L"\t<tableprops>\n");
	OutputSRowToFile(lpMsgData->fMessageProps, lpSRow, lpMessage);

	OutputToFile(lpMsgData->fMessageProps, L"\t</tableprops>\n");

	if (lpAttach)
	{
		ULONG ulAllProps = 0;
		LPSPropValue lpAllProps = nullptr;
		// Let's get all props from the message and dump them.
		WC_H_GETPROPS(GetPropsNULL(lpAttach,
			fMapiUnicode,
			&ulAllProps,
			&lpAllProps));
		if (lpAllProps)
		{
			OutputToFile(lpMsgData->fMessageProps, L"\t<getprops>\n");
			OutputPropertiesToFile(lpMsgData->fMessageProps, ulAllProps, lpAllProps, lpAttach, m_bRetryStreamProps);
			OutputToFile(lpMsgData->fMessageProps, L"\t</getprops>\n");
		}

		MAPIFreeBuffer(lpAllProps);
		lpAllProps = nullptr;
	}

	OutputToFile(lpMsgData->fMessageProps, L"</attachment>\n");
}

void CDumpStore::EndAttachmentWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
{
	if (!lpData) return;
	if (m_bOutputMSG) return; // When outputting message files, no attachment work is needed
	if (m_bOutputList) return;
	OutputToFile((static_cast<LPMESSAGEDATA>(lpData))->fMessageProps, L"</attachments>\n");
}

void CDumpStore::EndMessageWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
{
	if (m_bOutputMSG) return; // When outputting message files, no end message work is needed
	if (m_bOutputList) return;
	auto lpMsgData = static_cast<LPMESSAGEDATA>(lpData);

	if (lpMsgData->fMessageProps)
	{
		OutputToFile(lpMsgData->fMessageProps, L"</message>\n");
		CloseFile(lpMsgData->fMessageProps);
	}
	delete lpMsgData;
}