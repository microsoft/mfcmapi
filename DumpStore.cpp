//  MODULE:   DumpStore.cpp
//
//  PURPOSE:  Contains routines used in dumping the contents of the Exchange store
//            in to a log file

#include "stdafx.h"
#include "dumpstore.h"
#include "MAPIFunctions.h"
#include "String.h"
#include "file.h"
#include "InterpretProp2.h"
#include "ImportProcs.h"
#include "ExtraPropTags.h"

CDumpStore::CDumpStore()
{
	m_szMessageFileName[0] = L'\0';
	m_szMailboxTablePathRoot[0] = L'\0';
	m_szFolderPathRoot[0] = L'\0';

	m_szFolderPath = NULL;

	m_fFolderProps = NULL;
	m_fFolderContents = NULL;
	m_fMailboxTable = NULL;

	m_bRetryStreamProps = true;
	m_bOutputAttachments = true;
	m_bOutputMSG = false;
	m_bOutputList = false;
} // CDumpStore::CDumpStore

CDumpStore::~CDumpStore()
{
	if (m_fFolderProps) CloseFile(m_fFolderProps);
	if (m_fFolderContents) CloseFile(m_fFolderContents);
	if (m_fMailboxTable) CloseFile(m_fMailboxTable);
	MAPIFreeBuffer(m_szFolderPath);
} // CDumpStore::~CDumpStore

// --------------------------------------------------------------------------------- //

void CDumpStore::InitMessagePath(_In_z_ LPCWSTR szMessageFileName)
{
	if (szMessageFileName)
	{
		HRESULT hRes = S_OK;
		EC_H(StringCchCopyW(m_szMessageFileName, _countof(m_szMessageFileName), szMessageFileName));
	}
	else
	{
		m_szMessageFileName[0] = L'\0';
	}
} // CDumpStore::InitMessagePath

void CDumpStore::InitFolderPathRoot(_In_z_ LPCWSTR szFolderPathRoot)
{
	if (szFolderPathRoot)
	{
		HRESULT hRes = S_OK;
		EC_H(StringCchCopyW(m_szFolderPathRoot, _countof(m_szFolderPathRoot), szFolderPathRoot));
	}
	else
	{
		m_szFolderPathRoot[0] = L'\0';
	}
} // CDumpStore::InitFolderPathRoot

void CDumpStore::InitMailboxTablePathRoot(_In_z_ LPCWSTR szMailboxTablePathRoot)
{
	if (szMailboxTablePathRoot)
	{
		HRESULT hRes = S_OK;
		EC_H(StringCchCopyW(m_szMailboxTablePathRoot, _countof(m_szMailboxTablePathRoot), szMailboxTablePathRoot));
	}
	else
	{
		m_szMailboxTablePathRoot[0] = L'\0';
	}
} // CDumpStore::InitMailboxTablePathRoot

void CDumpStore::EnableMSG()
{
	m_bOutputMSG = true;
} // CDumpStore::EnableMSG

void CDumpStore::EnableList()
{
	m_bOutputList = true;
} // CDumpStore::EnableMSG

void CDumpStore::DisableStreamRetry()
{
	m_bRetryStreamProps = false;
}

void CDumpStore::DisableEmbeddedAttachments()
{
	m_bOutputAttachments = false;
}

// --------------------------------------------------------------------------------- //

void CDumpStore::BeginMailboxTableWork(_In_z_ LPCTSTR szExchangeServerName)
{
	if (m_bOutputList) return;
	HRESULT hRes = S_OK;
	WCHAR	szTableContentsFile[MAX_PATH];
	WC_H(StringCchPrintfW(szTableContentsFile, _countof(szTableContentsFile),
		L"%s\\MAILBOX_TABLE.xml", // STRING_OK
		m_szMailboxTablePathRoot));
	m_fMailboxTable = MyOpenFile(szTableContentsFile, true);
	if (m_fMailboxTable)
	{
		OutputToFile(m_fMailboxTable, g_szXMLHeader);
		OutputToFilef(m_fMailboxTable, _T("<mailboxtable server=\"%s\">\n"), szExchangeServerName);
	}
} // CDumpStore::BeginMailboxTableWork

void CDumpStore::DoMailboxTablePerRowWork(_In_ LPMDB lpMDB, _In_ LPSRow lpSRow, ULONG /*ulCurRow*/)
{
	if (!lpSRow || !m_fMailboxTable) return;
	if (m_bOutputList) return;
	HRESULT			hRes = S_OK;
	LPSPropValue	lpEmailAddress = NULL;
	LPSPropValue	lpDisplayName = NULL;

	lpEmailAddress = PpropFindProp(
		lpSRow->lpProps,
		lpSRow->cValues,
		PR_EMAIL_ADDRESS);

	lpDisplayName = PpropFindProp(
		lpSRow->lpProps,
		lpSRow->cValues,
		PR_DISPLAY_NAME);

	OutputToFile(m_fMailboxTable, _T("<mailbox prdisplayname=\""));
	if (CheckStringProp(lpDisplayName, PT_TSTRING))
	{
		OutputToFile(m_fMailboxTable, lpDisplayName->Value.LPSZ);
	}
	OutputToFile(m_fMailboxTable, _T("\" premailaddress=\""));
	if (!CheckStringProp(lpEmailAddress, PT_TSTRING))
	{
		OutputToFile(m_fMailboxTable, lpEmailAddress->Value.LPSZ);
	}
	OutputToFile(m_fMailboxTable, _T("\">\n"));

	OutputSRowToFile(m_fMailboxTable, lpSRow, lpMDB);

	// build a path for our store's folder output:
	if (CheckStringProp(lpEmailAddress, PT_TSTRING) && CheckStringProp(lpDisplayName, PT_TSTRING))
	{
		TCHAR szTemp[MAX_PATH / 2] = { 0 };
		// Clean up the file name before appending it to the path
		EC_H(SanitizeFileName(szTemp, _countof(szTemp), lpDisplayName->Value.LPSZ, _countof(szTemp)));

#ifdef UNICODE
		EC_H(StringCchPrintfW(m_szFolderPathRoot,_countof(m_szFolderPathRoot),
			L"%s\\%ws", // STRING_OK
			m_szMailboxTablePathRoot,szTemp));
#else
		EC_H(StringCchPrintfW(m_szFolderPathRoot, _countof(m_szFolderPathRoot),
			L"%s\\%hs", // STRING_OK
			m_szMailboxTablePathRoot, szTemp));
#endif

		// suppress any error here since the folder may already exist
		WC_B(CreateDirectoryW(m_szFolderPathRoot, NULL));
		hRes = S_OK;
	}

	OutputToFile(m_fMailboxTable, _T("</mailbox>\n"));
} // CDumpStore::DoMailboxTablePerRowWork

void CDumpStore::EndMailboxTableWork()
{
	if (m_bOutputList) return;
	if (m_fMailboxTable)
	{
		OutputToFile(m_fMailboxTable, _T("</mailboxtable>\n"));
		CloseFile(m_fMailboxTable);
	}
	m_fMailboxTable = NULL;
} // CDumpStore::EndMailboxTableWork(

void CDumpStore::BeginStoreWork()
{
} // CDumpStore::BeginStoreWork

void CDumpStore::EndStoreWork()
{
} // CDumpStore::EndStoreWork

void CDumpStore::BeginFolderWork()
{
	HRESULT hRes = S_OK;
	WCHAR	szFolderPath[MAX_PATH];
#ifdef UNICODE
	WC_H(StringCchPrintfW(szFolderPath,_countof(szFolderPath),
		L"%s%ws", // STRING_OK
		m_szFolderPathRoot,m_szFolderOffset));
#else
	WC_H(StringCchPrintfW(szFolderPath, _countof(szFolderPath),
		L"%s%hs", // STRING_OK
		m_szFolderPathRoot, m_szFolderOffset));
#endif

	WC_H(CopyStringW(&m_szFolderPath, szFolderPath, NULL));

	// We've done all the setup we need. If we're just outputting a list, we don't need to do the rest
	if (m_bOutputList) return;

	WC_B(CreateDirectoryW(m_szFolderPath, NULL));
	hRes = S_OK; // ignore the error - the directory may exist already

	WCHAR	szFolderPropsFile[MAX_PATH]; // Holds file/path name for folder props

	// Dump the folder props to a file
	WC_H(StringCchPrintfW(szFolderPropsFile, _countof(szFolderPropsFile),
		L"%sFOLDER_PROPS.xml", // STRING_OK
		m_szFolderPath));
	m_fFolderProps = MyOpenFile(szFolderPropsFile, true);
	if (!m_fFolderProps) return;

	OutputToFile(m_fFolderProps, g_szXMLHeader);
	OutputToFile(m_fFolderProps, _T("<folderprops>\n"));

	LPSPropValue	lpAllProps = NULL;
	ULONG			cValues = 0L;

	WC_H_GETPROPS(GetPropsNULL(m_lpFolder,
		fMapiUnicode,
		&cValues,
		&lpAllProps));
	if (FAILED(hRes))
	{
		OutputToFilef(m_fFolderProps, _T("<properties error=\"0x%08X\" />\n"), hRes);
	}
	else if (lpAllProps)
	{
		OutputToFile(m_fFolderProps, _T("<properties listtype=\"summary\">\n"));

		OutputPropertiesToFile(m_fFolderProps, cValues, lpAllProps, m_lpFolder, m_bRetryStreamProps);

		OutputToFile(m_fFolderProps, _T("</properties>\n"));

		MAPIFreeBuffer(lpAllProps);
	}

	OutputToFile(m_fFolderProps, _T("<HierarchyTable>\n"));
} // CDumpStore::BeginFolderWork

void CDumpStore::DoFolderPerHierarchyTableRowWork(_In_ LPSRow lpSRow)
{
	if (m_bOutputList) return;
	if (!m_fFolderProps || !lpSRow) return;
	OutputToFile(m_fFolderProps, _T("<row>\n"));
	OutputSRowToFile(m_fFolderProps, lpSRow, m_lpMDB);
	OutputToFile(m_fFolderProps, _T("</row>\n"));
} // CDumpStore::DoFolderPerHierarchyTableRowWork

void CDumpStore::EndFolderWork()
{
	if (m_bOutputList) return;
	if (m_fFolderProps)
	{
		OutputToFile(m_fFolderProps, _T("</HierarchyTable>\n"));
		OutputToFile(m_fFolderProps, _T("</folderprops>\n"));
		CloseFile(m_fFolderProps);
	}
	m_fFolderProps = NULL;
} // CDumpStore::EndFolderWork

void CDumpStore::BeginContentsTableWork(ULONG ulFlags, ULONG ulCountRows)
{
	if (!m_szFolderPathRoot) return;
	if (m_bOutputList)
	{
		_tprintf(_T("Subject, Message Class, Filename\n"));
		return;
	}

	HRESULT hRes = S_OK;
	WCHAR	szContentsTableFile[MAX_PATH]; // Holds file/path name for contents table output

	WC_H(StringCchPrintfW(szContentsTableFile, _countof(szContentsTableFile),
		(ulFlags & MAPI_ASSOCIATED) ? L"%sASSOCIATED_CONTENTS_TABLE.xml" : L"%sCONTENTS_TABLE.xml", // STRING_OK
		m_szFolderPath));
	m_fFolderContents = MyOpenFile(szContentsTableFile, true);
	if (m_fFolderContents)
	{
		OutputToFile(m_fFolderContents, g_szXMLHeader);
		OutputToFilef(m_fFolderContents, _T("<ContentsTable TableType=\"%s\" messagecount=\"0x%08X\">\n"),
			(ulFlags & MAPI_ASSOCIATED) ? _T("Associated Contents") : _T("Contents"), // STRING_OK
			ulCountRows);
	}
} // CDumpStore::BeginContentsTableWork

// Outputs a single message's details to the screen, so as to produce a list of messages
void OutputMessageList(
	_In_ LPSRow lpSRow,
	_In_z_ LPWSTR szFolderPath,
	bool bOutputMSG)
{
	HRESULT hRes = S_OK;

	if (!lpSRow || !szFolderPath) return;

	WCHAR szFileName[MAX_PATH] = { 0 };

	LPCWSTR szSubj = NULL;
	LPSBinary lpRecordKey = NULL;
	LPSPropValue lpTemp = NULL;
	LPSPropValue lpMessageClass = NULL;
	LPWSTR szTemp = NULL;
	LPCWSTR szExt = L".xml"; // STRING_OK
	if (bOutputMSG) szExt = L".msg"; // STRING_OK

	// Get required properties from the message
	lpTemp = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_SUBJECT_W);
	if (lpTemp && CheckStringProp(lpTemp, PT_UNICODE))
	{
		szSubj = lpTemp->Value.lpszW;
	}
	else
	{
		lpTemp = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_SUBJECT_A);
		if (lpTemp && CheckStringProp(lpTemp, PT_STRING8))
		{
			WC_H(AnsiToUnicode(lpTemp->Value.lpszA, &szTemp));
			if (SUCCEEDED(hRes)) szSubj = szTemp;
			hRes = S_OK;
		}
	}
	lpTemp = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, PR_RECORD_KEY);
	if (lpTemp && PR_RECORD_KEY == lpTemp->ulPropTag)
	{
		lpRecordKey = &lpTemp->Value.bin;
	}
	lpMessageClass = PpropFindProp(lpSRow->lpProps, lpSRow->cValues, CHANGE_PROP_TYPE(PR_MESSAGE_CLASS, PT_UNSPECIFIED));

	WC_H(BuildFileNameAndPath(szFileName, _countof(szFileName), szExt, 4, szSubj, lpRecordKey, szFolderPath));

	_tprintf(_T("\"%ws\""), szSubj ? szSubj : L"");
	if (lpMessageClass)
	{
		if (PT_STRING8 == PROP_TYPE(lpMessageClass->ulPropTag))
		{
			_tprintf(_T(",\"%hs\""), lpMessageClass->Value.lpszA ? lpMessageClass->Value.lpszA : "");
		}
		else if (PT_UNICODE == PROP_TYPE(lpMessageClass->ulPropTag))
		{
			_tprintf(_T(",\"%ws\""), lpMessageClass->Value.lpszW ? lpMessageClass->Value.lpszW : L"");
		}
	}
	_tprintf(_T(",\"%ws\"\n"), szFileName);
	delete[] szTemp;
} // OutputMessageList

bool CDumpStore::DoContentsTablePerRowWork(_In_ LPSRow lpSRow, ULONG ulCurRow)
{
	if (m_bOutputList)
	{
		OutputMessageList(lpSRow, m_szFolderPath, m_bOutputMSG);
		return false;
	}
	if (!m_fFolderContents || !m_lpFolder) return true;

	OutputToFilef(m_fFolderContents, _T("<message num=\"0x%08X\">\n"), ulCurRow);

	OutputSRowToFile(m_fFolderContents, lpSRow, m_lpFolder);

	OutputToFile(m_fFolderContents, _T("</message>\n"));
	return true;
} // CDumpStore::DoContentsTablePerRowWork

void CDumpStore::EndContentsTableWork()
{
	if (m_bOutputList) return;
	if (m_fFolderContents)
	{
		OutputToFile(m_fFolderContents, _T("</ContentsTable>\n"));
		CloseFile(m_fFolderContents);
	}
	m_fFolderContents = NULL;
} // CDumpStore::EndContentsTableWork

// TODO: This fails in unicode builds since PR_RTF_COMPRESSED is always ascii.
void OutputBody(_In_ FILE* fMessageProps, _In_ LPMESSAGE lpMessage, ULONG ulBodyTag, _In_z_ LPCTSTR szBodyName, bool bWrapEx, ULONG ulCPID)
{
	HRESULT hRes = S_OK;
	LPSTREAM lpStream = NULL;
	LPSTREAM lpRTFUncompressed = NULL;
	LPSTREAM lpOutputStream = NULL;

	WC_MAPI(lpMessage->OpenProperty(
		ulBodyTag,
		&IID_IStream,
		STGM_READ,
		NULL,
		(LPUNKNOWN *)&lpStream));
	// The only error we suppress is MAPI_E_NOT_FOUND, so if a body type isn't in the output, it wasn't on the message
	if (MAPI_E_NOT_FOUND != hRes)
	{
		OutputToFilef(fMessageProps, _T("<body property=\"%s\""), szBodyName);
		if (!lpStream)
		{
			OutputToFilef(fMessageProps, _T(" error=\"0x%08X\">\n"), hRes);
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
					wstring szFlags = InterpretFlags(flagStreamFlag, ulStreamFlags);
					OutputToFilef(fMessageProps, _T(" ulStreamFlags = \"0x%08X\" szStreamFlags= \"%ws\""), ulStreamFlags, szFlags.c_str());
					OutputToFilef(fMessageProps, _T(" CodePageIn = \"%u\" CodePageOut = \"%d\""), ulCPID, CP_ACP);
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
						NULL));
				}
				if (!lpRTFUncompressed || FAILED(hRes))
				{
					OutputToFilef(fMessageProps, _T(" rtfWrapError=\"0x%08X\""), hRes);
				}
				else
				{
					lpOutputStream = lpRTFUncompressed;
				}
			}

			OutputToFile(fMessageProps, _T(">\n"));
			if (lpOutputStream)
			{
				OutputCDataOpen(DBGNoDebug, fMessageProps);
				OutputStreamToFile(fMessageProps, lpOutputStream);
				OutputCDataClose(DBGNoDebug, fMessageProps);
			}
		}
		OutputToFile(fMessageProps, _T("</body>\n"));
	}
	if (lpRTFUncompressed) lpRTFUncompressed->Release();
	if (lpStream) lpStream->Release();
} // OutputBody

void OutputMessageXML(
	_In_ LPMESSAGE lpMessage,
	_In_ LPVOID lpParentMessageData,
	_In_z_ LPWSTR szMessageFileName,
	_In_z_ LPWSTR szFolderPath,
	bool bRetryStreamProps,
	_Deref_out_ LPVOID* lpData)
{
	if (!lpMessage || !lpData) return;

	HRESULT hRes = S_OK;

	*lpData = (LPVOID) new(MessageData);
	if (!*lpData) return;

	LPMESSAGEDATA lpMsgData = (LPMESSAGEDATA)*lpData;

	LPSPropValue lpAllProps = NULL;
	LPSPropValue lpTemp = NULL;
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

	lpMsgData->szFilePath[0] = NULL;
	// If we've got a parent message, we're an attachment - use attachment filename logic
	if (lpParentMessageData)
	{
		// Take the parent filename/path, remove any extension, and append -Attach.xml
		// Should we append something for attachment number?
		size_t  cchLen = 0;

		// Copy the source string over
		WC_H(StringCchCopyW(lpMsgData->szFilePath, _countof(lpMsgData->szFilePath), ((LPMESSAGEDATA)lpParentMessageData)->szFilePath));

		// Remove any extension
		WC_H(StringCchLengthW(lpMsgData->szFilePath, _countof(lpMsgData->szFilePath), &cchLen));
		while (cchLen > 0 && lpMsgData->szFilePath[cchLen] != L'.') cchLen--;
		if (lpMsgData->szFilePath[cchLen] == L'.') lpMsgData->szFilePath[cchLen] = L'\0';

		// build a string for appending
		WCHAR szNewExt[MAX_PATH];
		WC_H(StringCchPrintfW(szNewExt, _countof(szNewExt), L"-Attach%u.xml", ((LPMESSAGEDATA)lpParentMessageData)->ulCurAttNum)); // STRING_OK

		// append our string
		WC_H(StringCchCatW(lpMsgData->szFilePath, _countof(lpMsgData->szFilePath), szNewExt));

		OutputToFilef(((LPMESSAGEDATA)lpParentMessageData)->fMessageProps, _T("<embeddedmessage path=\"%ws\"/>\n"), lpMsgData->szFilePath);
	}
	else if (szMessageFileName[0]) // if we've got a file name, use it
	{
		WC_H(StringCchCopyW(lpMsgData->szFilePath, _countof(lpMsgData->szFilePath), szMessageFileName));
	}
	else
	{
		LPCWSTR szSubj = NULL; // BuildFileNameAndPath will substitute a subject if we don't find one
		LPWSTR szTemp = NULL;
		LPSBinary lpRecordKey = NULL;

		lpTemp = PpropFindProp(lpAllProps, cValues, PR_SUBJECT_W);
		if (lpTemp && CheckStringProp(lpTemp, PT_UNICODE))
		{
			szSubj = lpTemp->Value.lpszW;
		}
		else
		{
			lpTemp = PpropFindProp(lpAllProps, cValues, PR_SUBJECT_A);
			if (lpTemp && CheckStringProp(lpTemp, PT_STRING8))
			{
				WC_H(AnsiToUnicode(lpTemp->Value.lpszA, &szTemp));
				if (SUCCEEDED(hRes)) szSubj = szTemp;
				hRes = S_OK;
			}
		}
		lpTemp = PpropFindProp(lpAllProps, cValues, PR_RECORD_KEY);
		if (lpTemp && PR_RECORD_KEY == lpTemp->ulPropTag)
		{
			lpRecordKey = &lpTemp->Value.bin;
		}
		WC_H(BuildFileNameAndPath(lpMsgData->szFilePath, _countof(lpMsgData->szFilePath), L".xml", 4, szSubj, lpRecordKey, szFolderPath)); // STRING_OK
		delete[] szTemp;
	}

	DebugPrint(DBGGeneric, _T("OutputMessagePropertiesToFile: Saving %p to \"%ws\"\n"), lpMessage, lpMsgData->szFilePath);
	lpMsgData->fMessageProps = MyOpenFile(lpMsgData->szFilePath, true);

	if (lpMsgData->fMessageProps)
	{
		OutputToFile(lpMsgData->fMessageProps, g_szXMLHeader);
		OutputToFile(lpMsgData->fMessageProps, _T("<message>\n"));
		if (lpAllProps)
		{
			OutputToFile(lpMsgData->fMessageProps, _T("<properties listtype=\"summary\">\n"));
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

			ULONG i = 0;
			for (i = 0; i < NUMPROPS; i++)
			{
				lpTemp = PpropFindProp(lpAllProps, cValues, sptCols.aulPropTag[i]);
				if (lpTemp)
				{
					OutputPropertyToFile(lpMsgData->fMessageProps, lpTemp, lpMessage, bRetryStreamProps);
				}
			}

			OutputToFile(lpMsgData->fMessageProps, _T("</properties>\n"));
		}

		// Log Body
		OutputBody(lpMsgData->fMessageProps, lpMessage, PR_BODY, _T("PR_BODY"), false, NULL);
		OutputBody(lpMsgData->fMessageProps, lpMessage, PR_BODY_HTML, _T("PR_BODY_HTML"), false, NULL);
		OutputBody(lpMsgData->fMessageProps, lpMessage, PR_RTF_COMPRESSED, _T("PR_RTF_COMPRESSED"), false, NULL);

		ULONG ulInCodePage = CP_ACP; // picking CP_ACP as our default
		lpTemp = PpropFindProp(lpAllProps, cValues, PR_INTERNET_CPID);
		if (lpTemp && PR_INTERNET_CPID == lpTemp->ulPropTag)
		{
			ulInCodePage = lpTemp->Value.l;
		}

		OutputBody(lpMsgData->fMessageProps, lpMessage, PR_RTF_COMPRESSED, _T("WrapCompressedRTFEx best body"), true, ulInCodePage);

		if (lpAllProps)
		{
			OutputToFile(lpMsgData->fMessageProps, _T("<properties listtype=\"FullPropList\">\n"));

			OutputPropertiesToFile(lpMsgData->fMessageProps, cValues, lpAllProps, lpMessage, bRetryStreamProps);

			OutputToFile(lpMsgData->fMessageProps, _T("</properties>\n"));
		}
	}

	MAPIFreeBuffer(lpAllProps);
}

void OutputMessageMSG(
	_In_ LPMESSAGE lpMessage,
	_In_z_ LPWSTR szFolderPath)
{
	HRESULT hRes = S_OK;

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

	if (!lpMessage || !szFolderPath) return;

	DebugPrint(DBGGeneric, _T("OutputMessageMSG: Saving message to \"%ws\"\n"), szFolderPath);

	WCHAR szFileName[MAX_PATH] = { 0 };

	LPCWSTR szSubj = NULL;

	ULONG cProps = 0;
	LPSPropValue lpsProps = NULL;
	LPSBinary lpRecordKey = NULL;


	// Get required properties from the message
	EC_H_GETPROPS(lpMessage->GetProps(
		(LPSPropTagArray)&msgProps,
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
	WC_H(BuildFileNameAndPath(szFileName, _countof(szFileName), L".msg", 4, szSubj, lpRecordKey, szFolderPath)); // STRING_OK

	DebugPrint(DBGGeneric, _T("Saving to = \"%ws\"\n"), szFileName);

	WC_H(SaveToMSG(
		lpMessage,
		szFileName,
		fMapiUnicode ? true : false,
		NULL,
		false));
} // OutputMessageMSG

bool CDumpStore::BeginMessageWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpParentMessageData, _Deref_out_opt_ LPVOID* lpData)
{
	if (lpData) *lpData = NULL;
	if (lpParentMessageData && !m_bOutputAttachments) return false;
	if (m_bOutputList) return false;

	if (m_bOutputMSG)
	{
		OutputMessageMSG(lpMessage, m_szFolderPath);
		return false; // no more work necessary
	}
	else
	{
		OutputMessageXML(lpMessage, lpParentMessageData, m_szMessageFileName, m_szFolderPath, m_bRetryStreamProps, lpData);
		return true;
	}
}

bool CDumpStore::BeginRecipientWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
{
	if (!lpData) return false;
	if (m_bOutputMSG) return false; // When outputting message files, no recipient work is needed
	if (m_bOutputList) return false;
	OutputToFile(((LPMESSAGEDATA)lpData)->fMessageProps, _T("<recipients>\n"));
	return true;
} // CDumpStore::BeginRecipientWork

void CDumpStore::DoMessagePerRecipientWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData, _In_ LPSRow lpSRow, ULONG ulCurRow)
{
	if (!lpMessage || !lpData || !lpSRow) return;
	if (m_bOutputMSG) return; // When outputting message files, no recipient work is needed
	if (m_bOutputList) return;

	LPMESSAGEDATA lpMsgData = (LPMESSAGEDATA)lpData;

	OutputToFilef(lpMsgData->fMessageProps, _T("<recipient num=\"0x%08X\">\n"), ulCurRow);

	OutputSRowToFile(lpMsgData->fMessageProps, lpSRow, lpMessage);

	OutputToFile(lpMsgData->fMessageProps, _T("</recipient>\n"));
} // CDumpStore::DoMessagePerRecipientWork

void CDumpStore::EndRecipientWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
{
	if (!lpData) return;
	if (m_bOutputMSG) return; // When outputting message files, no recipient work is needed
	if (m_bOutputList) return;
	OutputToFile(((LPMESSAGEDATA)lpData)->fMessageProps, _T("</recipients>\n"));
} // CDumpStore::EndRecipientWork

bool CDumpStore::BeginAttachmentWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
{
	if (!lpData) return false;
	if (m_bOutputMSG) return false; // When outputting message files, no attachment work is needed
	if (m_bOutputList) return false;
	OutputToFile(((LPMESSAGEDATA)lpData)->fMessageProps, _T("<attachments>\n"));
	return true;
} // CDumpStore::BeginAttachmentWork

void CDumpStore::DoMessagePerAttachmentWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData, _In_ LPSRow lpSRow, _In_ LPATTACH lpAttach, ULONG ulCurRow)
{
	if (!lpMessage || !lpData || !lpSRow) return;
	if (m_bOutputMSG) return; // When outputting message files, no attachment work is needed
	if (m_bOutputList) return;
	HRESULT hRes = S_OK;

	LPMESSAGEDATA lpMsgData = (LPMESSAGEDATA)lpData;

	lpMsgData->ulCurAttNum = ulCurRow; // set this so we can pull it if this is an embedded message

	LPSPropValue lpAttachName = NULL;

	hRes = S_OK;

	OutputToFilef(lpMsgData->fMessageProps, _T("<attachment num=\"0x%08X\" filename=\""), ulCurRow);

	lpAttachName = PpropFindProp(
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
		OutputToFile(lpMsgData->fMessageProps, lpAttachName->Value.LPSZ);
	else
		OutputToFile(lpMsgData->fMessageProps, _T("PR_ATTACH_FILENAME not found"));

	OutputToFile(lpMsgData->fMessageProps, _T("\">\n"));

	OutputToFile(lpMsgData->fMessageProps, _T("\t<tableprops>\n"));
	OutputSRowToFile(lpMsgData->fMessageProps, lpSRow, lpMessage);

	OutputToFile(lpMsgData->fMessageProps, _T("\t</tableprops>\n"));

	if (lpAttach)
	{
		ULONG			ulAllProps = 0;
		LPSPropValue	lpAllProps = NULL;
		// Let's get all props from the message and dump them.
		WC_H_GETPROPS(GetPropsNULL(lpAttach,
			fMapiUnicode,
			&ulAllProps,
			&lpAllProps));
		if (lpAllProps)
		{
			// We've reported any error - mask it now
			hRes = S_OK;

			OutputToFile(lpMsgData->fMessageProps, _T("\t<getprops>\n"));
			OutputPropertiesToFile(lpMsgData->fMessageProps, ulAllProps, lpAllProps, lpAttach, m_bRetryStreamProps);
			OutputToFile(lpMsgData->fMessageProps, _T("\t</getprops>\n"));
		}

		MAPIFreeBuffer(lpAllProps);
		lpAllProps = NULL;
	}

	OutputToFile(lpMsgData->fMessageProps, _T("</attachment>\n"));
}

void CDumpStore::EndAttachmentWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
{
	if (!lpData) return;
	if (m_bOutputMSG) return; // When outputting message files, no attachment work is needed
	if (m_bOutputList) return;
	OutputToFile(((LPMESSAGEDATA)lpData)->fMessageProps, _T("</attachments>\n"));
} // CDumpStore::EndAttachmentWork

void CDumpStore::EndMessageWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
{
	if (m_bOutputMSG) return; // When outputting message files, no end message work is needed
	if (m_bOutputList) return;
	LPMESSAGEDATA lpMsgData = (LPMESSAGEDATA)lpData;

	if (lpMsgData->fMessageProps)
	{
		OutputToFile(lpMsgData->fMessageProps, _T("</message>\n"));
		CloseFile(lpMsgData->fMessageProps);
	}
	delete lpMsgData;
} // CDumpStore::EndMessageWork