//  MODULE:   DumpStore.cpp
//
//  PURPOSE:  Contains routines used in dumping the contents of the Exchange store
//            in to a log file

#include "stdafx.h"
#include "dumpstore.h"
#include "MAPIFunctions.h"
#include "file.h"
#include "InterpretProp.h"
#include "InterpretProp2.h"
#include "ImportProcs.h"
#include "ExtraPropTags.h"
#include "ColumnTags.h"
#include "PropTagArray.h"
#include "SmartView.h"

void OutputPropertyToFile(_In_ FILE* fFile, _In_ LPSPropValue lpProp, _In_ LPMAPIPROP lpObj)
{
	if (!lpProp) return;

	CString PropType;

	OutputToFilef(fFile,_T("\t<property tag = \"0x%08X\" type = \"%s\">\n"),lpProp->ulPropTag,(LPCTSTR) TypeToString(lpProp->ulPropTag));

	CString PropString;
	CString AltPropString;
	LPTSTR szExactMatches = NULL;
	LPTSTR szPartialMatches = NULL;
	LPTSTR szSmartView = NULL;
	LPTSTR szNamedPropName = NULL;
	LPTSTR szNamedPropGUID = NULL;

	InterpretProp(
		lpProp,
		lpProp->ulPropTag,
		lpObj,
		NULL,
		NULL,
		false,
		&szExactMatches, // Built from ulPropTag & bIsAB
		&szPartialMatches, // Built from ulPropTag & bIsAB
		&PropType,
		NULL,
		&PropString,
		&AltPropString,
		&szNamedPropName, // Built from lpProp & lpMAPIProp
		&szNamedPropGUID, // Built from lpProp & lpMAPIProp
		NULL);

	InterpretPropSmartView(
		lpProp,
		lpObj,
		NULL,
		NULL,
		false,
		&szSmartView);

	OutputXMLValueToFile(fFile,PropXMLNames[pcPROPEXACTNAMES].uidName,szExactMatches,2);
	OutputXMLValueToFile(fFile,PropXMLNames[pcPROPPARTIALNAMES].uidName,szPartialMatches,2);
	OutputXMLValueToFile(fFile,PropXMLNames[pcPROPNAMEDIID].uidName, szNamedPropGUID,2);
	OutputXMLValueToFile(fFile,PropXMLNames[pcPROPNAMEDNAME].uidName, szNamedPropName,2);

	switch(PROP_TYPE(lpProp->ulPropTag))
	{
	case PT_STRING8:
	case PT_UNICODE:
	case PT_MV_STRING8:
	case PT_MV_UNICODE:
		{
			OutputXMLCDataValueToFile(fFile,PropXMLNames[pcPROPVAL].uidName,(LPCTSTR) PropString,2);
			OutputXMLValueToFile(fFile,PropXMLNames[pcPROPVALALT].uidName,(LPCTSTR) AltPropString,2);
			break;
		}
	case PT_BINARY:
	case PT_MV_BINARY:
		{
			OutputXMLValueToFile(fFile,PropXMLNames[pcPROPVAL].uidName,(LPCTSTR) PropString,2);
			OutputXMLCDataValueToFile(fFile,PropXMLNames[pcPROPVALALT].uidName,(LPCTSTR) AltPropString,2);
			break;
		}
	default:
		{
			OutputXMLValueToFile(fFile,PropXMLNames[pcPROPVAL].uidName,(LPCTSTR) PropString,2);
			OutputXMLValueToFile(fFile,PropXMLNames[pcPROPVALALT].uidName,(LPCTSTR) AltPropString,2);
			break;
		}
	}

	if (szSmartView) OutputXMLCDataValueToFile(fFile,PropXMLNames[pcPROPSMARTVIEW].uidName,szSmartView,2);
	OutputToFile(fFile,_T("\t</property>\n"));

	delete[] szPartialMatches;
	delete[] szExactMatches;
	FreeNameIDStrings(szNamedPropName, szNamedPropGUID, NULL);
	delete[] szSmartView;
} // OutputPropertyToFile

void OutputPropertiesToFile(_In_ FILE* fFile, ULONG cProps, _In_count_(cProps) LPSPropValue lpProps, _In_ LPMAPIPROP lpObj)
{
	if (cProps && !lpProps) return;

	// sort the list first
	// insertion sort on lpProps
	ULONG iUnsorted = 0;
	ULONG iLoc = 0;
	for (iUnsorted = 1; iUnsorted < cProps; iUnsorted++)
	{
		SPropValue NextItem = lpProps[iUnsorted];
		for (iLoc = iUnsorted; iLoc > 0; iLoc--)
		{
			if (lpProps[iLoc-1].ulPropTag < NextItem.ulPropTag) break;
			lpProps[iLoc] = lpProps[iLoc-1];
		}
		lpProps[iLoc] = NextItem;
	}

	ULONG i = 0;
	for (i = 0; i < cProps; i++)
	{
		OutputPropertyToFile(fFile, &lpProps[i], lpObj);
	}
} // OutputPropertiesToFile

void OutputSRowToFile(_In_ FILE* fFile, _In_ LPSRow lpSRow, _In_ LPMAPIPROP lpObj)
{
	if (!lpSRow) return;

	if (lpSRow->cValues && !lpSRow->lpProps) return;

	OutputPropertiesToFile(fFile, lpSRow->cValues, lpSRow->lpProps,lpObj);
} // OutputSRowToFile

CDumpStore::CDumpStore()
{
	m_szMessageFileName[0] = L'\0';
	m_szMailboxTablePathRoot[0] = L'\0';
	m_szFolderPathRoot[0] = L'\0';

	m_szFolderPath = NULL;

	m_fFolderProps = NULL;
	m_fFolderContents = NULL;
	m_fMailboxTable = NULL;
} // CDumpStore::CDumpStore

CDumpStore::~CDumpStore()
{
	if (m_fFolderProps)		CloseFile(m_fFolderProps);
	if (m_fFolderContents)	CloseFile(m_fFolderContents);
	if (m_fMailboxTable)	CloseFile(m_fMailboxTable);
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

// --------------------------------------------------------------------------------- //

void CDumpStore::BeginMailboxTableWork(_In_z_ LPCTSTR szExchangeServerName)
{
	HRESULT hRes = S_OK;
	WCHAR	szTableContentsFile[MAX_PATH];
	WC_H(StringCchPrintfW(szTableContentsFile,_countof(szTableContentsFile),
		L"%s\\MAILBOX_TABLE.xml", // STRING_OK
		m_szMailboxTablePathRoot));
	m_fMailboxTable = OpenFile(szTableContentsFile,true);
	if (m_fMailboxTable)
	{
		OutputToFile(m_fMailboxTable,_T("<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?>\n"));
		OutputToFilef(m_fMailboxTable,_T("<mailboxtable server=\"%s\">\n"),szExchangeServerName);
	}
} // CDumpStore::BeginMailboxTableWork

void CDumpStore::DoMailboxTablePerRowWork(_In_ LPMDB lpMDB, _In_ LPSRow lpSRow, ULONG /*ulCurRow*/)
{
	if (!lpSRow || !m_fMailboxTable) return;
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

	OutputToFile(m_fMailboxTable,_T("<mailbox prdisplayname=\""));
	if (CheckStringProp(lpDisplayName,PT_TSTRING))
	{
		OutputToFile(m_fMailboxTable,lpDisplayName->Value.LPSZ);
	}
	OutputToFile(m_fMailboxTable,_T("\" premailaddress=\""));
	if (!CheckStringProp(lpEmailAddress,PT_TSTRING))
	{
		OutputToFile(m_fMailboxTable,lpEmailAddress->Value.LPSZ);
	}
	OutputToFile(m_fMailboxTable,_T("\">\n"));

	OutputSRowToFile(m_fMailboxTable,lpSRow,lpMDB);

	// build a path for our store's folder output:
	if (CheckStringProp(lpEmailAddress,PT_TSTRING) && CheckStringProp(lpDisplayName,PT_TSTRING))
	{
		TCHAR szTemp[MAX_PATH/2];
		// Clean up the file name before appending it to the path
		EC_H(SanitizeFileName(szTemp,_countof(szTemp),lpDisplayName->Value.LPSZ,_countof(szTemp)));

#ifdef UNICODE
		EC_H(StringCchPrintfW(m_szFolderPathRoot,_countof(m_szFolderPathRoot),
			L"%s\\%ws", // STRING_OK
			m_szMailboxTablePathRoot,szTemp));
#else
		EC_H(StringCchPrintfW(m_szFolderPathRoot,_countof(m_szFolderPathRoot),
			L"%s\\%hs", // STRING_OK
			m_szMailboxTablePathRoot,szTemp));
#endif

		// suppress any error here since the folder may already exist
		WC_B(CreateDirectoryW(m_szFolderPathRoot,NULL));
		hRes = S_OK;
	}

	OutputToFile(m_fMailboxTable,_T("</mailbox>\n"));
} // CDumpStore::DoMailboxTablePerRowWork

void CDumpStore::EndMailboxTableWork()
{
	if (m_fMailboxTable)
	{
		OutputToFile(m_fMailboxTable,_T("</mailboxtable>\n"));
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
	WC_H(StringCchPrintfW(szFolderPath,_countof(szFolderPath),
		L"%s%hs", // STRING_OK
		m_szFolderPathRoot,m_szFolderOffset));
#endif

	WC_H(CopyStringW(&m_szFolderPath,szFolderPath,NULL));

	WC_B(CreateDirectoryW(m_szFolderPath,NULL));
	hRes = S_OK; // ignore the error - the directory may exist already

	WCHAR	szFolderPropsFile[MAX_PATH]; // Holds file/path name for folder props

	// Dump the folder props to a file
	WC_H(StringCchPrintfW(szFolderPropsFile,_countof(szFolderPropsFile),
		L"%sFOLDER_PROPS.xml", // STRING_OK
		m_szFolderPath));
	m_fFolderProps = OpenFile(szFolderPropsFile,true);
	if (!m_fFolderProps) return;

	OutputToFile(m_fFolderProps,_T("<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?>\n"));
	OutputToFile(m_fFolderProps,_T("<folderprops>\n"));

	LPSPropValue	lpAllProps = NULL;
	ULONG			cValues = 0L;

	hRes = GetPropsNULL(m_lpFolder,
		fMapiUnicode,
		&cValues,
		&lpAllProps);
	WARNHRESMSG(hRes,IDS_DUMPFOLDERGETPROPFAILED);
	if (FAILED(hRes))
	{
		OutputToFilef(m_fFolderProps,_T("<properties error=\"0x%08X\" />\n"),hRes);
	}
	else if (lpAllProps)
	{
		OutputToFile(m_fFolderProps,_T("<properties listtype=\"summary\">\n"));

		OutputPropertiesToFile(m_fFolderProps, cValues, lpAllProps, m_lpFolder);

		OutputToFile(m_fFolderProps,_T("</properties>\n"));

		MAPIFreeBuffer(lpAllProps);
	}

	OutputToFile(m_fFolderProps,_T("<HierarchyTable>\n"));
} // CDumpStore::BeginFolderWork

void CDumpStore::DoFolderPerHierarchyTableRowWork(_In_ LPSRow lpSRow)
{
	if (!m_fFolderProps || !lpSRow) return;
	OutputToFile(m_fFolderProps,_T("<row>\n"));
	OutputSRowToFile(m_fFolderProps,lpSRow,m_lpMDB);
	OutputToFile(m_fFolderProps,_T("</row>\n"));
} // CDumpStore::DoFolderPerHierarchyTableRowWork

void CDumpStore::EndFolderWork()
{
	if (m_fFolderProps)
	{
		OutputToFile(m_fFolderProps,_T("</HierarchyTable>\n"));
		OutputToFile(m_fFolderProps,_T("</folderprops>\n"));
		CloseFile(m_fFolderProps);
	}
	m_fFolderProps = NULL;
} // CDumpStore::EndFolderWork

void CDumpStore::BeginContentsTableWork(ULONG ulFlags, ULONG ulCountRows)
{
	if (!m_szFolderPathRoot) return;

	HRESULT hRes = S_OK;
	WCHAR	szContentsTableFile[MAX_PATH]; // Holds file/path name for contents table output

	WC_H(StringCchPrintfW(szContentsTableFile,_countof(szContentsTableFile),
		(ulFlags & MAPI_ASSOCIATED)?L"%sASSOCIATED_CONTENTS_TABLE.xml":L"%sCONTENTS_TABLE.xml", // STRING_OK
		m_szFolderPath));
	m_fFolderContents = OpenFile(szContentsTableFile,true);
	if (m_fFolderContents)
	{
		OutputToFile(m_fFolderContents,_T("<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?>\n"));
		OutputToFilef(m_fFolderContents,_T("<ContentsTable TableType=\"%s\" messagecount=\"0x%08X\">\n"),
			(ulFlags & MAPI_ASSOCIATED)?_T("Associated Contents"):_T("Contents"), // STRING_OK
			ulCountRows);
	}
} // CDumpStore::BeginContentsTableWork

void CDumpStore::DoContentsTablePerRowWork(_In_ LPSRow lpSRow, ULONG ulCurRow)
{
	if (!m_fFolderContents || !m_lpFolder) return;

	OutputToFilef(m_fFolderContents,_T("<message num=\"0x%08X\">\n"),ulCurRow);

	OutputSRowToFile(m_fFolderContents,lpSRow,m_lpFolder);

	OutputToFile(m_fFolderContents,_T("</message>\n"));
} // CDumpStore::DoContentsTablePerRowWork

void CDumpStore::EndContentsTableWork()
{
	if (m_fFolderContents)
	{
		OutputToFile(m_fFolderContents,_T("</ContentsTable>\n"));
		CloseFile(m_fFolderContents);
	}
	m_fFolderContents = NULL;
} // CDumpStore::EndContentsTableWork

void CDumpStore::BeginMessageWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpParentMessageData, _Deref_out_ LPVOID* lpData)
{
	if (!lpMessage || !lpData) return;

	HRESULT hRes = S_OK;

	*lpData = (LPVOID) new(MessageData);
	if (!*lpData) return;

	LPMESSAGEDATA lpMsgData = (LPMESSAGEDATA) *lpData;

	enum {ePR_MESSAGE_CLASS,
		ePR_MESSAGE_DELIVERY_TIME,
		ePR_SUBJECT_W,
		ePR_SENDER_ADDRTYPE,
		ePR_SENDER_EMAIL_ADDRESS,
		ePR_INTERNET_CPID,
		ePR_ENTRYID,
		ePR_SEARCH_KEY,
		NUM_COLS};
	// These tags represent the message information we would like to pick up
	static SizedSPropTagArray(NUM_COLS,sptCols) = { NUM_COLS,
		PR_MESSAGE_CLASS,
		PR_MESSAGE_DELIVERY_TIME,
		PR_SUBJECT_W,
		PR_SENDER_ADDRTYPE,
		PR_SENDER_EMAIL_ADDRESS,
		PR_INTERNET_CPID,
		PR_ENTRYID,
		PR_SEARCH_KEY
	};

	LPSPropValue	lpPropsMsg = NULL;
	ULONG			cValues = 0L;

	WC_H_GETPROPS(lpMessage->GetProps(
		(LPSPropTagArray)&sptCols,
		0L,
		&cValues,
		&lpPropsMsg));

	lpMsgData->szFilePath[0] = NULL;
	// If we've got a parent message, we're an attachment - use attachment filename logic
	if (lpParentMessageData)
	{
		// Take the parent filename/path, remove any extension, and append -Attach.xml
		// Should we append something for attachment number?
		size_t  cchLen = 0;

		// Copy the source string over
		WC_H(StringCchCopyW(lpMsgData->szFilePath, _countof(lpMsgData->szFilePath), ((LPMESSAGEDATA) lpParentMessageData)->szFilePath));

		// Remove any extension
		WC_H(StringCchLengthW(lpMsgData->szFilePath,_countof(lpMsgData->szFilePath),&cchLen));
		while (cchLen > 0 && lpMsgData->szFilePath[cchLen] != L'.') cchLen--;
		if (lpMsgData->szFilePath[cchLen] == L'.') lpMsgData->szFilePath[cchLen] = L'\0';

		// build a string for appending
		WCHAR	szNewExt[MAX_PATH];
		WC_H(StringCchPrintfW(szNewExt,_countof(szNewExt),L"-Attach%d.xml",((LPMESSAGEDATA) lpParentMessageData)->ulCurAttNum)); // STRING_OK

		// append our string
		WC_H(StringCchCatW(lpMsgData->szFilePath,_countof(lpMsgData->szFilePath),szNewExt));

		OutputToFilef(((LPMESSAGEDATA) lpParentMessageData)->fMessageProps,_T("<embeddedmessage path=\"%ws\"/>\n"),lpMsgData->szFilePath);
	}
	else if (m_szMessageFileName[0]) // if we've got a file name, use it
	{
		WC_H(StringCchCopyW(lpMsgData->szFilePath, _countof(lpMsgData->szFilePath), m_szMessageFileName));
	}
	else
	{
		LPCWSTR szSubj = NULL; // BuildFileNameAndPath will substitute a subject if we don't find one
		LPSBinary lpSearchKey = NULL;

		if (CheckStringProp(&lpPropsMsg[ePR_SUBJECT_W],PT_UNICODE))
		{
			szSubj = lpPropsMsg[ePR_SUBJECT_W].Value.lpszW;
		}
		if (PR_SEARCH_KEY == lpPropsMsg[ePR_SEARCH_KEY].ulPropTag)
		{
			lpSearchKey = &lpPropsMsg[ePR_SEARCH_KEY].Value.bin;
		}
		WC_H(BuildFileNameAndPath(lpMsgData->szFilePath,_countof(lpMsgData->szFilePath),L".xml",4,szSubj,lpSearchKey,m_szFolderPath)); // STRING_OK
	}

	DebugPrint(DBGGeneric,_T("OutputMessagePropertiesToFile: Saving %p to \"%ws\"\n"),lpMessage,lpMsgData->szFilePath);
	lpMsgData->fMessageProps = OpenFile(lpMsgData->szFilePath,true);

	if (lpMsgData->fMessageProps)
	{
		OutputToFile(lpMsgData->fMessageProps,_T("<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?>\n"));
		OutputToFile(lpMsgData->fMessageProps,_T("<message>\n"));
		if (lpPropsMsg)
		{
			OutputToFile(lpMsgData->fMessageProps,_T("<properties listtype=\"summary\">\n"));

			OutputPropertiesToFile(lpMsgData->fMessageProps, cValues, lpPropsMsg, lpMessage);

			OutputToFile(lpMsgData->fMessageProps,_T("</properties>\n"));
			MAPIFreeBuffer(lpPropsMsg);
		}

		hRes = S_OK; // we've logged our warning if we had one - mask it now.

		LPSTREAM lpStream = NULL;
		// Log Body
		WC_H(lpMessage->OpenProperty(
			PR_BODY,
			&IID_IStream,
			STGM_READ,
			NULL,
			(LPUNKNOWN *) &lpStream));
		if (lpStream)
		{
			OutputToFile(lpMsgData->fMessageProps,_T("<body property=\"PR_BODY\">\n"));
			OutputCDataOpen(lpMsgData->fMessageProps);
			OutputStreamToFile(lpMsgData->fMessageProps, lpStream);
			lpStream->Release();
			lpStream = NULL;
			OutputCDataClose(lpMsgData->fMessageProps);
			OutputToFile(lpMsgData->fMessageProps,_T("</body>\n"));
		}

		hRes = S_OK; // we've logged our warning if we had one - mask it now.

		// Log HTML Body
		WC_H(lpMessage->OpenProperty(
			PR_BODY_HTML,
			&IID_IStream,
			STGM_READ,
			NULL,
			(LPUNKNOWN *) &lpStream));
		if (lpStream)
		{
			OutputToFile(lpMsgData->fMessageProps,_T("<body property=\"PR_BODY_HTML\">\n"));
			OutputCDataOpen(lpMsgData->fMessageProps);
			OutputStreamToFile(lpMsgData->fMessageProps, lpStream);
			lpStream->Release();
			lpStream = NULL;
			OutputCDataClose(lpMsgData->fMessageProps);
			OutputToFile(lpMsgData->fMessageProps,_T("</body>\n"));
		}
		OutputToFile(lpMsgData->fMessageProps,_T("\n"));


		hRes = S_OK; // we've logged our warning if we had one - mask it now.

		// Log RTF Body
		LPSTREAM lpRTFUncompressed = NULL;

		// TODO: This fails in unicode builds since PR_RTF_COMPRESSED is always ascii.
		WC_H(lpMessage->OpenProperty(
			PR_RTF_COMPRESSED,
			&IID_IStream,
			STGM_READ,
			NULL,
			(LPUNKNOWN *)&lpStream));
		if (lpStream)
		{
			WC_H(WrapStreamForRTF(
				lpStream,
				false,
				NULL,
				NULL,
				NULL,
				&lpRTFUncompressed,
				NULL));
			if (lpRTFUncompressed)
			{
				OutputToFile(lpMsgData->fMessageProps,_T("<body property=\"PR_RTF_COMPRESSED\">\n"));
				OutputCDataOpen(lpMsgData->fMessageProps);
				OutputStreamToFile(lpMsgData->fMessageProps, lpRTFUncompressed);
				lpRTFUncompressed->Release();
				lpRTFUncompressed = NULL;
				OutputCDataClose(lpMsgData->fMessageProps);
				OutputToFile(lpMsgData->fMessageProps,_T("</body>\n"));
			}
			lpStream->Release();
			lpStream = NULL;
		}

		if (pfnWrapEx)
		{
			// Log best Body using WrapCompressedRTFEx
			WC_H(lpMessage->OpenProperty(
				PR_RTF_COMPRESSED,
				&IID_IStream,
				STGM_READ,
				NULL,
				(LPUNKNOWN *)&lpStream));
			if (lpStream)
			{
				ULONG ulStreamFlags = NULL;
				ULONG ulInCodePage = CP_ACP; // picking CP_ACP as our default

				if (PT_LONG == PROP_TYPE(lpPropsMsg[ePR_INTERNET_CPID].ulPropTag))
				{
					ulInCodePage = lpPropsMsg[ePR_INTERNET_CPID].Value.l;
				}
				WC_H(WrapStreamForRTF(
					lpStream,
					true,
					MAPI_NATIVE_BODY,
					ulInCodePage,
					CP_ACP, // requesting ANSI code page - check if this will be valid in UNICODE builds
					&lpRTFUncompressed,
					&ulStreamFlags));
				if (lpRTFUncompressed)
				{
					OutputToFile(lpMsgData->fMessageProps,_T("<body property=\"WrapCompressedRTFEx best body\""));
					LPTSTR szFlags = NULL;
					InterpretFlags(flagStreamFlag, ulStreamFlags, &szFlags);
					OutputToFilef(lpMsgData->fMessageProps,_T(" ulStreamFlags = \"0x%08X\" szStreamFlags= \"%s\""),ulStreamFlags,szFlags);
					delete[] szFlags;
					szFlags = NULL;
					OutputToFilef(lpMsgData->fMessageProps,_T(" CodePageIn = \"%d\" CodePageOut = \"%d\">\n"),ulInCodePage,CP_ACP);
					OutputCDataOpen(lpMsgData->fMessageProps);
					OutputStreamToFile(lpMsgData->fMessageProps, lpRTFUncompressed);
					lpRTFUncompressed->Release();
					lpRTFUncompressed = NULL;
					OutputCDataClose(lpMsgData->fMessageProps);
					OutputToFile(lpMsgData->fMessageProps,_T("</body>\n"));
				}
				lpStream->Release();
				lpStream = NULL;
			}
		}

		hRes = S_OK; // we've logged our warning if we had one - mask it now.

		LPSPropValue lpAllProps = NULL;
		WC_H_GETPROPS(GetPropsNULL(lpMessage,
			fMapiUnicode,
			&cValues,
			&lpAllProps));
		if (lpAllProps)
		{
			OutputToFile(lpMsgData->fMessageProps,_T("<properties listtype=\"FullPropList\">\n"));

			OutputPropertiesToFile(lpMsgData->fMessageProps, cValues, lpAllProps, lpMessage);

			OutputToFile(lpMsgData->fMessageProps,_T("</properties>\n"));
			MAPIFreeBuffer(lpAllProps);
		}
	}
} // CDumpStore::BeginMessageWork

void CDumpStore::BeginRecipientWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
{
	if (!lpData) return;
	OutputToFile(((LPMESSAGEDATA) lpData)->fMessageProps,_T("<recipients>\n"));
} // CDumpStore::BeginRecipientWork

void CDumpStore::DoMessagePerRecipientWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData, _In_ LPSRow lpSRow, ULONG ulCurRow)
{
	if (!lpMessage || !lpData || !lpSRow) return;

	LPMESSAGEDATA lpMsgData = (LPMESSAGEDATA) lpData;

	OutputToFilef(lpMsgData->fMessageProps,_T("<recipient num=\"0x%08X\">\n"),ulCurRow);

	OutputSRowToFile(lpMsgData->fMessageProps,lpSRow, lpMessage);

	OutputToFile(lpMsgData->fMessageProps,_T("</recipient>\n"));
} // CDumpStore::DoMessagePerRecipientWork

void CDumpStore::EndRecipientWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
{
	if (!lpData) return;
	OutputToFile(((LPMESSAGEDATA) lpData)->fMessageProps,_T("</recipients>\n"));
} // CDumpStore::EndRecipientWork

void CDumpStore::BeginAttachmentWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
{
	if (!lpData) return;
	OutputToFile(((LPMESSAGEDATA) lpData)->fMessageProps,_T("<attachments>\n"));
} // CDumpStore::BeginAttachmentWork

void CDumpStore::DoMessagePerAttachmentWork(_In_ LPMESSAGE lpMessage, _In_ LPVOID lpData, _In_ LPSRow lpSRow, _In_ LPATTACH lpAttach, ULONG ulCurRow)
{
	if (!lpMessage || !lpData || !lpSRow) return;
	HRESULT hRes = S_OK;

	LPMESSAGEDATA lpMsgData = (LPMESSAGEDATA) lpData;

	lpMsgData->ulCurAttNum = ulCurRow; // set this so we can pull it if this is an embedded message

	LPSPropValue lpAttachName = NULL;

	hRes = S_OK;

	lpAttachName = PpropFindProp(
		lpSRow->lpProps,
		lpSRow->cValues,
		PR_ATTACH_FILENAME);

	OutputToFilef(lpMsgData->fMessageProps,_T("<attachment num=\"0x%08X\" filename=\""),ulCurRow);

	if (CheckStringProp(lpAttachName,PT_TSTRING))
		OutputToFile(lpMsgData->fMessageProps,lpAttachName->Value.LPSZ);
	else
		OutputToFile(lpMsgData->fMessageProps,_T("PR_ATTACH_FILENAME not found"));
	OutputToFile(lpMsgData->fMessageProps,_T("\">\n"));

	OutputToFile(lpMsgData->fMessageProps,_T("\t<tableprops>\n"));
	OutputSRowToFile(lpMsgData->fMessageProps,lpSRow, lpMessage);

	OutputToFile(lpMsgData->fMessageProps,_T("\t</tableprops>\n"));

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

			OutputToFile(lpMsgData->fMessageProps,_T("\t<getprops>\n"));
			OutputPropertiesToFile(lpMsgData->fMessageProps,ulAllProps,lpAllProps,lpMessage);
			OutputToFile(lpMsgData->fMessageProps,_T("\t</getprops>\n"));

		}
		MAPIFreeBuffer(lpAllProps);
		lpAllProps = NULL;
	}
	OutputToFile(lpMsgData->fMessageProps,_T("</attachment>\n"));
} // CDumpStore::DoMessagePerAttachmentWork

void CDumpStore::EndAttachmentWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
{
	if (!lpData) return;
	OutputToFile(((LPMESSAGEDATA) lpData)->fMessageProps,_T("</attachments>\n"));
} // CDumpStore::EndAttachmentWork

void CDumpStore::EndMessageWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID lpData)
{
	LPMESSAGEDATA lpMsgData = (LPMESSAGEDATA) lpData;

	if (lpMsgData->fMessageProps)
	{
		OutputToFile(lpMsgData->fMessageProps,_T("</message>\n"));
		CloseFile(lpMsgData->fMessageProps);
	}
	delete lpMsgData;
} // CDumpStore::EndMessageWork