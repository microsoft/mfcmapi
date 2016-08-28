#include "stdafx.h"
#include "MAPIProcessor.h"
#include "MAPIFunctions.h"
#include "MAPIStoreFunctions.h"
#include "File.h"
#include "ColumnTags.h"

CMAPIProcessor::CMAPIProcessor()
{
	m_lpSession = NULL;
	m_lpMDB = NULL;
	m_lpFolder = NULL;
	m_szFolderOffset = NULL;
	m_lpListHead = NULL;
	m_lpListTail = NULL;
	m_lpResFolderContents = NULL;
	m_lpSort = NULL;
	m_ulCount = NULL;
}

CMAPIProcessor::~CMAPIProcessor()
{
	FreeFolderList();
	MAPIFreeBuffer(m_szFolderOffset);
	if (m_lpFolder) m_lpFolder->Release();
	if (m_lpMDB) m_lpMDB->Release();
	if (m_lpSession) m_lpSession->Release();
}

// --------------------------------------------------------------------------------- //

void CMAPIProcessor::InitSession(_In_ LPMAPISESSION lpSession)
{
	if (m_lpSession) m_lpSession->Release();
	m_lpSession = lpSession;
	if (m_lpSession) m_lpSession->AddRef();
}

void CMAPIProcessor::InitMDB(_In_ LPMDB lpMDB)
{
	if (m_lpMDB) m_lpMDB->Release();
	m_lpMDB = lpMDB;
	if (m_lpMDB) m_lpMDB->AddRef();
}

void CMAPIProcessor::InitFolder(_In_ LPMAPIFOLDER lpFolder)
{
	if (m_lpFolder) m_lpFolder->Release();
	m_lpFolder = lpFolder;
	if (m_lpFolder) m_lpFolder->AddRef();
	MAPIFreeBuffer(m_szFolderOffset);
	HRESULT hRes = S_OK;
	WC_H(CopyString(&m_szFolderOffset, _T("\\"), NULL)); // STRING_OK
}

void CMAPIProcessor::InitFolderContentsRestriction(_In_opt_ LPSRestriction lpRes)
{
	// If we ever need to hold this past the scope of the caller we'll need to copy the restriction.
	// For now, just grab a pointer.
	m_lpResFolderContents = lpRes;
}

void CMAPIProcessor::InitMaxOutput(_In_ ULONG ulCount)
{
	m_ulCount = ulCount;
}

void CMAPIProcessor::InitSortOrder(_In_ LPSSortOrderSet lpSort)
{
	// If we ever need to hold this past the scope of the caller we'll need to copy the sort order.
	// For now, just grab a pointer.
	m_lpSort = lpSort;
}

// --------------------------------------------------------------------------------- //

// Server name MUST be passed
void CMAPIProcessor::ProcessMailboxTable(
	_In_ wstring szExchangeServerName)
{
	if (szExchangeServerName.empty()) return;
	HRESULT hRes = S_OK;

	LPMAPITABLE lpMailBoxTable = NULL;
	LPSRowSet lpRows = NULL;
	LPMDB lpPrimaryMDB = NULL;
	ULONG ulOffset = 0;
	ULONG ulRowNum = 0;
	ULONG ulFlags = LOGOFF_NO_WAIT;

	if (!m_lpSession) return;

	BeginMailboxTableWork(szExchangeServerName);

	WC_H(OpenMessageStoreGUID(m_lpSession, pbExchangeProviderPrimaryUserGuid, &lpPrimaryMDB));

	if (lpPrimaryMDB && StoreSupportsManageStore(lpPrimaryMDB)) do
	{
		hRes = S_OK;
		WC_H(GetMailboxTable(
			lpPrimaryMDB,
			szExchangeServerName,
			ulOffset,
			&lpMailBoxTable));
		if (lpMailBoxTable)
		{
			WC_MAPI(lpMailBoxTable->SetColumns((LPSPropTagArray)&sptMBXCols, NULL));
			hRes = S_OK;

			// go to the first row
			WC_MAPI(lpMailBoxTable->SeekRow(
				BOOKMARK_BEGINNING,
				0,
				NULL));

			// get each row in turn and process it
			if (!FAILED(hRes)) for (ulRowNum = 0;; ulRowNum++)
			{
				if (lpRows) FreeProws(lpRows);
				lpRows = NULL;

				hRes = S_OK;
				WC_MAPI(lpMailBoxTable->QueryRows(
					1,
					NULL,
					&lpRows));
				if (FAILED(hRes) || !lpRows || !lpRows->cRows) break;

				DoMailboxTablePerRowWork(lpPrimaryMDB, lpRows->aRow, ulRowNum + ulOffset + 1);

				LPSPropValue lpEmailAddress = PpropFindProp(lpRows->aRow->lpProps, lpRows->aRow->cValues, PR_EMAIL_ADDRESS);

				if (!CheckStringProp(lpEmailAddress, PT_TSTRING)) continue;

				if (m_lpMDB)
				{
					WC_MAPI(m_lpMDB->StoreLogoff(&ulFlags));
					m_lpMDB->Release();
					m_lpMDB = NULL;
					hRes = S_OK;
				}

				WC_H(OpenOtherUsersMailbox(
					m_lpSession,
					lpPrimaryMDB,
					szExchangeServerName,
					LPCTSTRToWstring(lpEmailAddress->Value.LPSZ),
					OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_TAKE_OWNERSHIP,
					false,
					&m_lpMDB));

				if (m_lpMDB) ProcessStore();
			}

			if (lpRows) FreeProws(lpRows);
			lpRows = NULL;
			lpMailBoxTable->Release();
			ulOffset += ulRowNum;
		}
	} while (ulRowNum);

	if (m_lpMDB)
	{
		WC_MAPI(m_lpMDB->StoreLogoff(&ulFlags));
		m_lpMDB->Release();
		m_lpMDB = NULL;
	}

	if (lpPrimaryMDB) lpPrimaryMDB->Release();
	EndMailboxTableWork();
}

void CMAPIProcessor::ProcessStore()
{
	if (!m_lpMDB) return;

	BeginStoreWork();

	AddFolderToFolderList(NULL, _T("\\")); // STRING_OK

	ProcessFolders(true, true, true);

	EndStoreWork();
}

void CMAPIProcessor::ProcessFolders(bool bDoRegular, bool bDoAssociated, bool bDoDescent)
{
	BeginProcessFoldersWork();

	if (ContinueProcessingFolders())
	{
		if (!m_lpFolder)
		{
			OpenFirstFolderInList();
		}

		while (m_lpFolder)
		{
			DoProcessFoldersPerFolderWork();
			ProcessFolder(bDoRegular, bDoAssociated, bDoDescent);
			if (!ContinueProcessingFolders()) break;
			OpenFirstFolderInList();
		}
	}
	EndProcessFoldersWork();
}

bool CMAPIProcessor::ContinueProcessingFolders()
{
	return true;
}

bool CMAPIProcessor::ShouldProcessContentsTable()
{
	return true;
}

void CMAPIProcessor::ProcessFolder(bool bDoRegular,
	bool bDoAssociated,
	bool bDoDescent)
{
	if (!m_lpMDB || !m_lpFolder) return;

	HRESULT hRes = S_OK;

	BeginFolderWork();

	if (bDoRegular && ShouldProcessContentsTable())
	{
		ProcessContentsTable(NULL);
	}

	if (bDoAssociated && ShouldProcessContentsTable())
	{
		ProcessContentsTable(MAPI_ASSOCIATED);
	}

	// If we're not processing subfolders, then get outta here
	if (bDoDescent)
	{
		LPMAPITABLE lpHierarchyTable = NULL;
		// We need to walk down the tree
		// and get the list of kids of the folder
		WC_MAPI(m_lpFolder->GetHierarchyTable(fMapiUnicode, &lpHierarchyTable));
		if (lpHierarchyTable)
		{
			enum
			{
				NAME,
				EID,
				SUBFOLDERS,
				FLAGS,
				NUMCOLS
			};
			static const SizedSPropTagArray(NUMCOLS, sptHierarchyCols) =
			{
			NUMCOLS,
			PR_DISPLAY_NAME,
			PR_ENTRYID,
			PR_SUBFOLDERS,
			PR_CONTAINER_FLAGS,
			};

			LPSRowSet lpRows = NULL;
			// If I don't do this, the MSPST provider can blow chunks (MAPI_E_EXTENDED_ERROR) for some folders when I get a row
			// For some reason, this fixes it.
			WC_MAPI(lpHierarchyTable->SetColumns(
				(LPSPropTagArray)&sptHierarchyCols,
				TBL_BATCH));

			// go to the first row
			WC_MAPI(lpHierarchyTable->SeekRow(
				BOOKMARK_BEGINNING,
				0,
				NULL));

			if (S_OK == hRes) for (;;)
			{
				hRes = S_OK;
				if (lpRows) FreeProws(lpRows);
				lpRows = NULL;
				WC_MAPI(lpHierarchyTable->QueryRows(
					255,
					NULL,
					&lpRows));
				if (FAILED(hRes))
				{
					hRes = S_OK;
					break;
				}
				if (!lpRows || !lpRows->cRows) break;

				ULONG ulRow = 0;

				for (ulRow = 0; ulRow < lpRows->cRows; ulRow++)
				{
					DoFolderPerHierarchyTableRowWork(&lpRows->aRow[ulRow]);
					if (lpRows->aRow[ulRow].lpProps)
					{
						TCHAR szSubFolderOffset[MAX_PATH]; // Holds subfolder name

						LPSPropValue lpFolderDisplayName = NULL;

						lpFolderDisplayName = PpropFindProp(
							lpRows->aRow[ulRow].lpProps,
							lpRows->aRow[ulRow].cValues,
							PR_DISPLAY_NAME);

						if (CheckStringProp(lpFolderDisplayName, PT_TSTRING))
						{
							TCHAR szTemp[MAX_PATH / 2];
							// Clean up the folder name before appending it to the offset
							WC_H(SanitizeFileName(szTemp, _countof(szTemp), lpFolderDisplayName->Value.LPSZ, _countof(szTemp)));

							WC_H(StringCchPrintf(szSubFolderOffset, _countof(szSubFolderOffset),
								_T("%s%s\\"), // STRING_OK
								m_szFolderOffset, szTemp));
						}
						else
						{
							WC_H(StringCchPrintf(szSubFolderOffset, _countof(szSubFolderOffset),
								_T("%s\\UnknownFolder\\"), // STRING_OK
								m_szFolderOffset));
						}

						AddFolderToFolderList(&lpRows->aRow[ulRow].lpProps[EID].Value.bin, szSubFolderOffset);
					}
				}
			}

			if (lpRows) FreeProws(lpRows);
			lpHierarchyTable->Release();
		}
	}
	EndFolderWork();
}

void CMAPIProcessor::ProcessContentsTable(ULONG ulFlags)
{
	if (!m_lpFolder) return;

	enum
	{
		contPR_SUBJECT,
		contPR_MESSAGE_CLASS,
		contPR_MESSAGE_DELIVERY_TIME,
		contPR_HASATTACH,
		contPR_ENTRYID,
		contPR_SEARCH_KEY,
		contPR_RECORD_KEY,
		contPidTagMid,
		contNUM_COLS
	};
	static const SizedSPropTagArray(contNUM_COLS, contCols) =
	{
	contNUM_COLS,
	PR_MESSAGE_CLASS,
	PR_SUBJECT,
	PR_MESSAGE_DELIVERY_TIME,
	PR_HASATTACH,
	PR_ENTRYID,
	PR_SEARCH_KEY,
	PR_RECORD_KEY,
	PidTagMid,
	};
	HRESULT hRes = S_OK;

	LPMAPITABLE lpContentsTable = NULL;

	WC_MAPI(m_lpFolder->GetContentsTable(
		ulFlags | fMapiUnicode,
		&lpContentsTable));
	if (SUCCEEDED(hRes) && lpContentsTable)
	{
		WC_MAPI(lpContentsTable->SetColumns((LPSPropTagArray)&contCols, TBL_BATCH));
	}

	if (SUCCEEDED(hRes) && lpContentsTable && m_lpResFolderContents)
	{
		DebugPrintRestriction(DBGGeneric, m_lpResFolderContents, NULL);
		WC_MAPI(lpContentsTable->Restrict(m_lpResFolderContents, TBL_BATCH));
		hRes = S_OK;
	}

	if (SUCCEEDED(hRes) && lpContentsTable && m_lpSort)
	{
		WC_MAPI(lpContentsTable->SortTable(m_lpSort, TBL_BATCH));
		hRes = S_OK;
	}

	if (SUCCEEDED(hRes) && lpContentsTable)
	{
		ULONG ulCountRows = 0;
		WC_MAPI(lpContentsTable->GetRowCount(0, &ulCountRows));

		BeginContentsTableWork(ulFlags, ulCountRows);

		LPSRowSet lpRows = NULL;
		ULONG i = 0;
		for (;;)
		{
			// If we've output enough rows, stop
			if (m_ulCount && (i >= m_ulCount)) break;
			hRes = S_OK;
			if (lpRows) FreeProws(lpRows);
			lpRows = NULL;
			WC_MAPI(lpContentsTable->QueryRows(
				255,
				NULL,
				&lpRows));
			if (FAILED(hRes))
			{
				hRes = S_OK;
				break;
			}
			if (!lpRows || !lpRows->cRows) break;

			ULONG iRow = 0;
			for (iRow = 0; iRow < lpRows->cRows; iRow++)
			{
				// If we've output enough rows, stop
				if (m_ulCount && (i >= m_ulCount)) break;
				if (!DoContentsTablePerRowWork(&lpRows->aRow[iRow], i++)) continue;

				LPSPropValue lpMsgEID = NULL;
				lpMsgEID = PpropFindProp(
					lpRows->aRow[iRow].lpProps,
					lpRows->aRow[iRow].cValues,
					PR_ENTRYID);

				if (!lpMsgEID) continue;

				LPMESSAGE lpMessage = NULL;
				WC_H(CallOpenEntry(
					NULL,
					NULL,
					m_lpFolder,
					NULL,
					lpMsgEID->Value.bin.cb,
					(LPENTRYID)lpMsgEID->Value.bin.lpb,
					NULL,
					MAPI_BEST_ACCESS,
					NULL,
					(LPUNKNOWN*)&lpMessage));

				if (lpMessage)
				{
					LPSPropValue lpMsgHasAttach = NULL;
					bool bHasAttach = true;
					lpMsgHasAttach = PpropFindProp(
						lpRows->aRow[iRow].lpProps,
						lpRows->aRow[iRow].cValues,
						PR_HASATTACH);
					if (lpMsgHasAttach)
					{
						bHasAttach = (0 != lpMsgHasAttach->Value.b);
					}
					ProcessMessage(lpMessage, bHasAttach, NULL);
					lpMessage->Release();
				}

				lpMessage = NULL;
			}
		}

		if (lpRows) FreeProws(lpRows);
		lpContentsTable->Release();
	}
	EndContentsTableWork();
}

void CMAPIProcessor::ProcessMessage(_In_ LPMESSAGE lpMessage, bool bHasAttach, _In_opt_ LPVOID lpParentMessageData)
{
	if (!lpMessage) return;

	LPVOID lpData = NULL;
	if (BeginMessageWork(lpMessage, lpParentMessageData, &lpData))
	{
		ProcessRecipients(lpMessage, lpData);
		ProcessAttachments(lpMessage, bHasAttach, lpData);
		EndMessageWork(lpMessage, lpData);
	}
}

void CMAPIProcessor::ProcessRecipients(_In_ LPMESSAGE lpMessage, _In_opt_ LPVOID lpData)
{
	if (!lpMessage) return;

	HRESULT hRes = S_OK;
	LPMAPITABLE lpRecipTable = NULL;
	LPSRowSet lpRows = NULL;

	if (!BeginRecipientWork(lpMessage, lpData)) return;

	// get the recipient table
	WC_MAPI(lpMessage->GetRecipientTable(
		NULL,
		&lpRecipTable));
	if (lpRecipTable)
	{
		ULONG i = 0;
		for (;;)
		{
			hRes = S_OK;
			if (lpRows) FreeProws(lpRows);
			lpRows = NULL;
			WC_MAPI(lpRecipTable->QueryRows(
				255,
				NULL,
				&lpRows));
			if (FAILED(hRes))
			{
				hRes = S_OK;
				break;
			}
			if (!lpRows || !lpRows->cRows) break;

			ULONG iRow = 0;
			for (iRow = 0; iRow < lpRows->cRows; iRow++)
			{
				DoMessagePerRecipientWork(lpMessage, lpData, &lpRows->aRow[iRow], i++);
			}
		}

		if (lpRows) FreeProws(lpRows);
		if (lpRecipTable) lpRecipTable->Release();
	}

	EndRecipientWork(lpMessage, lpData);
}

void CMAPIProcessor::ProcessAttachments(_In_ LPMESSAGE lpMessage, bool bHasAttach, _In_opt_ LPVOID lpData)
{
	if (!lpMessage) return;

	enum
	{
		attPR_ATTACH_NUM,
		atPR_ATTACH_METHOD,
		attPR_ATTACH_FILENAME,
		attPR_DISPLAY_NAME,
		attNUM_COLS
	};
	static const SizedSPropTagArray(attNUM_COLS, attCols) =
	{
	attNUM_COLS,
	PR_ATTACH_NUM,
	PR_ATTACH_METHOD,
	PR_ATTACH_FILENAME,
	PR_DISPLAY_NAME,
	};
	HRESULT hRes = S_OK;
	LPMAPITABLE lpAttachTable = NULL;
	LPSRowSet lpRows = NULL;

	if (!BeginAttachmentWork(lpMessage, lpData)) return;

	// Only get the attachment table if PR_HASATTACH was true or didn't exist.
	if (bHasAttach)
	{
		// get the attachment table
		WC_MAPI(lpMessage->GetAttachmentTable(
			NULL,
			&lpAttachTable));
	}

	if (SUCCEEDED(hRes) && lpAttachTable)
	{
		// Make sure we have the columns we need
		WC_MAPI(lpAttachTable->SetColumns((LPSPropTagArray)&attCols, TBL_BATCH));
	}

	if (SUCCEEDED(hRes) && lpAttachTable)
	{
		ULONG i = 0;
		for (;;)
		{
			hRes = S_OK;
			if (lpRows) FreeProws(lpRows);
			lpRows = NULL;
			WC_MAPI(lpAttachTable->QueryRows(
				255,
				NULL,
				&lpRows));
			if (FAILED(hRes))
			{
				hRes = S_OK;
				break;
			}
			if (!lpRows || !lpRows->cRows) break;

			ULONG iRow = 0;
			for (iRow = 0; iRow < lpRows->cRows; iRow++)
			{
				hRes = S_OK;
				LPSPropValue lpAttachNum = NULL;
				lpAttachNum = PpropFindProp(
					lpRows->aRow[iRow].lpProps,
					lpRows->aRow[iRow].cValues,
					PR_ATTACH_NUM);

				if (lpAttachNum)
				{
					LPATTACH lpAttach = NULL;
					WC_MAPI(lpMessage->OpenAttach(
						lpAttachNum->Value.l,
						NULL,
						MAPI_BEST_ACCESS,
						(LPATTACH*)&lpAttach));

					DoMessagePerAttachmentWork(lpMessage, lpData, &lpRows->aRow[iRow], lpAttach, i++);
					// Check if the attachment is an embedded message - if it is, parse it as such!
					LPSPropValue lpAttachMethod = NULL;

					lpAttachMethod = PpropFindProp(
						lpRows->aRow->lpProps,
						lpRows->aRow->cValues,
						PR_ATTACH_METHOD);

					if (lpAttach && lpAttachMethod && ATTACH_EMBEDDED_MSG == lpAttachMethod->Value.l)
					{
						LPMESSAGE lpAttachMsg = NULL;

						WC_MAPI(lpAttach->OpenProperty(
							PR_ATTACH_DATA_OBJ,
							(LPIID)&IID_IMessage,
							0,
							NULL, // MAPI_MODIFY,
							(LPUNKNOWN *)&lpAttachMsg));
						if (lpAttachMsg)
						{
							// Just assume this message might have attachments
							ProcessMessage(lpAttachMsg, true, lpData);
							lpAttachMsg->Release();
						}
					}
					if (lpAttach)
					{
						lpAttach->Release();
						lpAttach = NULL;
					}
				}
			}
		}
		if (lpRows) FreeProws(lpRows);
		lpAttachTable->Release();
	}

	EndAttachmentWork(lpMessage, lpData);
}

// --------------------------------------------------------------------------------- //
// List Functions
// --------------------------------------------------------------------------------- //
void CMAPIProcessor::AddFolderToFolderList(_In_opt_ LPSBinary lpFolderEID, _In_z_ LPCTSTR szFolderOffsetPath)
{
	HRESULT hRes = S_OK;
	LPFOLDERNODE lpNewNode = NULL;
	WC_H(MAPIAllocateBuffer(sizeof(FolderNode), (LPVOID*)&lpNewNode));
	lpNewNode->lpNextFolder = NULL;

	lpNewNode->lpFolderEID = NULL;
	if (lpFolderEID)
	{
		WC_H(MAPIAllocateMore(
			(ULONG)sizeof(SBinary),
			lpNewNode,
			(LPVOID*)&lpNewNode->lpFolderEID));
		WC_H(CopySBinary(lpNewNode->lpFolderEID, lpFolderEID, lpNewNode));
	}

	WC_H(CopyString(&lpNewNode->szFolderOffsetPath, szFolderOffsetPath, lpNewNode));

	if (!m_lpListHead)
	{
		m_lpListHead = lpNewNode;
		m_lpListTail = lpNewNode;
	}
	else
	{
		m_lpListTail->lpNextFolder = lpNewNode;
		m_lpListTail = lpNewNode;
	}
}

// Call OpenEntry on the first folder in the list, remove it from the list
// If we fail to open a folder, move on to the next item in the list
void CMAPIProcessor::OpenFirstFolderInList()
{
	HRESULT hRes = S_OK;
	LPMAPIFOLDER lpFolder = NULL;
	MAPIFreeBuffer(m_szFolderOffset);
	m_szFolderOffset = NULL;

	if (m_lpFolder) m_lpFolder->Release();
	m_lpFolder = NULL;

	// loop over nodes until we open one or run out
	while (!lpFolder && m_lpListHead)
	{
		hRes = S_OK;

		WC_H(CallOpenEntry(
			m_lpMDB,
			NULL,
			NULL,
			NULL,
			m_lpListHead->lpFolderEID,
			NULL,
			MAPI_BEST_ACCESS,
			NULL,
			(LPUNKNOWN*)&lpFolder));
		if (m_lpListHead->szFolderOffsetPath)
		{
			WC_H(CopyString(&m_szFolderOffset, m_lpListHead->szFolderOffsetPath, NULL));
		}

		LPFOLDERNODE lpTempNode = NULL;
		lpTempNode = m_lpListHead;
		m_lpListHead = m_lpListHead->lpNextFolder;
		if (m_lpListHead == lpTempNode) m_lpListTail = NULL;
		MAPIFreeBuffer(lpTempNode);
	}

	m_lpFolder = lpFolder;
}

// Clean up the list
void CMAPIProcessor::FreeFolderList()
{
	if (!m_lpListHead) return;
	LPFOLDERNODE lpTempNode = NULL;
	while (m_lpListHead)
	{
		lpTempNode = m_lpListHead->lpNextFolder;
		MAPIFreeBuffer(m_lpListHead);
		m_lpListHead = lpTempNode;
	}
	m_lpListTail = NULL;
}

// --------------------------------------------------------------------------------- //
// Worker Functions
// --------------------------------------------------------------------------------- //

void CMAPIProcessor::BeginMailboxTableWork(_In_ wstring /*szExchangeServerName*/)
{
}

void CMAPIProcessor::DoMailboxTablePerRowWork(_In_ LPMDB /*lpMDB*/, _In_ LPSRow /*lpSRow*/, ULONG /*ulCurRow*/)
{
}

void CMAPIProcessor::EndMailboxTableWork()
{
}

void CMAPIProcessor::BeginStoreWork()
{
}

void CMAPIProcessor::EndStoreWork()
{
}

void CMAPIProcessor::BeginProcessFoldersWork()
{
}

void CMAPIProcessor::DoProcessFoldersPerFolderWork()
{
}

void CMAPIProcessor::EndProcessFoldersWork()
{
}

void CMAPIProcessor::BeginFolderWork()
{
}

void CMAPIProcessor::DoFolderPerHierarchyTableRowWork(_In_ LPSRow /*lpSRow*/)
{
}

void CMAPIProcessor::EndFolderWork()
{
}

void CMAPIProcessor::BeginContentsTableWork(ULONG /*ulFlags*/, ULONG /*ulCountRows*/)
{
}

bool CMAPIProcessor::DoContentsTablePerRowWork(_In_ LPSRow /*lpSRow*/, ULONG /*ulCurRow*/)
{
	return true; // Keep processing
}

void CMAPIProcessor::EndContentsTableWork()
{
}

bool CMAPIProcessor::BeginMessageWork(_In_ LPMESSAGE /*lpMessage*/, _In_opt_ LPVOID /*lpParentMessageData*/, _Deref_out_opt_ LPVOID* lpData)
{
	if (lpData) *lpData = NULL;
	return true; // Keep processing
}

bool CMAPIProcessor::BeginRecipientWork(_In_ LPMESSAGE /*lpMessage*/, _In_opt_ LPVOID /*lpData*/)
{
	return true; // Keep processing
}

void CMAPIProcessor::DoMessagePerRecipientWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID /*lpData*/, _In_ LPSRow /*lpSRow*/, ULONG /*ulCurRow*/)
{
}

void CMAPIProcessor::EndRecipientWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID /*lpData*/)
{
}

bool CMAPIProcessor::BeginAttachmentWork(_In_ LPMESSAGE /*lpMessage*/, _In_opt_ LPVOID /*lpData*/)
{
	return true; // Keep processing
}

void CMAPIProcessor::DoMessagePerAttachmentWork(_In_ LPMESSAGE /*lpMessage*/, _In_ LPVOID /*lpData*/, _In_ LPSRow /*lpSRow*/, _In_ LPATTACH /*lpAttach*/, ULONG /*ulCurRow*/)
{
}

void CMAPIProcessor::EndAttachmentWork(_In_ LPMESSAGE /*lpMessage*/, _In_opt_ LPVOID /*lpData*/)
{
}

void CMAPIProcessor::EndMessageWork(_In_ LPMESSAGE /*lpMessage*/, _In_opt_ LPVOID /*lpData*/)
{
}
