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
}

CMAPIProcessor::~CMAPIProcessor()
{
	FreeFolderList();
	MAPIFreeBuffer(m_szFolderOffset);
	if (m_lpFolder) m_lpFolder->Release();
	if (m_lpMDB) m_lpMDB->Release();
	if (m_lpSession) m_lpSession->Release();
}

//---------------------------------------------------------------------------------//

void CMAPIProcessor::InitSession(LPMAPISESSION lpSession)
{
	if (m_lpSession) m_lpSession->Release();
	m_lpSession = lpSession;
	if (m_lpSession) m_lpSession->AddRef();
}

void CMAPIProcessor::InitMDB(LPMDB lpMDB)
{
	if (m_lpMDB) m_lpMDB->Release();
	m_lpMDB = lpMDB;
	if (m_lpMDB) m_lpMDB->AddRef();
}

void CMAPIProcessor::InitFolder(LPMAPIFOLDER lpFolder)
{
	if (m_lpFolder) m_lpFolder->Release();
	m_lpFolder = lpFolder;
	if (m_lpFolder) m_lpFolder->AddRef();
	MAPIFreeBuffer(m_szFolderOffset);
	HRESULT hRes = S_OK;
	WC_H(CopyString(&m_szFolderOffset,_T("\\"),NULL));// STRING_OK
}

//---------------------------------------------------------------------------------//

//Server name MUST be passed
void CMAPIProcessor::ProcessMailboxTable(
	LPCTSTR szExchangeServerName)
{
	if (!szExchangeServerName) return;
	HRESULT		hRes			= S_OK;

	LPMAPITABLE	lpMailBoxTable	= NULL;
	LPSRowSet	lpRows			= NULL;
	LPMDB		lpPrimaryMDB	= NULL;
	ULONG		ulOffset		= 0;
	ULONG		ulRowNum		= 0;
	ULONG		ulFlags			= LOGOFF_NO_WAIT;

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
			WC_H(lpMailBoxTable->SetColumns((LPSPropTagArray)&sptMBXCols,NULL));
			hRes = S_OK;
			//go to the first row
			WC_H(lpMailBoxTable->SeekRow(
				BOOKMARK_BEGINNING,
				0,
				NULL));

			//get each row in turn and process it
			if (!FAILED(hRes)) for (ulRowNum=0;;ulRowNum++)
			{
				if (lpRows) FreeProws(lpRows);
				lpRows = NULL;

				hRes = S_OK;
				WC_H(lpMailBoxTable->QueryRows(
					1,
					NULL,
					&lpRows));
				if (FAILED(hRes) || !lpRows || !lpRows->cRows) break;

				DoMailboxTablePerRowWork(lpPrimaryMDB,lpRows->aRow,ulRowNum+ulOffset+1);

				LPSPropValue lpEmailAddress = PpropFindProp(lpRows->aRow->lpProps,lpRows->aRow->cValues,PR_EMAIL_ADDRESS);

				if (!CheckStringProp(lpEmailAddress,PT_TSTRING)) continue;

				if (m_lpMDB)
				{
					WC_H(m_lpMDB->StoreLogoff(&ulFlags));
					m_lpMDB->Release();
					m_lpMDB = NULL;
					hRes = S_OK;
				}

				WC_H(OpenOtherUsersMailbox(
					m_lpSession,
					lpPrimaryMDB,
					szExchangeServerName,
					lpEmailAddress->Value.LPSZ,
					true,
					&m_lpMDB));

				if (m_lpMDB) ProcessStore();
			}//loop

			if (lpRows) FreeProws(lpRows);
			lpRows = NULL;
			lpMailBoxTable->Release();
			ulOffset += ulRowNum;
		}
	} while (ulRowNum);

	if (m_lpMDB)
	{
		WC_H(m_lpMDB->StoreLogoff(&ulFlags));
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

	AddFolderToFolderList(NULL,_T("\\")); // STRING_OK

	ProcessFolders(true,true,true);

	EndStoreWork();
}

void CMAPIProcessor::ProcessFolders(BOOL bDoRegular, BOOL bDoAssociated, BOOL bDoDescent)
{
	BeginProcessFoldersWork();

	if (!m_lpFolder)
	{
		OpenFirstFolderInList();
	}

	while (m_lpFolder)
	{
		DoProcessFoldersPerFolderWork();
		ProcessFolder(bDoRegular, bDoAssociated, bDoDescent);
		OpenFirstFolderInList();
	}
	EndProcessFoldersWork();
}

void CMAPIProcessor::ProcessFolder(BOOL bDoRegular,
									  BOOL bDoAssociated,
									  BOOL bDoDescent)
{
	if (!m_lpMDB || !m_lpFolder) return;

    HRESULT         hRes = S_OK;

	BeginFolderWork();

	if (bDoRegular)
	{
		ProcessContentsTable(NULL);
	}

	if (bDoAssociated)
	{
		ProcessContentsTable(MAPI_ASSOCIATED);
	}

	//If we're not processing subfolders, then get outta here
	if (bDoDescent)
	{
	    LPMAPITABLE		lpHierarchyTable	= NULL;
		//We need to walk down the tree
		//and get the list of kids of the folder
		WC_H(m_lpFolder->GetHierarchyTable(fMapiUnicode,&lpHierarchyTable));
		if (lpHierarchyTable)
		{
			enum{NAME,EID,SUBFOLDERS,FLAGS,NUMCOLS};
			SizedSPropTagArray(NUMCOLS,sptHierarchyCols) = {NUMCOLS,
				PR_DISPLAY_NAME,
				PR_ENTRYID,
				PR_SUBFOLDERS,
				PR_CONTAINER_FLAGS};

			LPSRowSet lpRows = NULL;
			//If I don't do this, the MSPST provider can blow chunks (MAPI_E_EXTENDED_ERROR) for some folders when I get a row
			//For some reason, this fixes it.
			WC_H(lpHierarchyTable->SetColumns(
				(LPSPropTagArray) &sptHierarchyCols,
				TBL_BATCH));

			//go to the first row
			WC_H(lpHierarchyTable->SeekRow(
				BOOKMARK_BEGINNING,
				0,
				NULL));

			if (S_OK == hRes) for (;;)
			{
				hRes = S_OK;
				if (lpRows) FreeProws(lpRows);
				lpRows = NULL;
				WC_H(lpHierarchyTable->QueryRows(
					1,
					NULL,
					&lpRows));
				if (FAILED(hRes))
				{
					hRes = S_OK;
					break;
				}
				if (!lpRows || !lpRows->cRows) break;

				DoFolderPerHierarchyTableRowWork(lpRows->aRow);
				if (lpRows->aRow[0].lpProps)
				{
					TCHAR	szSubFolderOffset[MAX_PATH];//Holds subfolder name

					LPSPropValue	lpFolderDisplayName = NULL;

					lpFolderDisplayName = PpropFindProp(
						lpRows->aRow[0].lpProps,
						lpRows->aRow[0].cValues,
						PR_DISPLAY_NAME);

					if (CheckStringProp(lpFolderDisplayName,PT_TSTRING))
					{
						TCHAR szTemp[MAX_PATH/2];
						//Clean up the folder name before appending it to the offset
						WC_H(SanitizeFileName(szTemp,CCH(szTemp),lpFolderDisplayName->Value.LPSZ,CCH(szTemp)));

						WC_H(StringCchPrintf(szSubFolderOffset,CCH(szSubFolderOffset),
							_T("%s%s\\"),// STRING_OK
							m_szFolderOffset,szTemp));
					}
					else
					{
						WC_H(StringCchPrintf(szSubFolderOffset,CCH(szSubFolderOffset),
							_T("%s\\UnknownFolder\\"), // STRING_OK
							m_szFolderOffset));
					}

					AddFolderToFolderList(&lpRows->aRow[0].lpProps[EID].Value.bin,szSubFolderOffset);
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

	HRESULT		hRes = S_OK;

	LPMAPITABLE lpContentsTable = NULL;

	WC_H(m_lpFolder->GetContentsTable(
		ulFlags | fMapiUnicode,
		&lpContentsTable));
	if (lpContentsTable)
	{
		ULONG ulCountRows = 0;
		WC_H(lpContentsTable->GetRowCount(0,&ulCountRows));

		BeginContentsTableWork(ulFlags,ulCountRows);

		LPSRowSet lpRows = NULL;
		ULONG i = 0;
		for (i = 0;;i++)
		{
			hRes = S_OK;
			if (lpRows) FreeProws(lpRows);
			lpRows = NULL;
			WC_H(lpContentsTable->QueryRows(
				1,
				NULL,
				&lpRows));
			if (FAILED(hRes))
			{
				hRes = S_OK;
				break;
			}
			if (!lpRows || !lpRows->cRows) break;

			LPSPropValue	lpMsgSubject = NULL;
			LPSPropValue	lpMsgEID = NULL;

			DoContentsTablePerRowWork(lpRows->aRow, i);

			lpMsgSubject = PpropFindProp(
				lpRows->aRow->lpProps,
				lpRows->aRow->cValues,
				PR_SUBJECT);

			lpMsgEID = PpropFindProp(
				lpRows->aRow->lpProps,
				lpRows->aRow->cValues,
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
				ProcessMessage(lpMessage, NULL);
				lpMessage->Release();
			}

			lpMessage = NULL;
		}

		if (lpRows) FreeProws(lpRows);
		lpContentsTable->Release();
	}
	EndContentsTableWork();
}

void CMAPIProcessor::ProcessMessage(LPMESSAGE lpMessage, LPVOID lpParentMessageData)
{
	if (!lpMessage) return;

	LPVOID lpData = NULL;
	BeginMessageWork(lpMessage, lpParentMessageData, &lpData);
	ProcessRecipients(lpMessage,lpData);
	ProcessAttachments(lpMessage,lpData);
	EndMessageWork(lpMessage,lpData);
}

void CMAPIProcessor::ProcessRecipients(LPMESSAGE lpMessage, LPVOID lpData)
{
	if (!lpMessage) return;

	HRESULT			hRes = S_OK;
	LPMAPITABLE		lpRecipTable = NULL;
	LPSRowSet		lpRows = NULL;

	BeginRecipientWork(lpMessage,lpData);

	//get the recipient table
	WC_H(lpMessage->GetRecipientTable(
		NULL,
		&lpRecipTable));
	if (lpRecipTable)
	{
		ULONG i = 0;
		for (i = 0;;i++)
		{
			hRes = S_OK;
			if (lpRows) FreeProws(lpRows);
			lpRows = NULL;
			WC_H(lpRecipTable->QueryRows(
				1,
				NULL,
				&lpRows));
			if (FAILED(hRes))
			{
				hRes = S_OK;
				break;
			}
			if (!lpRows || !lpRows->cRows) break;

			DoMessagePerRecipientWork(lpMessage,lpData,lpRows->aRow,i);
		}

		if (lpRows) FreeProws(lpRows);
		if (lpRecipTable) lpRecipTable->Release();
	}

	EndRecipientWork(lpMessage,lpData);
}

void CMAPIProcessor::ProcessAttachments(LPMESSAGE lpMessage, LPVOID lpData)
{
	if (!lpMessage) return;

	HRESULT			hRes = S_OK;
	LPMAPITABLE		lpAttachTable = NULL;
	LPSRowSet		lpRows = NULL;

	BeginAttachmentWork(lpMessage,lpData);

	//get the attachment table
	WC_H(lpMessage->GetAttachmentTable(
		NULL,
		&lpAttachTable));

	if (lpAttachTable)
	{
		ULONG i = 0;
		for (i = 0;;i++)
		{
			hRes = S_OK;
			if (lpRows) FreeProws(lpRows);
			lpRows = NULL;
			WC_H(lpAttachTable->QueryRows(
				1,
				NULL,
				&lpRows));
			if (FAILED(hRes))
			{
				hRes = S_OK;
				break;
			}
			if (!lpRows || !lpRows->cRows) break;

			LPSPropValue lpAttachNum = NULL;
			lpAttachNum = PpropFindProp(
				lpRows->aRow->lpProps,
				lpRows->aRow->cValues,
				PR_ATTACH_NUM);

			if (lpAttachNum)
			{
				LPATTACH lpAttach = NULL;
				WC_H(lpMessage->OpenAttach(
					lpAttachNum->Value.l,
					NULL,
					MAPI_BEST_ACCESS,
					(LPATTACH*)&lpAttach));
				hRes = S_OK;

				DoMessagePerAttachmentWork(lpMessage,lpData,lpRows->aRow,lpAttach,i);
				//Check if the attachment is an embedded message - if it is, parse it as such!
				LPSPropValue lpAttachMethod = NULL;

				lpAttachMethod = PpropFindProp(
					lpRows->aRow->lpProps,
					lpRows->aRow->cValues,
					PR_ATTACH_METHOD);

				if (lpAttachMethod && ATTACH_EMBEDDED_MSG == lpAttachMethod->Value.l)
				{
					LPMESSAGE lpAttachMsg = NULL;

					WC_H(lpAttach->OpenProperty(
						PR_ATTACH_DATA_OBJ,
						(LPIID)&IID_IMessage,
						0,
						NULL,//MAPI_MODIFY,
						(LPUNKNOWN *)&lpAttachMsg));
					if (lpAttachMsg)
					{
						ProcessMessage(lpAttachMsg,lpData);
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
		if (lpRows) FreeProws(lpRows);
		lpAttachTable->Release();
	}

	EndAttachmentWork(lpMessage,lpData);
}

//---------------------------------------------------------------------------------//
//List Functions
//---------------------------------------------------------------------------------//
void CMAPIProcessor::AddFolderToFolderList(LPSBinary lpFolderEID, LPCTSTR szFolderOffsetPath)
{
	HRESULT hRes = S_OK;
	LPFOLDERNODE lpNewNode = NULL;
	WC_H(MAPIAllocateBuffer(sizeof(FolderNode),(LPVOID*) &lpNewNode));
	lpNewNode->lpNextFolder = NULL;

	lpNewNode->lpFolderEID = NULL;
	if (lpFolderEID)
	{
		WC_H(MAPIAllocateMore(
			(ULONG)sizeof(SBinary),
			lpNewNode,
			(LPVOID*) &lpNewNode->lpFolderEID));
		WC_H(CopySBinary(lpNewNode->lpFolderEID,lpFolderEID,lpNewNode));
	}

	WC_H(CopyString(&lpNewNode->szFolderOffsetPath,szFolderOffsetPath, lpNewNode));

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

//Call OpenEntry on the first folder in the list, remove it from the list
void CMAPIProcessor::OpenFirstFolderInList()
{
	MAPIFreeBuffer(m_szFolderOffset);
	m_szFolderOffset = NULL;

	if (m_lpFolder) m_lpFolder->Release();
	m_lpFolder = NULL;

	if (!m_lpListHead) return;
	SBinary sBin = {0};

	if (m_lpListHead->lpFolderEID)
	{
		sBin.cb = m_lpListHead->lpFolderEID->cb;
		sBin.lpb = m_lpListHead->lpFolderEID->lpb;
	}

	HRESULT hRes = S_OK;
	if (m_lpListHead->szFolderOffsetPath)
	{
		WC_H(CopyString(&m_szFolderOffset,m_lpListHead->szFolderOffsetPath,NULL));
	}

	LPMAPIFOLDER lpFolder = NULL;
	WC_H(CallOpenEntry(
		m_lpMDB,
		NULL,
		NULL,
		NULL,
		sBin.cb,
		(LPENTRYID) sBin.lpb,
		NULL,
		MAPI_BEST_ACCESS,
		NULL,
		(LPUNKNOWN*)&lpFolder));

	LPFOLDERNODE lpTempNode = NULL;
	lpTempNode = m_lpListHead;
	m_lpListHead = m_lpListHead->lpNextFolder;
	if (m_lpListHead == lpTempNode) m_lpListTail = NULL;

	MAPIFreeBuffer(lpTempNode);

	m_lpFolder = lpFolder;
}

//Clean up the list
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


//---------------------------------------------------------------------------------//
//Worker Functions
//---------------------------------------------------------------------------------//

void CMAPIProcessor::BeginMailboxTableWork(LPCTSTR /*szExchangeServerName*/)
{
}

void CMAPIProcessor::DoMailboxTablePerRowWork(LPMDB /*lpMDB*/, LPSRow /*lpSRow*/, ULONG /*ulCurRow*/)
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

void CMAPIProcessor::DoFolderPerHierarchyTableRowWork(LPSRow /*lpSRow*/)
{
}

void CMAPIProcessor::EndFolderWork()
{
}

void CMAPIProcessor::BeginContentsTableWork(ULONG /*ulFlags*/, ULONG /*ulCountRows*/)
{
}

void CMAPIProcessor::DoContentsTablePerRowWork(LPSRow /*lpSRow*/, ULONG /*ulCurRow*/)
{
}

void CMAPIProcessor::EndContentsTableWork()
{
}

void CMAPIProcessor::BeginMessageWork(LPMESSAGE /*lpMessage*/,LPVOID /*lpParentMessageData*/, LPVOID* /*lpData*/)
{
}

void CMAPIProcessor::BeginRecipientWork(LPMESSAGE /*lpMessage*/,LPVOID /*lpData*/)
{
}

void CMAPIProcessor::DoMessagePerRecipientWork(LPMESSAGE /*lpMessage*/, LPVOID /*lpData*/, LPSRow /*lpSRow*/, ULONG /*ulCurRow*/)
{
}

void CMAPIProcessor::EndRecipientWork(LPMESSAGE /*lpMessage*/,LPVOID /*lpData*/)
{
}

void CMAPIProcessor::BeginAttachmentWork(LPMESSAGE /*lpMessage*/,LPVOID /*lpData*/)
{
}

void CMAPIProcessor::DoMessagePerAttachmentWork(LPMESSAGE /*lpMessage*/, LPVOID /*lpData*/, LPSRow /*lpSRow*/, LPATTACH /*lpAttach*/, ULONG /*ulCurRow*/)
{
}

void CMAPIProcessor::EndAttachmentWork(LPMESSAGE /*lpMessage*/,LPVOID /*lpData*/)
{
}

void CMAPIProcessor::EndMessageWork(LPMESSAGE /*lpMessage*/,LPVOID /*lpData*/)
{
}
