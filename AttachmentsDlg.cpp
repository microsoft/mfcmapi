// AttachmentsDlg.cpp : implementation file
// Displays the attachment table for a message

#include "stdafx.h"
#include "AttachmentsDlg.h"
#include "ContentsTableListCtrl.h"
#include "File.h"
#include "MapiObjects.h"
#include "ColumnTags.h"
#include "SingleMAPIPropListCtrl.h"
#include "MAPIProgress.h"
#include "MFCUtilityFunctions.h"

static TCHAR* CLASS = _T("CAttachmentsDlg");

/////////////////////////////////////////////////////////////////////////////
// CAttachmentsDlg dialog

CAttachmentsDlg::CAttachmentsDlg(
								 CParentWnd* pParentWnd,
								 CMapiObjects* lpMapiObjects,
								 LPMAPITABLE lpMAPITable,
								 LPMESSAGE lpMessage,
								 BOOL bSaveMessageAtClose
								 ):
CContentsTableDlg(
				  pParentWnd,
				  lpMapiObjects,
				  IDS_ATTACHMENTS,
				  mfcmapiDO_NOT_CALL_CREATE_DIALOG,
				  lpMAPITable,
				  (LPSPropTagArray) &sptATTACHCols,
				  NUMATTACHCOLUMNS,
				  ATTACHColumns,
				  IDR_MENU_ATTACHMENTS_POPUP,
				  MENU_CONTEXT_ATTACHMENT_TABLE)
{
	TRACE_CONSTRUCTOR(CLASS);
	m_lpMessage = lpMessage;
	if (m_lpMessage) m_lpMessage->AddRef();

	m_bDisplayAttachAsEmbeddedMessage = FALSE;
	m_bUseMapiModifyOnEmbeddedMessage = FALSE;
	m_bSaveMessageAtClose = bSaveMessageAtClose;
	m_lpAttach = NULL;

	CreateDialogAndMenu(IDR_MENU_ATTACHMENTS);
}

CAttachmentsDlg::~CAttachmentsDlg()
{
	TRACE_DESTRUCTOR(CLASS);
	if (m_lpAttach) m_lpAttach->Release();
	if (m_lpMessage && m_bSaveMessageAtClose)
	{
		HRESULT hRes = S_OK;

		EC_H(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
	}
	if (m_lpMessage) m_lpMessage->Release();
}

BEGIN_MESSAGE_MAP(CAttachmentsDlg, CContentsTableDlg)
	ON_COMMAND(ID_DELETESELECTEDITEM, OnDeleteSelectedItem)
	ON_COMMAND(ID_MODIFYSELECTEDITEM, OnModifySelectedItem)
	ON_COMMAND(ID_SAVECHANGES, OnSaveChanges)
	ON_COMMAND(ID_SAVETOFILE, OnSaveToFile)
	ON_COMMAND(ID_VIEWEMBEDDEDMESSAGEPROPERTIES, OnViewEmbeddedMessageProps)
	ON_COMMAND(ID_USEMAPIMODIFYONATTACHMENTS, OnUseMapiModify)
	ON_COMMAND(ID_ATTACHMENTPROPERTIES, OnAttachmentProperties)
	ON_COMMAND(ID_RECIPIENTPROPERTIES, OnRecipientProperties)
END_MESSAGE_MAP()

void CAttachmentsDlg::OnInitMenu(CMenu* pMenu)
{
	if (pMenu)
	{
		if (m_lpContentsTableListCtrl)
		{
			int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
			if (m_lpMapiObjects)
			{
				ULONG ulStatus = m_lpMapiObjects->GetBufferStatus();
				pMenu->EnableMenuItem(ID_PASTE,DIM(ulStatus & BUFFER_ATTACHMENTS));
			}
			pMenu->EnableMenuItem(ID_COPY,DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_DELETESELECTEDITEM,DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_MODIFYSELECTEDITEM,DIMMSOK(1 == iNumSel));
			pMenu->EnableMenuItem(ID_SAVETOFILE,DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_ATTACHMENTPROPERTIES,DIM(m_bDisplayAttachAsEmbeddedMessage && 1 == iNumSel));
			pMenu->EnableMenuItem(ID_RECIPIENTPROPERTIES,DIM(m_bDisplayAttachAsEmbeddedMessage && 1 == iNumSel));
		}
		pMenu->CheckMenuItem(ID_VIEWEMBEDDEDMESSAGEPROPERTIES,CHECK(m_bDisplayAttachAsEmbeddedMessage));
		pMenu->CheckMenuItem(ID_USEMAPIMODIFYONATTACHMENTS,CHECK(m_bUseMapiModifyOnEmbeddedMessage));

	}
	CContentsTableDlg::OnInitMenu(pMenu);
}

HRESULT CAttachmentsDlg::OpenItemProp(
									  int iSelectedItem,
									  __mfcmapiModifyEnum /*bModify*/,
									  LPMAPIPROP* lppMAPIProp)
{
	HRESULT			hRes = S_OK;
	SortListData*	lpListData = NULL;

	DebugPrintEx(DBGOpenItemProp,CLASS,_T("OpenItemProp"),_T("iSelectedItem = 0x%X\n"),iSelectedItem);

	*lppMAPIProp = NULL;
	if (m_lpAttach) m_lpAttach->Release();
	m_lpAttach = NULL;

	if (!m_lpContentsTableListCtrl || !lppMAPIProp) return MAPI_E_INVALID_PARAMETER;

	// Find the highlighted item AttachNum
	lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(NULL);

	if (lpListData)
	{
		ULONG ulAttachNum = 0;
		ULONG ulAttachMethod = 0;
		ulAttachNum = lpListData->data.Contents.ulAttachNum;
		ulAttachMethod = lpListData->data.Contents.ulAttachMethod;

		if (m_bDisplayAttachAsEmbeddedMessage && ATTACH_EMBEDDED_MSG == ulAttachMethod)
		{
			EC_H(m_lpMessage->OpenAttach(
				ulAttachNum,
				NULL,
				m_bUseMapiModifyOnEmbeddedMessage?MAPI_MODIFY:MAPI_BEST_ACCESS,
				(LPATTACH*)&m_lpAttach));

			if (m_lpAttach)
			{
				EC_H(m_lpAttach->OpenProperty(
					PR_ATTACH_DATA_OBJ,
					(LPIID)&IID_IMessage,
					0,
					m_bUseMapiModifyOnEmbeddedMessage?MAPI_MODIFY:0,
					(LPUNKNOWN *)lppMAPIProp));
				if (hRes == MAPI_E_INTERFACE_NOT_SUPPORTED)
				{
					WARNHRESMSG(hRes,IDS_ATTNOTEMBEDDEDMSG);
				}
				else if (hRes == MAPI_E_NOT_FOUND)
				{
					WARNHRESMSG(hRes,IDS_ATTNOTEMBEDDEDMSG);
				}
			}
		}
		else
		{
			EC_H(m_lpMessage->OpenAttach(
				ulAttachNum,
				NULL,
				MAPI_BEST_ACCESS,
				(LPATTACH*)&m_lpAttach));
			*lppMAPIProp = m_lpAttach;
			if (*lppMAPIProp) (*lppMAPIProp)->AddRef();
		}
	}
	return hRes;
}

BOOL CAttachmentsDlg::HandleCopy()
{
	if (!m_lpContentsTableListCtrl || !m_lpMessage) return false;
	HRESULT hRes = S_OK;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric,CLASS,_T("HandleCopy"),_T("\n"));
	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return false;

	ULONG*			lpAttNumList = NULL;
	SortListData*	lpListData = NULL;

	ULONG ulNumSelected = m_lpContentsTableListCtrl->GetSelectedCount();

	if (ulNumSelected && ulNumSelected < ULONG_MAX/sizeof(ULONG))
	{
		EC_H(MAPIAllocateBuffer(
			ulNumSelected * sizeof(ULONG),
			(LPVOID*) &lpAttNumList));
		if (lpAttNumList)
		{
			ZeroMemory(lpAttNumList, ulNumSelected * sizeof(ULONG));
			ULONG ulSelection = 0;
			for (ulSelection = 0 ; ulSelection < ulNumSelected ; ulSelection++)
			{
				int	iItem = -1;
				lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
				if (lpListData)
				{
					lpAttNumList[ulSelection] = lpListData->data.Contents.ulAttachNum;
				}
			}

			// m_lpMapiObjects takes over ownership of lpAttNumList - don't free now
			m_lpMapiObjects->SetAttachmentsToCopy(m_lpMessage, ulNumSelected, lpAttNumList);
		}
	}

	return true;
} // CAttachmentsDlg::HandleCopy

BOOL CAttachmentsDlg::HandlePaste()
{
	if (CBaseDialog::HandlePaste()) return true;

	if (!m_lpContentsTableListCtrl || !m_lpMessage || !m_lpMapiObjects) return false;
	DebugPrintEx(DBGGeneric,CLASS,_T("HandlePaste"),_T("\n"));

	HRESULT		hRes = S_OK;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	ULONG ulStatus = m_lpMapiObjects->GetBufferStatus();
	if (!(ulStatus & BUFFER_ATTACHMENTS) || !(ulStatus & BUFFER_SOURCEPROPOBJ)) return false;

	ULONG* lpAttNumList = m_lpMapiObjects->GetAttachmentsToCopy();
	ULONG iNumSelected = m_lpMapiObjects->GetNumAttachments();
	LPMESSAGE lpSourceMessage = (LPMESSAGE) m_lpMapiObjects->GetSourcePropObject();

	if (lpAttNumList && iNumSelected && lpSourceMessage)
	{
		// If we failed on one pass, try the rest
		hRes = S_OK;
		// Go through each attachment and copy it
		ULONG ulAtt = 0;
		for (ulAtt = 0; ulAtt < iNumSelected; ++ulAtt)
		{
			LPATTACH lpAttSrc = NULL;
			LPATTACH lpAttDst = NULL;
			LPSPropProblemArray lpProblems = NULL;

			// Open the attachment source
			EC_H(lpSourceMessage->OpenAttach(
				lpAttNumList[ulAtt],
				NULL,
				MAPI_DEFERRED_ERRORS,
				&lpAttSrc));

			if (lpAttSrc)
			{
				ULONG ulAttNum = NULL;
				// Create the attachment destination
				EC_H(m_lpMessage->CreateAttach(NULL, MAPI_DEFERRED_ERRORS, &ulAttNum, &lpAttDst));
				if (lpAttDst)
				{
					LPMAPIPROGRESS lpProgress = GetMAPIProgress(_T("IAttach::CopyTo"), m_hWnd); // STRING_OK

					// Copy from source to destination
					EC_H(lpAttSrc->CopyTo(
						0,
						NULL,
						0,
						lpProgress ? (ULONG_PTR)m_hWnd : NULL,
						lpProgress,
						(LPIID) &IID_IAttachment,
						lpAttDst,
						lpProgress ? MAPI_DIALOG : 0,
						&lpProblems));

					if (lpProgress) lpProgress->Release();
					lpProgress = NULL;

					EC_PROBLEMARRAY(lpProblems);
					MAPIFreeBuffer(lpProblems);
				}
			}

			if (lpAttSrc) lpAttSrc->Release();
			lpAttSrc = NULL;

			if (lpAttDst)
			{
				EC_H(lpAttDst->SaveChanges(KEEP_OPEN_READWRITE));
				lpAttDst->Release();
				lpAttDst = NULL;
			}
		}
		EC_H(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
		OnRefreshView(); // Update the view since we don't have notifications here.
	}

	return true;
} // CAttachmentsDlg::HandlePaste

void CAttachmentsDlg::OnDeleteSelectedItem()
{
	if (!m_lpContentsTableListCtrl || !m_lpMessage) return;
	HRESULT			hRes = S_OK;
	ULONG*			lpAttNumList = NULL;
	CWaitCursor		Wait; // Change the mouse to an hourglass while we work.
	int				iItem = -1;
	SortListData*	lpListData = NULL;

	int iNumSelected = m_lpContentsTableListCtrl->GetSelectedCount();

	if (iNumSelected && iNumSelected < ULONG_MAX/sizeof(ULONG))
	{
		EC_H(MAPIAllocateBuffer(
			iNumSelected * sizeof(ULONG),
			(LPVOID*) &lpAttNumList));
		if (lpAttNumList)
		{
			ZeroMemory(lpAttNumList, iNumSelected * sizeof(ULONG));
			int iSelection = 0;
			for (iSelection = 0 ; iSelection < iNumSelected ; iSelection++)
			{
				lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
				if (lpListData)
				{
					lpAttNumList[iSelection] = lpListData->data.Contents.ulAttachNum;
				}
			}

			for (iSelection = 0 ; iSelection < iNumSelected ; iSelection++)
			{
				DebugPrintEx(DBGDeleteSelectedItem,CLASS,_T("OnDeleteSelectedItem"),_T("Deleting attachment 0x%08X\n"),lpAttNumList[iSelection]);

				LPMAPIPROGRESS lpProgress = GetMAPIProgress(_T("IMessage::DeleteAttach"), m_hWnd); // STRING_OK

				EC_H(m_lpMessage->DeleteAttach(
					lpAttNumList[iSelection],
					lpProgress ? (ULONG_PTR)m_hWnd : NULL,
					lpProgress,
					lpProgress ? ATTACH_DIALOG : 0));

				if(lpProgress)
					lpProgress->Release();

				lpProgress = NULL;
			}

			MAPIFreeBuffer(lpAttNumList);
			EC_H(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
			OnRefreshView(); // Update the view since we don't have notifications here.
		}
	}
} // CAttachmentsDlg::OnDeleteSelectedItem

void CAttachmentsDlg::OnModifySelectedItem()
{
	if (m_lpAttach)
	{
		HRESULT hRes = S_OK;

		EC_H(m_lpAttach->SaveChanges(KEEP_OPEN_READWRITE));
	}
}

void CAttachmentsDlg::OnSaveChanges()
{
	if (m_lpMessage)
	{
		HRESULT hRes = S_OK;

		EC_H(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
	}
}

void CAttachmentsDlg::OnSaveToFile()
{
	HRESULT			hRes = S_OK;
	LPATTACH		lpAttach = NULL;
	ULONG			ulAttachNum = 0;
	int				iItem = -1;
	SortListData*	lpListData = NULL;
	CWaitCursor		Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl || !m_lpMessage) return;

	do
	{
		// Find the highlighted item AttachNum
		lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
		if (S_OK != hRes && -1 != iItem)
		{
			if (bShouldCancel(this,hRes)) break;
			hRes = S_OK;
		}
		if (lpListData)
		{
			ulAttachNum = lpListData->data.Contents.ulAttachNum;

			EC_H(m_lpMessage->OpenAttach(
				ulAttachNum,
				NULL,
				MAPI_BEST_ACCESS, // TODO: Is best access really needed?
				(LPATTACH*)&lpAttach));

			if (lpAttach)
			{
				WC_H(WriteAttachmentToFile(lpAttach, m_hWnd));

				lpAttach->Release();
				lpAttach = NULL;
			}
		}
	}
	while (iItem != -1);
} // CAttachmentsDlg::OnSaveToFile

void CAttachmentsDlg::OnViewEmbeddedMessageProps()
{
	m_bDisplayAttachAsEmbeddedMessage = !m_bDisplayAttachAsEmbeddedMessage;
	OnRefreshView();
} // CAttachmentsDlg::OnViewEmbeddedMessageProps

void CAttachmentsDlg::OnUseMapiModify()
{
	m_bUseMapiModifyOnEmbeddedMessage = !m_bUseMapiModifyOnEmbeddedMessage;
} // CAttachmentsDlg::OnUseMapiModify

HRESULT CAttachmentsDlg::GetEmbeddedMessage(int iIndex, LPMESSAGE *lppMessage)
{
	HRESULT hRes = S_OK;
	LPMAPIPROP lpMAPIProp = NULL;
	__mfcmapiModifyEnum	fRequestModify = m_bUseMapiModifyOnEmbeddedMessage ? mfcmapiREQUEST_MODIFY : mfcmapiDO_NOT_REQUEST_MODIFY;

	if (!m_bDisplayAttachAsEmbeddedMessage) return MAPI_E_CALL_FAILED;

	EC_H(OpenItemProp(
		iIndex,
		fRequestModify,
		&lpMAPIProp));

	if (NULL != lpMAPIProp)
	{
		EC_H(lpMAPIProp->QueryInterface(
			IID_IMessage,
			(LPVOID*)lppMessage));

		lpMAPIProp->Release();
	}

	return hRes;
} // CAttachmentsDlg::GetEmbeddedMessage

void CAttachmentsDlg::OnAttachmentProperties()
{
	HRESULT hRes = S_OK;
	LPMESSAGE lpMessage = NULL;

	EC_H(GetEmbeddedMessage(
		-1,
		&lpMessage));

	if (NULL != lpMessage)
	{
		EC_H(OpenAttachmentsFromMessage(lpMessage, m_bUseMapiModifyOnEmbeddedMessage));

		lpMessage->Release();
	}

} // CAttachmentsDlg::OnAttachmentProperties

void CAttachmentsDlg::OnRecipientProperties()
{
	HRESULT hRes = S_OK;
	LPMESSAGE lpMessage = NULL;

	EC_H(GetEmbeddedMessage(
		-1,
		&lpMessage));

	if (lpMessage)
	{
		EC_H(OpenRecipientsFromMessage(lpMessage));

		lpMessage->Release();
	}

} // CAttachmentsDlg::OnRecipientProperties

void CAttachmentsDlg::HandleAddInMenuSingle(
	LPADDINMENUPARAMS lpParams,
	LPMAPIPROP lpMAPIProp,
	LPMAPICONTAINER /*lpContainer*/)
{
	if (lpParams)
	{
		lpParams->lpMessage = m_lpMessage;
		lpParams->lpAttach = (LPATTACH) lpMAPIProp; // OpenItemProp returns LPATTACH
	}

	InvokeAddInMenu(lpParams);
} // CAttachmentsDlg::HandleAddInMenuSingle
