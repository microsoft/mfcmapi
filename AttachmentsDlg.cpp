// AttachmentsDlg.cpp : implementation file
// Displays the attachment table for a message

#include "stdafx.h"
#include "Error.h"

#include "AttachmentsDlg.h"

#include "ContentsTableListCtrl.h"
#include "File.h"
#include "MapiObjects.h"
#include "ColumnTags.h"
#include "SingleMAPIPropListCtrl.h"
#include "MAPIProgress.h"
#include "MFCUtilityFunctions.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

static TCHAR* CLASS = _T("CAttachmentsDlg");

/////////////////////////////////////////////////////////////////////////////
// CAttachmentsDlg dialog

CAttachmentsDlg::CAttachmentsDlg(
							   CParentWnd* pParentWnd,
							   CMapiObjects *lpMapiObjects,
							   LPMAPITABLE lpMAPITable,
							   LPMESSAGE lpMessage
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
	m_lpAttach = NULL;

	CreateDialogAndMenu(IDR_MENU_ATTACHMENTS);
}

CAttachmentsDlg::~CAttachmentsDlg()
{
	TRACE_DESTRUCTOR(CLASS);
	if (m_lpAttach) m_lpAttach->Release();
	if (m_lpMessage)
	{
		HRESULT hRes = S_OK;

		EC_H(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
	}
	if (m_lpMessage) m_lpMessage->Release();
}

BEGIN_MESSAGE_MAP(CAttachmentsDlg, CContentsTableDlg)
//{{AFX_MSG_MAP(CAttachmentsDlg)
	ON_COMMAND(ID_DELETESELECTEDITEM, OnDeleteSelectedItem)
	ON_COMMAND(ID_MODIFYSELECTEDITEM, OnModifySelectedItem)
	ON_COMMAND(ID_SAVECHANGES, OnSaveChanges)
	ON_COMMAND(ID_SAVETOFILE, OnSaveToFile)
	ON_COMMAND(ID_VIEWEMBEDDEDMESSAGEPROPERTIES, OnViewEmbeddedMessageProps)
	ON_COMMAND(ID_USEMAPIMODIFYONATTACHMENTS, OnUseMapiModify)
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

void CAttachmentsDlg::OnInitMenu(CMenu* pMenu)
{
	if (pMenu)
	{
		if (m_lpContentsTableListCtrl)
		{
			int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
			pMenu->EnableMenuItem(ID_DELETESELECTEDITEM,DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_MODIFYSELECTEDITEM,DIMMSOK(1 == iNumSel));
			pMenu->EnableMenuItem(ID_SAVETOFILE,DIMMSOK(iNumSel));
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
	ULONG			ulAttachNum = 0;

	DebugPrintEx(DBGOpenItemProp,CLASS,_T("OpenItemProp"),_T("iSelectedItem = 0x%X\n"),iSelectedItem);

	*lppMAPIProp = NULL;
	if (m_lpAttach) m_lpAttach->Release();
	m_lpAttach = NULL;

	if (!m_lpContentsTableListCtrl || !lppMAPIProp) return MAPI_E_INVALID_PARAMETER;

	//Find the highlighted item AttachNum
	lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(NULL);

	if (lpListData)
	{
		ulAttachNum = lpListData->data.Contents.ulAttachNum;

		if (!m_bDisplayAttachAsEmbeddedMessage)
		{
			EC_H(m_lpMessage->OpenAttach(
				ulAttachNum,
				NULL,
				MAPI_BEST_ACCESS,
				(LPATTACH*)&m_lpAttach));
			*lppMAPIProp = m_lpAttach;
			if (*lppMAPIProp) (*lppMAPIProp)->AddRef();
		}
		else
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
	}
	return hRes;
}

void CAttachmentsDlg::OnDeleteSelectedItem()
{
	if (!m_lpContentsTableListCtrl || !m_lpMessage) return;
	HRESULT			hRes = S_OK;
	ULONG*			lpAttNumList = NULL;
	CWaitCursor		Wait;//Change the mouse to an hourglass while we work.
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

				LPMAPIPROGRESS lpProgress = GetMAPIProgress(_T("IMessage::DeleteAttach"), m_hWnd);// STRING_OK

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
			OnRefreshView();//Update the view since we don't have notifications here.
		}
	}
}//CAttachmentsDlg::OnDeleteSelectedItem

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
	CWaitCursor		Wait;//Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl || !m_lpMessage) return;

	do
	{
		//Find the highlighted item AttachNum
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
				MAPI_BEST_ACCESS,//TODO: Is best access really needed?
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
} //CAttachmentsDlg::OnSaveToFile

void CAttachmentsDlg::OnViewEmbeddedMessageProps()
{
	m_bDisplayAttachAsEmbeddedMessage = !m_bDisplayAttachAsEmbeddedMessage;
	OnRefreshView();
}//CAttachmentsDlg::OnViewEmbeddedMessageProps

void CAttachmentsDlg::OnUseMapiModify()
{
	m_bUseMapiModifyOnEmbeddedMessage = !m_bUseMapiModifyOnEmbeddedMessage;
}//CAttachmentsDlg::OnUseMapiModify

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
