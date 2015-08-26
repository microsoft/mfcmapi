// AttachmentsDlg.cpp : implementation file
// Displays the attachment table for a message

#include "stdafx.h"
#include "AttachmentsDlg.h"
#include "ContentsTableListCtrl.h"
#include "File.h"
#include "FileDialogEx.h"
#include "MapiObjects.h"
#include "ColumnTags.h"
#include "SingleMAPIPropListCtrl.h"
#include "MAPIProgress.h"
#include "MFCUtilityFunctions.h"
#include "ImportProcs.h"

static wstring CLASS = L"CAttachmentsDlg";

/////////////////////////////////////////////////////////////////////////////
// CAttachmentsDlg dialog

CAttachmentsDlg::CAttachmentsDlg(
	_In_ CParentWnd* pParentWnd,
	_In_ CMapiObjects* lpMapiObjects,
	_In_ LPMAPITABLE lpMAPITable,
	_In_ LPMESSAGE lpMessage
	) :
	CContentsTableDlg(
	pParentWnd,
	lpMapiObjects,
	IDS_ATTACHMENTS,
	mfcmapiDO_NOT_CALL_CREATE_DIALOG,
	lpMAPITable,
	(LPSPropTagArray)&sptATTACHCols,
	NUMATTACHCOLUMNS,
	ATTACHColumns,
	IDR_MENU_ATTACHMENTS_POPUP,
	MENU_CONTEXT_ATTACHMENT_TABLE)
{
	TRACE_CONSTRUCTOR(CLASS);
	m_lpMessage = lpMessage;
	if (m_lpMessage) m_lpMessage->AddRef();

	m_bDisplayAttachAsEmbeddedMessage = false;
	m_lpAttach = NULL;
	m_ulAttachNum = (ULONG)-1;

	CreateDialogAndMenu(IDR_MENU_ATTACHMENTS);
} // CAttachmentsDlg::CAttachmentsDlg

CAttachmentsDlg::~CAttachmentsDlg()
{
	TRACE_DESTRUCTOR(CLASS);
	if (m_lpAttach) m_lpAttach->Release();
	if (m_lpMessage)
	{
		HRESULT hRes = S_OK;

		WC_MAPI(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
		m_lpMessage->Release();
	}
} // CAttachmentsDlg::~CAttachmentsDlg

BEGIN_MESSAGE_MAP(CAttachmentsDlg, CContentsTableDlg)
	ON_COMMAND(ID_DELETESELECTEDITEM, OnDeleteSelectedItem)
	ON_COMMAND(ID_MODIFYSELECTEDITEM, OnModifySelectedItem)
	ON_COMMAND(ID_SAVECHANGES, OnSaveChanges)
	ON_COMMAND(ID_SAVETOFILE, OnSaveToFile)
	ON_COMMAND(ID_VIEWEMBEDDEDMESSAGEPROPERTIES, OnViewEmbeddedMessageProps)
	ON_COMMAND(ID_ADDATTACHMENT, OnAddAttachment)
END_MESSAGE_MAP()

void CAttachmentsDlg::OnInitMenu(_In_ CMenu* pMenu)
{
	if (pMenu)
	{
		if (m_lpContentsTableListCtrl)
		{
			int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
			if (m_lpMapiObjects)
			{
				ULONG ulStatus = m_lpMapiObjects->GetBufferStatus();
				pMenu->EnableMenuItem(ID_PASTE, DIM(ulStatus & BUFFER_ATTACHMENTS));
			}
			pMenu->EnableMenuItem(ID_COPY, DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_DELETESELECTEDITEM, DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_MODIFYSELECTEDITEM, DIMMSOK(1 == iNumSel));
			pMenu->EnableMenuItem(ID_SAVETOFILE, DIMMSOK(iNumSel));
			pMenu->EnableMenuItem(ID_DISPLAYSELECTEDITEM, DIM(1 == iNumSel));
		}
		pMenu->CheckMenuItem(ID_VIEWEMBEDDEDMESSAGEPROPERTIES, CHECK(m_bDisplayAttachAsEmbeddedMessage));
	}
	CContentsTableDlg::OnInitMenu(pMenu);
} // CAttachmentsDlg::OnInitMenu

void CAttachmentsDlg::OnDisplayItem()
{
	HRESULT hRes = S_OK;
	int iItem = -1;
	SortListData* lpListData = NULL;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl || !m_lpMessage) return;
	if (!m_lpAttach) return;

	lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
	if (lpListData && ATTACH_EMBEDDED_MSG == lpListData->data.Contents.ulAttachMethod)
	{
		LPMESSAGE lpMessage = OpenEmbeddedMessage();
		if (lpMessage)
		{
			WC_H(DisplayObject(lpMessage, MAPI_MESSAGE, otDefault, this));
			lpMessage->Release();
			lpMessage = NULL;
		}
	}
}

_Check_return_ LPATTACH CAttachmentsDlg::OpenAttach(ULONG ulAttachNum)
{
	HRESULT hRes = S_OK;
	LPATTACH lpAttach = NULL;

	WC_MAPI(m_lpMessage->OpenAttach(
		ulAttachNum,
		NULL,
		MAPI_MODIFY,
		&lpAttach));
	if (MAPI_E_NO_ACCESS == hRes)
	{
		hRes = S_OK;
		WC_MAPI(m_lpMessage->OpenAttach(
			ulAttachNum,
			NULL,
			MAPI_BEST_ACCESS,
			&lpAttach));
	}

	return lpAttach;
}

_Check_return_ LPMESSAGE CAttachmentsDlg::OpenEmbeddedMessage()
{
	if (!m_lpAttach) return NULL;
	HRESULT hRes = S_OK;

	LPMESSAGE lpMessage = NULL;
	WC_MAPI(m_lpAttach->OpenProperty(
		PR_ATTACH_DATA_OBJ,
		(LPIID)&IID_IMessage,
		0,
		MAPI_MODIFY,
		(LPUNKNOWN *)&lpMessage));
	if (hRes == MAPI_E_NO_ACCESS)
	{
		hRes = S_OK;
		WC_MAPI(m_lpAttach->OpenProperty(
			PR_ATTACH_DATA_OBJ,
			(LPIID)&IID_IMessage,
			0,
			MAPI_BEST_ACCESS,
			(LPUNKNOWN *)&lpMessage));
	}
	if (hRes == MAPI_E_INTERFACE_NOT_SUPPORTED ||
		hRes == MAPI_E_NOT_FOUND)
	{
		WARNHRESMSG(hRes, IDS_ATTNOTEMBEDDEDMSG);
	}

	return lpMessage;
}

_Check_return_ HRESULT CAttachmentsDlg::OpenItemProp(
	int iSelectedItem,
	__mfcmapiModifyEnum /*bModify*/,
	_Deref_out_opt_ LPMAPIPROP* lppMAPIProp)
{
	HRESULT hRes = S_OK;
	SortListData* lpListData = NULL;

	DebugPrintEx(DBGOpenItemProp, CLASS, L"OpenItemProp", L"iSelectedItem = 0x%X\n", iSelectedItem);

	if (!m_lpContentsTableListCtrl || !lppMAPIProp) return MAPI_E_INVALID_PARAMETER;

	*lppMAPIProp = NULL;

	// Find the highlighted item AttachNum
	lpListData = (SortListData*)m_lpContentsTableListCtrl->GetItemData(iSelectedItem);

	if (lpListData)
	{
		ULONG ulAttachNum = 0;
		ULONG ulAttachMethod = 0;
		ulAttachNum = lpListData->data.Contents.ulAttachNum;
		ulAttachMethod = lpListData->data.Contents.ulAttachMethod;

		// Check for matching cached attachment to avoid reopen
		if (ulAttachNum != m_ulAttachNum || !m_lpAttach)
		{
			if (m_lpAttach) m_lpAttach->Release();
			m_lpAttach = OpenAttach(ulAttachNum);
			m_ulAttachNum = (ULONG)-1;
			if (m_lpAttach)
			{
				m_ulAttachNum = ulAttachNum;
			}
		}

		if (m_lpAttach && m_bDisplayAttachAsEmbeddedMessage && ATTACH_EMBEDDED_MSG == ulAttachMethod)
		{
			// Reopening an embedded message can fail
			// The view might be holding the embedded message we're trying to open, so we clear
			// it from the view to allow us to reopen it.
			// TODO: Consider caching our embedded message so this isn't necessary
			OnUpdateSingleMAPIPropListCtrl(NULL, NULL);
			*lppMAPIProp = OpenEmbeddedMessage();
		}
		else
		{
			*lppMAPIProp = m_lpAttach;
			if (*lppMAPIProp) (*lppMAPIProp)->AddRef();
		}
	}

	return hRes;
} // CAttachmentsDlg::OpenItemProp

void CAttachmentsDlg::HandleCopy()
{
	if (!m_lpContentsTableListCtrl || !m_lpMessage) return;
	HRESULT hRes = S_OK;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric, CLASS, L"HandleCopy", L"\n");
	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

	ULONG*			lpAttNumList = NULL;
	SortListData*	lpListData = NULL;

	ULONG ulNumSelected = m_lpContentsTableListCtrl->GetSelectedCount();

	if (ulNumSelected && ulNumSelected < ULONG_MAX / sizeof(ULONG))
	{
		EC_H(MAPIAllocateBuffer(
			ulNumSelected * sizeof(ULONG),
			(LPVOID*)&lpAttNumList));
		if (lpAttNumList)
		{
			ZeroMemory(lpAttNumList, ulNumSelected * sizeof(ULONG));
			ULONG ulSelection = 0;
			for (ulSelection = 0; ulSelection < ulNumSelected; ulSelection++)
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
} // CAttachmentsDlg::HandleCopy

_Check_return_ bool CAttachmentsDlg::HandlePaste()
{
	if (CBaseDialog::HandlePaste()) return true;

	if (!m_lpContentsTableListCtrl || !m_lpMessage || !m_lpMapiObjects) return false;
	DebugPrintEx(DBGGeneric, CLASS, L"HandlePaste", L"\n");

	HRESULT		hRes = S_OK;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	ULONG ulStatus = m_lpMapiObjects->GetBufferStatus();
	if (!(ulStatus & BUFFER_ATTACHMENTS) || !(ulStatus & BUFFER_SOURCEPROPOBJ)) return false;

	ULONG* lpAttNumList = m_lpMapiObjects->GetAttachmentsToCopy();
	ULONG iNumSelected = m_lpMapiObjects->GetNumAttachments();
	LPMESSAGE lpSourceMessage = (LPMESSAGE)m_lpMapiObjects->GetSourcePropObject();

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
			EC_MAPI(lpSourceMessage->OpenAttach(
				lpAttNumList[ulAtt],
				NULL,
				MAPI_DEFERRED_ERRORS,
				&lpAttSrc));

			if (lpAttSrc)
			{
				ULONG ulAttNum = NULL;
				// Create the attachment destination
				EC_MAPI(m_lpMessage->CreateAttach(NULL, MAPI_DEFERRED_ERRORS, &ulAttNum, &lpAttDst));
				if (lpAttDst)
				{
					LPMAPIPROGRESS lpProgress = GetMAPIProgress(_T("IAttach::CopyTo"), m_hWnd); // STRING_OK

					// Copy from source to destination
					EC_MAPI(lpAttSrc->CopyTo(
						0,
						NULL,
						0,
						lpProgress ? (ULONG_PTR)m_hWnd : NULL,
						lpProgress,
						(LPIID)&IID_IAttachment,
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
				EC_MAPI(lpAttDst->SaveChanges(KEEP_OPEN_READWRITE));
				lpAttDst->Release();
				lpAttDst = NULL;
			}
		}
		EC_MAPI(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
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

	if (iNumSelected && iNumSelected < ULONG_MAX / sizeof(ULONG))
	{
		EC_H(MAPIAllocateBuffer(
			iNumSelected * sizeof(ULONG),
			(LPVOID*)&lpAttNumList));
		if (lpAttNumList)
		{
			ZeroMemory(lpAttNumList, iNumSelected * sizeof(ULONG));
			int iSelection = 0;
			for (iSelection = 0; iSelection < iNumSelected; iSelection++)
			{
				lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
				if (lpListData)
				{
					lpAttNumList[iSelection] = lpListData->data.Contents.ulAttachNum;
				}
			}

			for (iSelection = 0; iSelection < iNumSelected; iSelection++)
			{
				DebugPrintEx(DBGDeleteSelectedItem, CLASS, L"OnDeleteSelectedItem", L"Deleting attachment 0x%08X\n", lpAttNumList[iSelection]);

				LPMAPIPROGRESS lpProgress = GetMAPIProgress(_T("IMessage::DeleteAttach"), m_hWnd); // STRING_OK

				EC_MAPI(m_lpMessage->DeleteAttach(
					lpAttNumList[iSelection],
					lpProgress ? (ULONG_PTR)m_hWnd : NULL,
					lpProgress,
					lpProgress ? ATTACH_DIALOG : 0));

				if (lpProgress)
					lpProgress->Release();

				lpProgress = NULL;
			}

			MAPIFreeBuffer(lpAttNumList);
			EC_MAPI(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
			OnRefreshView(); // Update the view since we don't have notifications here.
		}
	}
} // CAttachmentsDlg::OnDeleteSelectedItem

void CAttachmentsDlg::OnModifySelectedItem()
{
	if (m_lpAttach)
	{
		HRESULT hRes = S_OK;

		EC_MAPI(m_lpAttach->SaveChanges(KEEP_OPEN_READWRITE));
	}
} // CAttachmentsDlg::OnModifySelectedItem

void CAttachmentsDlg::OnSaveChanges()
{
	if (m_lpMessage)
	{
		HRESULT hRes = S_OK;

		EC_MAPI(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
	}
} // CAttachmentsDlg::OnSaveChanges

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
			if (bShouldCancel(this, hRes)) break;
			hRes = S_OK;
		}
		if (lpListData)
		{
			ulAttachNum = lpListData->data.Contents.ulAttachNum;

			EC_MAPI(m_lpMessage->OpenAttach(
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
	} while (iItem != -1);
} // CAttachmentsDlg::OnSaveToFile

void CAttachmentsDlg::OnViewEmbeddedMessageProps()
{
	m_bDisplayAttachAsEmbeddedMessage = !m_bDisplayAttachAsEmbeddedMessage;
	OnRefreshView();
} // CAttachmentsDlg::OnViewEmbeddedMessageProps

void CAttachmentsDlg::OnAddAttachment()
{
	HRESULT hRes = 0;
	INT_PTR iDlgRet = 0;
	CStringW szFileSpec;
	EC_B(szFileSpec.LoadString(IDS_ALLFILES));

	CFileDialogExW dlgFilePicker;

	EC_D_DIALOG(dlgFilePicker.DisplayDialog(
		true,
		NULL,
		NULL,
		NULL,
		szFileSpec));
	if (iDlgRet == IDOK)
	{
		LPATTACH lpAttachment = NULL;
		ULONG ulAttachNum = 0;

		EC_MAPI(m_lpMessage->CreateAttach(NULL, NULL, &ulAttachNum, &lpAttachment));

		if (SUCCEEDED(hRes) && lpAttachment)
		{
			LPWSTR szAttachName = dlgFilePicker.GetFileName();
			SPropValue spvAttach[4];
			spvAttach[0].ulPropTag = PR_ATTACH_METHOD;
			spvAttach[0].Value.l = ATTACH_BY_VALUE;
			spvAttach[1].ulPropTag = PR_RENDERING_POSITION;
			spvAttach[1].Value.l = -1;
			spvAttach[2].ulPropTag = PR_ATTACH_FILENAME_W;
			spvAttach[2].Value.lpszW = szAttachName;
			spvAttach[3].ulPropTag = PR_DISPLAY_NAME_W;
			spvAttach[3].Value.lpszW = szAttachName;

			EC_MAPI(lpAttachment->SetProps(_countof(spvAttach), spvAttach, NULL));
			if (SUCCEEDED(hRes))
			{
				LPSTREAM pStreamFile = NULL;

				EC_MAPI(MyOpenStreamOnFile(
					MAPIAllocateBuffer,
					MAPIFreeBuffer,
					STGM_READ,
					szAttachName,
					NULL,
					&pStreamFile));
				if (SUCCEEDED(hRes) && pStreamFile)
				{
					LPSTREAM pStreamAtt = NULL;
					STATSTG StatInfo = { 0 };

					EC_MAPI(lpAttachment->OpenProperty(
						PR_ATTACH_DATA_BIN,
						&IID_IStream,
						0,
						MAPI_MODIFY | MAPI_CREATE,
						(LPUNKNOWN *)&pStreamAtt));
					if (SUCCEEDED(hRes) && pStreamAtt)
					{
						EC_MAPI(pStreamFile->Stat(&StatInfo, STATFLAG_NONAME));
						EC_MAPI(pStreamFile->CopyTo(pStreamAtt, StatInfo.cbSize, NULL, NULL));
						EC_MAPI(pStreamAtt->Commit(STGC_DEFAULT));
						EC_MAPI(lpAttachment->SaveChanges(KEEP_OPEN_READWRITE));
						EC_MAPI(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
					}

					if (pStreamAtt) pStreamAtt->Release();
				}

				if (pStreamFile) pStreamFile->Release();
			}
		}

		if (lpAttachment) lpAttachment->Release();

		OnRefreshView();
	}
}

void CAttachmentsDlg::HandleAddInMenuSingle(
	_In_ LPADDINMENUPARAMS lpParams,
	_In_ LPMAPIPROP lpMAPIProp,
	_In_ LPMAPICONTAINER /*lpContainer*/)
{
	if (lpParams)
	{
		lpParams->lpMessage = m_lpMessage;
		lpParams->lpAttach = (LPATTACH)lpMAPIProp; // OpenItemProp returns LPATTACH
	}

	InvokeAddInMenu(lpParams);
} // CAttachmentsDlg::HandleAddInMenuSingle
