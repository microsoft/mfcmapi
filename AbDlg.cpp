// AbDlg.cpp : implementation file
// Displays the contents of a single address book

#include "stdafx.h"
#include "ABDlg.h"
#include "ContentsTableListCtrl.h"
#include "MapiObjects.h"
#include "SingleMAPIPropListCtrl.h"
#include "ColumnTags.h"
#include "MFCUtilityFunctions.h"
#include "Editor.h"
#include "MAPIABFunctions.h"
#include "MAPIProgress.h"
#include "MAPIFunctions.h"

static TCHAR* CLASS = _T("CAbDlg");

/////////////////////////////////////////////////////////////////////////////
// CAbDlg dialog


CAbDlg::CAbDlg(
			   CParentWnd* pParentWnd,
			   CMapiObjects* lpMapiObjects,
			   LPABCONT lpAdrBook
			   ):
CContentsTableDlg(
				  pParentWnd,
				  lpMapiObjects,
				  IDS_AB,
				  mfcmapiDO_NOT_CALL_CREATE_DIALOG,
				  NULL,
				  (LPSPropTagArray) &sptABCols,
				  NUMABCOLUMNS,
				  ABColumns,
				  IDR_MENU_AB_VIEW_POPUP,
				  MENU_CONTEXT_AB_CONTENTS)
{
	TRACE_CONSTRUCTOR(CLASS);
	m_lpContainer = lpAdrBook;
	if (m_lpContainer) m_lpContainer->AddRef();

	m_bIsAB = true;

	CreateDialogAndMenu(IDR_MENU_AB_VIEW);
}

CAbDlg::~CAbDlg()
{
	TRACE_DESTRUCTOR(CLASS);
}

BEGIN_MESSAGE_MAP(CAbDlg, CContentsTableDlg)
	ON_COMMAND(ID_DELETESELECTEDITEM, OnDeleteSelectedItem)
	ON_COMMAND(ID_DISPLAYDETAILS, OnDisplayDetails)
	ON_COMMAND(ID_OPENCONTACT, OnOpenContact)
	ON_COMMAND(ID_OPENMANAGER, OnOpenManager)
	ON_COMMAND(ID_OPENOWNER, OnOpenOwner)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CAbDlg message handlers

void CAbDlg::CreateDialogAndMenu(UINT nIDMenuResource)
{
	DebugPrintEx(DBGCreateDialog,CLASS,_T("CreateDialogAndMenu"),_T("id = 0x%X\n"),nIDMenuResource);
	CContentsTableDlg::CreateDialogAndMenu(nIDMenuResource);

	UpdateMenuString(
		this,
		ID_CREATEPROPERTYSTRINGRESTRICTION,
		IDS_ABRESMENU);
} // CAbDlg::CreateDialogAndMenu

void CAbDlg::OnInitMenu(CMenu* pMenu)
{
	if (pMenu && m_lpContentsTableListCtrl)
	{
		if (m_lpMapiObjects)
		{
			ULONG ulStatus = m_lpMapiObjects->GetBufferStatus();
			pMenu->EnableMenuItem(ID_PASTE,DIM(ulStatus & BUFFER_ABENTRIES));
		}

		int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
		pMenu->EnableMenuItem(ID_COPY,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_DELETESELECTEDITEM,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_DISPLAYDETAILS,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_OPENCONTACT,DIMMSNOK(iNumSel));
		pMenu->EnableMenuItem(ID_OPENMANAGER,DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_OPENOWNER,DIMMSOK(iNumSel));
	}
	CContentsTableDlg::OnInitMenu(pMenu);
}

void CAbDlg::OnDisplayDetails()
{
	DebugPrintEx(DBGGeneric,CLASS,_T("OnDisplayDetails"),_T("displaying Address Book entry details\n"));

	HRESULT		hRes = S_OK;
	if (!m_lpMapiObjects) return;
	LPADRBOOK	lpAddrBook = m_lpMapiObjects->GetAddrBook(false); // do not release
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.
	LPENTRYLIST	lpEIDs = NULL;

	if (lpAddrBook)
	{
		EC_H(m_lpContentsTableListCtrl->GetSelectedItemEIDs(&lpEIDs));

		if (lpEIDs && lpEIDs->cValues && lpEIDs->lpbin)
		{
			ULONG i = 0;

			for (i = 0 ; i < lpEIDs->cValues ; i++)
			{
				ULONG_PTR ulUIParam = (ULONG_PTR) (void*) m_hWnd;

				// Have to pass DIALOG_MODAL according to
				// http://support.microsoft.com/kb/171637
				EC_H(lpAddrBook->Details(
					&ulUIParam,
					NULL,
					NULL,
					lpEIDs->lpbin[i].cb,
					(LPENTRYID) lpEIDs->lpbin[i].lpb,
					NULL,
					NULL,
					NULL,
					DIALOG_MODAL)); // API doesn't like Unicode
				if (lpEIDs->cValues > i+1 && bShouldCancel(this,hRes)) break;
				hRes = S_OK;
			}
		}
		MAPIFreeBuffer(lpEIDs);
	}

	return;
}

void CAbDlg::OnOpenContact()
{
	HRESULT			hRes = S_OK;
	LPENTRYLIST		lpEntryList = NULL;
	LPMESSAGE		lpMessage = NULL;

	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl || !m_lpPropDisplay) return;
	LPMAPISESSION	lpSession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpSession) return;

	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	EC_H(m_lpContentsTableListCtrl->GetSelectedItemEIDs(&lpEntryList));

	if (SUCCEEDED(hRes) && lpEntryList && 1 == lpEntryList->cValues && sizeof(CONTAB_ENTRYID) <= lpEntryList->lpbin[0].cb)
	{
		LPCONTAB_ENTRYID lpContabEID = (LPCONTAB_ENTRYID)lpEntryList->lpbin[0].lpb;

		if (lpContabEID && lpContabEID->cbeid && lpContabEID->abeid)
		{
			EC_H(CallOpenEntry(
				NULL,
				NULL,
				NULL,
				lpSession,
				lpContabEID->cbeid,
				(LPENTRYID) lpContabEID->abeid,
				NULL,
				NULL,
				NULL,
				(LPUNKNOWN*)&lpMessage));
		}
	}

	m_lpPropDisplay->SetDataSource(lpMessage, NULL, false);
	if (lpMessage) lpMessage->Release();
	MAPIFreeBuffer(lpEntryList);
	return;
} // CAbDlg::OnOpenContact

void CAbDlg::OnOpenManager()
{
	HRESULT			hRes = S_OK;
	LPMAILUSER		lpMailUser = NULL;
	int				iItem = -1;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

	do
	{
		hRes = S_OK;
		if (lpMailUser) lpMailUser->Release();
		lpMailUser = NULL;
		EC_H(m_lpContentsTableListCtrl->OpenNextSelectedItemProp(
			&iItem,
			mfcmapiREQUEST_MODIFY,
			(LPMAPIPROP*)&lpMailUser));

		if (lpMailUser)
		{
			EC_H(DisplayTable(
				lpMailUser,
				PR_EMS_AB_MANAGER_O,
				otDefault, // oType,
				this));
		}
	}
	while (iItem != -1);

	if (lpMailUser) lpMailUser->Release();
	return;
} // CAbDlg::OnOpenManager

void CAbDlg::OnOpenOwner()
{
	HRESULT			hRes = S_OK;
	LPMAILUSER		lpMailUser = NULL;
	int				iItem = -1;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

	do
	{
		hRes = S_OK;
		if (lpMailUser) lpMailUser->Release();
		lpMailUser = NULL;
		EC_H(m_lpContentsTableListCtrl->OpenNextSelectedItemProp(
			&iItem,
			mfcmapiREQUEST_MODIFY,
			(LPMAPIPROP*)&lpMailUser));

		if (lpMailUser)
		{
			EC_H(DisplayTable(
				lpMailUser,
				PR_EMS_AB_OWNER_O,
				otDefault, // oType,
				this));
		}
	}
	while (iItem != -1);

	if (lpMailUser) lpMailUser->Release();
	return;
} // CAbDlg::OnOpenOwner

void CAbDlg::OnDeleteSelectedItem()
{
	HRESULT			hRes = S_OK;
	CEditor Query(
		this,
		IDS_DELETEABENTRY,
		IDS_DELETEABENTRYPROMPT,
		(ULONG) 0,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
	WC_H(Query.DisplayDialog());
	if (S_OK == hRes)
	{
		DebugPrintEx(DBGGeneric,CLASS,_T("OnDeleteSelectedItem"),_T("deleting address Book entries\n"));
		CWaitCursor		Wait; // Change the mouse to an hourglass while we work.
		LPENTRYLIST lpEIDs = NULL;

		EC_H(m_lpContentsTableListCtrl->GetSelectedItemEIDs(&lpEIDs));

		EC_H(((LPABCONT)m_lpContainer)->DeleteEntries(lpEIDs,NULL));

		MAPIFreeBuffer(lpEIDs);
	}

	return;
} // CAbDlg::OnDeleteSelectedItem

BOOL CAbDlg::HandleCopy()
{
	HRESULT			hRes = S_OK;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric,CLASS,_T("HandleCopy"),_T("\n"));
	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return false;

	LPENTRYLIST lpEIDs = NULL;

	EC_H(m_lpContentsTableListCtrl->GetSelectedItemEIDs(&lpEIDs));

	// m_lpMapiObjects takes over ownership of lpEIDs - don't free now
	m_lpMapiObjects->SetABEntriesToCopy(lpEIDs);

	return true;
} // CAbDlg::HandleCopy

BOOL CAbDlg::HandlePaste()
{
	if (CBaseDialog::HandlePaste()) return true;

	HRESULT			hRes = S_OK;
	CWaitCursor		Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric,CLASS,_T("HandlePaste"),_T("pasting address Book entries\n"));
	if (!m_lpMapiObjects || !m_lpContainer) return false;

	LPENTRYLIST lpEIDs = m_lpMapiObjects->GetABEntriesToCopy();

	if (lpEIDs)
	{
		CEditor MyData(
			this,
			IDS_CALLCOPYENTRIES,
			IDS_CALLCOPYENTRIESPROMPT,
			1,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

		MyData.InitSingleLine(0,IDS_FLAGS,NULL,false);
		MyData.SetHex(0,CREATE_CHECK_DUP_STRICT);

		WC_H(MyData.DisplayDialog());
		if (S_OK == hRes)
		{
			LPMAPIPROGRESS lpProgress = GetMAPIProgress(_T("IABContainer::CopyEntries"), m_hWnd); // STRING_OK

			EC_H(((LPABCONT)m_lpContainer)->CopyEntries(
				lpEIDs,
				lpProgress ? (ULONG_PTR)m_hWnd : NULL,
				lpProgress,
				MyData.GetHex(0)));

			if(lpProgress)
				lpProgress->Release();

			lpProgress = NULL;
		}
		return true; // handled pasted
	}
	return false; // did not handle paste
} // CAbDlg::HandlePaste

void CAbDlg::OnCreatePropertyStringRestriction()
{
	HRESULT			hRes = S_OK;
	LPSRestriction	lpRes = NULL;

	if (!m_lpContentsTableListCtrl) return;

	CEditor MyData(
		this,
		IDS_SEARCHCRITERIA,
		IDS_ABSEARCHCRITERIAPROMPT,
		1,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);

	MyData.InitSingleLine(0,IDS_NAME,NULL,false);

	WC_H(MyData.DisplayDialog());
	if (S_OK != hRes) return;

	// Allocate and create our SRestriction
	EC_H(CreateANRRestriction(
		PR_ANR,
		MyData.GetString(0),
		NULL,
		&lpRes));

	m_lpContentsTableListCtrl->SetRestriction(lpRes);

	SetRestrictionType(mfcmapiNORMAL_RESTRICTION);

	if (FAILED(hRes)) MAPIFreeBuffer(lpRes);
} // CAbDlg::OnCreatePropertyStringRestriction

void CAbDlg::HandleAddInMenuSingle(
								   LPADDINMENUPARAMS lpParams,
								   LPMAPIPROP lpMAPIProp,
								   LPMAPICONTAINER /*lpContainer*/)
{
	if (lpParams)
	{
		lpParams->lpAbCont = (LPABCONT) m_lpContainer;
		lpParams->lpMailUser = (LPMAILUSER) lpMAPIProp; // OpenItemProp returns LPMAILUSER
	}

	InvokeAddInMenu(lpParams);
} // CAbDlg::HandleAddInMenuSingle
