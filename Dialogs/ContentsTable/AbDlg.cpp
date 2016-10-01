// AbDlg.cpp : Displays the contents of a single address book

#include "stdafx.h"
#include "ABDlg.h"
#include "ContentsTableListCtrl.h"
#include "MapiObjects.h"
#include "SingleMAPIPropListCtrl.h"
#include "ColumnTags.h"
#include "MFCUtilityFunctions.h"
#include "UIFunctions.h"
#include <Dialogs/Editors/Editor.h>
#include "MAPIABFunctions.h"
#include "MAPIProgress.h"
#include "MAPIFunctions.h"
#include "GlobalCache.h"

static wstring CLASS = L"CAbDlg";

CAbDlg::CAbDlg(
	_In_ CParentWnd* pParentWnd,
	_In_ CMapiObjects* lpMapiObjects,
	_In_ LPABCONT lpAdrBook
) :
	CContentsTableDlg(
		pParentWnd,
		lpMapiObjects,
		IDS_AB,
		mfcmapiDO_NOT_CALL_CREATE_DIALOG,
		nullptr,
		LPSPropTagArray(&sptABCols),
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
	DebugPrintEx(DBGCreateDialog, CLASS, L"CreateDialogAndMenu", L"id = 0x%X\n", nIDMenuResource);
	CContentsTableDlg::CreateDialogAndMenu(nIDMenuResource);

	UpdateMenuString(
		m_hWnd,
		ID_CREATEPROPERTYSTRINGRESTRICTION,
		IDS_ABRESMENU);
}

void CAbDlg::OnInitMenu(_In_ CMenu* pMenu)
{
	if (pMenu && m_lpContentsTableListCtrl)
	{
		auto ulStatus = CGlobalCache::getInstance().GetBufferStatus();
		pMenu->EnableMenuItem(ID_PASTE, DIM(ulStatus & BUFFER_ABENTRIES));

		int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
		pMenu->EnableMenuItem(ID_COPY, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_DELETESELECTEDITEM, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_DISPLAYDETAILS, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_OPENCONTACT, DIMMSNOK(iNumSel));
		pMenu->EnableMenuItem(ID_OPENMANAGER, DIMMSOK(iNumSel));
		pMenu->EnableMenuItem(ID_OPENOWNER, DIMMSOK(iNumSel));
	}

	CContentsTableDlg::OnInitMenu(pMenu);
}

void CAbDlg::OnDisplayDetails()
{
	DebugPrintEx(DBGGeneric, CLASS, L"OnDisplayDetails", L"displaying Address Book entry details\n");

	auto hRes = S_OK;
	if (!m_lpMapiObjects) return;
	auto lpAddrBook = m_lpMapiObjects->GetAddrBook(false); // do not release
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (lpAddrBook)
	{
		LPENTRYLIST lpEIDs = nullptr;
		EC_H(m_lpContentsTableListCtrl->GetSelectedItemEIDs(&lpEIDs));

		if (lpEIDs && lpEIDs->cValues && lpEIDs->lpbin)
		{

			for (ULONG i = 0; i < lpEIDs->cValues; i++)
			{
				auto ulUIParam = reinterpret_cast<ULONG_PTR>(static_cast<void*>(m_hWnd));

				// Have to pass DIALOG_MODAL according to
				// http://support.microsoft.com/kb/171637
				EC_H_CANCEL(lpAddrBook->Details(
					&ulUIParam,
					NULL,
					NULL,
					lpEIDs->lpbin[i].cb,
					reinterpret_cast<LPENTRYID>(lpEIDs->lpbin[i].lpb),
					NULL,
					NULL,
					NULL,
					DIALOG_MODAL)); // API doesn't like Unicode
				if (lpEIDs->cValues > i + 1 && bShouldCancel(this, hRes)) break;
				hRes = S_OK;
			}
		}

		MAPIFreeBuffer(lpEIDs);
	}
}

void CAbDlg::OnOpenContact()
{
	auto hRes = S_OK;
	LPENTRYLIST lpEntryList = nullptr;
	LPMAPIPROP lpProp = nullptr;

	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl || !m_lpPropDisplay) return;
	auto lpSession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpSession) return;

	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	EC_H(m_lpContentsTableListCtrl->GetSelectedItemEIDs(&lpEntryList));

	if (SUCCEEDED(hRes) && lpEntryList && 1 == lpEntryList->cValues)
	{
		ULONG cb = 0;
		LPBYTE lpb = nullptr;
		if (UnwrapContactEntryID(lpEntryList->lpbin[0].cb, lpEntryList->lpbin[0].lpb, &cb, &lpb))
		{
			EC_H(CallOpenEntry(
				NULL,
				NULL,
				NULL,
				lpSession,
				cb,
				reinterpret_cast<LPENTRYID>(lpb),
				NULL,
				NULL,
				NULL,
				reinterpret_cast<LPUNKNOWN*>(&lpProp)));
		}
	}

	if (lpProp)
	{
		EC_H(DisplayObject(lpProp, NULL, otDefault, this));
		if (lpProp) lpProp->Release();
	}

	MAPIFreeBuffer(lpEntryList);
}

void CAbDlg::OnOpenManager()
{
	LPMAILUSER lpMailUser = nullptr;
	auto iItem = -1;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl) return;

	do
	{
		auto hRes = S_OK;
		if (lpMailUser) lpMailUser->Release();
		lpMailUser = nullptr;
		EC_H(m_lpContentsTableListCtrl->OpenNextSelectedItemProp(
			&iItem,
			mfcmapiREQUEST_MODIFY,
			reinterpret_cast<LPMAPIPROP*>(&lpMailUser)));

		if (lpMailUser)
		{
			EC_H(DisplayTable(
				lpMailUser,
				PR_EMS_AB_MANAGER_O,
				otDefault, // oType,
				this));
		}
	} while (iItem != -1);

	if (lpMailUser) lpMailUser->Release();
}

void CAbDlg::OnOpenOwner()
{
	LPMAILUSER lpMailUser = nullptr;
	auto iItem = -1;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpContentsTableListCtrl) return;

	do
	{
		auto hRes = S_OK;
		if (lpMailUser) lpMailUser->Release();
		lpMailUser = nullptr;
		EC_H(m_lpContentsTableListCtrl->OpenNextSelectedItemProp(
			&iItem,
			mfcmapiREQUEST_MODIFY,
			reinterpret_cast<LPMAPIPROP*>(&lpMailUser)));

		if (lpMailUser)
		{
			EC_H(DisplayTable(
				lpMailUser,
				PR_EMS_AB_OWNER_O,
				otDefault, // oType,
				this));
		}
	} while (iItem != -1);

	if (lpMailUser) lpMailUser->Release();
}

void CAbDlg::OnDeleteSelectedItem()
{
	auto hRes = S_OK;
	CEditor Query(
		this,
		IDS_DELETEABENTRY,
		IDS_DELETEABENTRYPROMPT,
		static_cast<ULONG>(0),
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	WC_H(Query.DisplayDialog());
	if (S_OK == hRes)
	{
		DebugPrintEx(DBGGeneric, CLASS, L"OnDeleteSelectedItem", L"deleting address Book entries\n");
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.
		LPENTRYLIST lpEIDs = nullptr;

		EC_H(m_lpContentsTableListCtrl->GetSelectedItemEIDs(&lpEIDs));

		EC_MAPI((static_cast<LPABCONT>(m_lpContainer))->DeleteEntries(lpEIDs, NULL));

		MAPIFreeBuffer(lpEIDs);
	}
}

void CAbDlg::HandleCopy()
{
	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric, CLASS, L"HandleCopy", L"\n");
	if (!m_lpContentsTableListCtrl) return;

	LPENTRYLIST lpEIDs = nullptr;

	EC_H(m_lpContentsTableListCtrl->GetSelectedItemEIDs(&lpEIDs));

	// CGlobalCache takes over ownership of lpEIDs - don't free now
	CGlobalCache::getInstance().SetABEntriesToCopy(lpEIDs);
}

_Check_return_ bool CAbDlg::HandlePaste()
{
	if (CBaseDialog::HandlePaste()) return true;

	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric, CLASS, L"HandlePaste", L"pasting address Book entries\n");
	if (!m_lpContainer) return false;

	auto lpEIDs = CGlobalCache::getInstance().GetABEntriesToCopy();

	if (lpEIDs)
	{
		CEditor MyData(
			this,
			IDS_CALLCOPYENTRIES,
			IDS_CALLCOPYENTRIESPROMPT,
			1,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyData.InitPane(0, CreateSingleLinePane(IDS_FLAGS, false));
		MyData.SetHex(0, CREATE_CHECK_DUP_STRICT);

		WC_H(MyData.DisplayDialog());
		if (S_OK == hRes)
		{
			LPMAPIPROGRESS lpProgress = GetMAPIProgress(L"IABContainer::CopyEntries", m_hWnd); // STRING_OK

			EC_MAPI((static_cast<LPABCONT>(m_lpContainer))->CopyEntries(
				lpEIDs,
				lpProgress ? reinterpret_cast<ULONG_PTR>(m_hWnd) : NULL,
				lpProgress,
				MyData.GetHex(0)));

			if (lpProgress)
				lpProgress->Release();
		}

		return true; // handled pasted
	}

	return false; // did not handle paste
}

void CAbDlg::OnCreatePropertyStringRestriction()
{
	auto hRes = S_OK;
	LPSRestriction lpRes = nullptr;

	if (!m_lpContentsTableListCtrl) return;

	CEditor MyData(
		this,
		IDS_SEARCHCRITERIA,
		IDS_ABSEARCHCRITERIAPROMPT,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	MyData.InitPane(0, CreateSingleLinePane(IDS_NAME, false));

	WC_H(MyData.DisplayDialog());
	if (S_OK != hRes) return;

	// Allocate and create our SRestriction
	EC_H(CreateANRRestriction(
		PR_ANR_W,
		MyData.GetStringW(0),
		NULL,
		&lpRes));

	m_lpContentsTableListCtrl->SetRestriction(lpRes);

	SetRestrictionType(mfcmapiNORMAL_RESTRICTION);

	if (FAILED(hRes)) MAPIFreeBuffer(lpRes);
}

void CAbDlg::HandleAddInMenuSingle(
	_In_ LPADDINMENUPARAMS lpParams,
	_In_ LPMAPIPROP lpMAPIProp,
	_In_ LPMAPICONTAINER /*lpContainer*/)
{
	if (lpParams)
	{
		lpParams->lpAbCont = static_cast<LPABCONT>(m_lpContainer);
		lpParams->lpMailUser = static_cast<LPMAILUSER>(lpMAPIProp); // OpenItemProp returns LPMAILUSER
	}

	InvokeAddInMenu(lpParams);
}