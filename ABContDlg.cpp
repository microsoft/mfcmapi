// AbContDlg.cpp : implementation file
// Displays the hierarchy tree of address books

#include "stdafx.h"
#include "Error.h"

#include "AbContDlg.h"

#include "HierarchyTableTreeCtrl.h"
#include "MapiFunctions.h"
#include "MapiObjects.h"
#include "SingleMAPIPropListCtrl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CAbContDlg dialog


static TCHAR* CLASS = _T("CAbContDlg");

CAbContDlg::CAbContDlg(
					   CParentWnd* pParentWnd,
					   CMapiObjects* lpMapiObjects
					   ):
CHierarchyTableDlg(
						   pParentWnd,
						   lpMapiObjects,
						   IDS_ABCONT,
						   NULL,
						   IDR_MENU_ABCONT_POPUP,
						   MENU_CONTEXT_AB_TREE)
{
	TRACE_CONSTRUCTOR(CLASS);

	HRESULT	hRes = S_OK;

	if (m_lpMapiObjects)
	{
		LPADRBOOK lpAddrBook = m_lpMapiObjects->GetAddrBook(false);//do not release
		if (lpAddrBook)
		{
			// Open root address book (container).
			EC_H(CallOpenEntry(
				NULL,lpAddrBook,NULL,NULL,
				0,
				NULL,
				MAPI_BEST_ACCESS,
				NULL,
				(LPUNKNOWN*)&m_lpContainer));
		}
	}

	CreateDialogAndMenu(IDR_MENU_ABCONT);
}

CAbContDlg::~CAbContDlg()
{
	TRACE_DESTRUCTOR(CLASS);
}

BEGIN_MESSAGE_MAP(CAbContDlg, CHierarchyTableDlg)
//{{AFX_MSG_MAP(CAbContDlg)
	ON_COMMAND(ID_SETPAB, OnSetPAB)
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////////////
//  Menu Commands

void CAbContDlg::OnSetPAB()
{
	HRESULT			hRes = S_OK;
	CWaitCursor	Wait;//Change the mouse to an hourglass while we work.

	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl) return;

	LPSBinary	lpItemEID = NULL;
	lpItemEID = m_lpHierarchyTableTreeCtrl->GetSelectedItemEID();

	if (lpItemEID)
	{
		LPADRBOOK lpAddrBook = m_lpMapiObjects->GetAddrBook(false);//do not release
		if (lpAddrBook)
		{
			EC_H(lpAddrBook->SetPAB(
				lpItemEID->cb,
				(LPENTRYID)lpItemEID->lpb));
		}
	}
	return;
} //CAbContDlg::OnSetPAB

void CAbContDlg::HandleAddInMenuSingle(
									   LPADDINMENUPARAMS lpParams,
									   LPMAPIPROP /*lpMAPIProp*/,
									   LPMAPICONTAINER lpContainer)
{
	if (lpParams)
	{
		lpParams->lpAbCont = (LPABCONT) lpContainer;
	}

	InvokeAddInMenu(lpParams);
} // CAbContDlg::HandleAddInMenuSingle