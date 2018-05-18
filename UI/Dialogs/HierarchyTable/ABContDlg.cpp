// Displays the hierarchy tree of address books
#include "stdafx.h"
#include "AbContDlg.h"
#include <UI/Controls/HierarchyTableTreeCtrl.h>
#include <MAPI/MAPIFunctions.h>
#include <MAPI/MapiObjects.h>

static std::wstring CLASS = L"CAbContDlg";

CAbContDlg::CAbContDlg(
	_In_ CParentWnd* pParentWnd,
	_In_ CMapiObjects* lpMapiObjects
) :
	CHierarchyTableDlg(
		pParentWnd,
		lpMapiObjects,
		IDS_ABCONT,
		nullptr,
		IDR_MENU_ABCONT_POPUP,
		MENU_CONTEXT_AB_TREE)
{
	TRACE_CONSTRUCTOR(CLASS);

	auto hRes = S_OK;

	m_bIsAB = true;

	if (m_lpMapiObjects)
	{
		auto lpAddrBook = m_lpMapiObjects->GetAddrBook(false); // do not release
		if (lpAddrBook)
		{
			// Open root address book (container).
			EC_H(CallOpenEntry(
				NULL, lpAddrBook, NULL, NULL,
				nullptr,
				NULL,
				MAPI_BEST_ACCESS,
				NULL,
				reinterpret_cast<LPUNKNOWN*>(&m_lpContainer)));
		}
	}

	CreateDialogAndMenu(IDR_MENU_ABCONT);
}

CAbContDlg::~CAbContDlg()
{
	TRACE_DESTRUCTOR(CLASS);
}

BEGIN_MESSAGE_MAP(CAbContDlg, CHierarchyTableDlg)
	ON_COMMAND(ID_SETDEFAULTDIR, OnSetDefaultDir)
	ON_COMMAND(ID_SETPAB, OnSetPAB)
END_MESSAGE_MAP()

void CAbContDlg::OnSetDefaultDir()
{
	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl) return;

	auto lpItemEID = m_lpHierarchyTableTreeCtrl->GetSelectedItemEID();

	if (lpItemEID)
	{
		auto lpAddrBook = m_lpMapiObjects->GetAddrBook(false); // Do not release
		if (lpAddrBook)
		{
			EC_MAPI(lpAddrBook->SetDefaultDir(
				lpItemEID->cb,
				reinterpret_cast<LPENTRYID>(lpItemEID->lpb)));
		}
	}
}

void CAbContDlg::OnSetPAB()
{
	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl) return;

	auto lpItemEID = m_lpHierarchyTableTreeCtrl->GetSelectedItemEID();

	if (lpItemEID)
	{
		auto lpAddrBook = m_lpMapiObjects->GetAddrBook(false); // do not release
		if (lpAddrBook)
		{
			EC_MAPI(lpAddrBook->SetPAB(
				lpItemEID->cb,
				reinterpret_cast<LPENTRYID>(lpItemEID->lpb)));
		}
	}
}

void CAbContDlg::HandleAddInMenuSingle(
	_In_ LPADDINMENUPARAMS lpParams,
	_In_ LPMAPIPROP /*lpMAPIProp*/,
	_In_ LPMAPICONTAINER lpContainer)
{
	if (lpParams)
	{
		lpParams->lpAbCont = static_cast<LPABCONT>(lpContainer);
	}

	InvokeAddInMenu(lpParams);
}