// Displays the hierarchy tree of address books
#include <StdAfx.h>
#include <UI/Dialogs/HierarchyTable/ABContDlg.h>
#include <UI/Controls/StyleTree/HierarchyTableTreeCtrl.h>
#include <core/mapi/cache/mapiObjects.h>
#include <UI/addinui.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFunctions.h>
#include <UI/UIFunctions.h>

namespace dialog
{
	static std::wstring CLASS = L"CAbContDlg";

	CAbContDlg::CAbContDlg(_In_ ui::CParentWnd* pParentWnd, _In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects)
		: CHierarchyTableDlg(
			  pParentWnd,
			  lpMapiObjects,
			  IDS_ABCONT,
			  nullptr,
			  IDR_MENU_ABCONT_POPUP,
			  MENU_CONTEXT_AB_TREE)
	{
		TRACE_CONSTRUCTOR(CLASS);

		m_bIsAB = true;

		if (m_lpMapiObjects)
		{
			const auto lpAddrBook = m_lpMapiObjects->GetAddrBook(false); // do not release
			if (lpAddrBook)
			{
				// Open root address book (container).
				const auto container = mapi::CallOpenEntry<LPUNKNOWN>(
					nullptr, lpAddrBook, nullptr, nullptr, nullptr, nullptr, MAPI_BEST_ACCESS, nullptr);
				SetRootContainer(container);
			}
		}

		CreateDialogAndMenu(IDR_MENU_ABCONT, IDR_MENU_AB_HIERARCHY_TABLE);
		if (m_lpHierarchyTableTreeCtrl)
		{
			// Override our menu callback
			m_lpHierarchyTableTreeCtrl.HandleContextMenuCallback = [&](auto x, auto y) {
				ui::DisplayContextMenu(IDR_MENU_ABCONT_POPUP, IDR_MENU_AB_HIERARCHY_TABLE, m_hWnd, x, y);
			};
			;
		}
	}

	CAbContDlg::~CAbContDlg() { TRACE_DESTRUCTOR(CLASS); }

	BEGIN_MESSAGE_MAP(CAbContDlg, CHierarchyTableDlg)
	ON_COMMAND(ID_SETDEFAULTDIR, OnSetDefaultDir)
	ON_COMMAND(ID_SETPAB, OnSetPAB)
	END_MESSAGE_MAP()

	void CAbContDlg::OnSetDefaultDir()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl) return;

		const auto lpItemEID = m_lpHierarchyTableTreeCtrl.GetSelectedItemEID();

		if (lpItemEID)
		{
			auto lpAddrBook = m_lpMapiObjects->GetAddrBook(false); // Do not release
			if (lpAddrBook)
			{
				EC_MAPI_S(lpAddrBook->SetDefaultDir(lpItemEID->cb, reinterpret_cast<LPENTRYID>(lpItemEID->lpb)));
			}
		}
	}

	void CAbContDlg::OnSetPAB()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpMapiObjects || !m_lpHierarchyTableTreeCtrl) return;

		const auto lpItemEID = m_lpHierarchyTableTreeCtrl.GetSelectedItemEID();

		if (lpItemEID)
		{
			auto lpAddrBook = m_lpMapiObjects->GetAddrBook(false); // do not release
			if (lpAddrBook)
			{
				EC_MAPI_S(lpAddrBook->SetPAB(lpItemEID->cb, reinterpret_cast<LPENTRYID>(lpItemEID->lpb)));
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
			lpParams->lpAbCont = mapi::safe_cast<LPABCONT>(lpContainer);
		}

		ui::addinui::InvokeAddInMenu(lpParams);

		if (lpParams && lpParams->lpAbCont)
		{
			lpParams->lpAbCont->Release();
			lpParams->lpAbCont = nullptr;
		}
	}
} // namespace dialog
