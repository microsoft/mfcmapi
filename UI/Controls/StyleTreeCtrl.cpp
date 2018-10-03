#include <StdAfx.h>
#include <UI/Controls/StyleTreeCtrl.h>
#include <UI/UIFunctions.h>

namespace controls
{
	static std::wstring CLASS = L"StyleTreeCtrl";

	void StyleTreeCtrl::Create(_In_ CWnd* pCreateParent, const UINT nIDContextMenu)
	{
		m_nIDContextMenu = nIDContextMenu;

		CTreeCtrl::Create(
			TVS_HASBUTTONS | TVS_LINESATROOT | TVS_EDITLABELS | TVS_DISABLEDRAGDROP | TVS_SHOWSELALWAYS |
				TVS_FULLROWSELECT | WS_CHILD | WS_TABSTOP | WS_CLIPSIBLINGS | WS_VISIBLE,
			CRect(0, 0, 0, 0),
			pCreateParent,
			IDC_FOLDER_TREE);
		TreeView_SetBkColor(m_hWnd, ui::MyGetSysColor(ui::cBackground));
		TreeView_SetTextColor(m_hWnd, ui::MyGetSysColor(ui::cText));
		::SendMessageA(m_hWnd, WM_SETFONT, reinterpret_cast<WPARAM>(ui::GetSegoeFont()), false);
	}

	BEGIN_MESSAGE_MAP(StyleTreeCtrl, CTreeCtrl)
	ON_NOTIFY_REFLECT(NM_CUSTOMDRAW, OnCustomDraw)
	ON_WM_GETDLGCODE()
	ON_NOTIFY_REFLECT(NM_RCLICK, OnRightClick)
	ON_WM_CONTEXTMENU()
	END_MESSAGE_MAP()

	LRESULT StyleTreeCtrl::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
	{
		// Read the current hover local, since we need to clear it before we do any drawing
		const auto hItemCurHover = m_hItemCurHover;
		switch (message)
		{
		case WM_MOUSEMOVE:
		{
			TVHITTESTINFO tvHitTestInfo = {0};
			tvHitTestInfo.pt.x = GET_X_LPARAM(lParam);
			tvHitTestInfo.pt.y = GET_Y_LPARAM(lParam);

			WC_B_S(::SendMessage(m_hWnd, TVM_HITTEST, 0, reinterpret_cast<LPARAM>(&tvHitTestInfo)));
			if (tvHitTestInfo.hItem)
			{
				if (tvHitTestInfo.flags & TVHT_ONITEMBUTTON)
				{
					m_HoverButton = true;
				}
				else
				{
					m_HoverButton = false;
				}

				// If this is a new glow, clean up the old glow and track for leaving the control
				if (hItemCurHover != tvHitTestInfo.hItem)
				{
					ui::DrawTreeItemGlow(m_hWnd, tvHitTestInfo.hItem);

					if (hItemCurHover)
					{
						m_hItemCurHover = nullptr;
						ui::DrawTreeItemGlow(m_hWnd, hItemCurHover);
					}

					m_hItemCurHover = tvHitTestInfo.hItem;

					TRACKMOUSEEVENT tmEvent = {0};
					tmEvent.cbSize = sizeof(TRACKMOUSEEVENT);
					tmEvent.dwFlags = TME_LEAVE;
					tmEvent.hwndTrack = m_hWnd;

					WC_B_S(TrackMouseEvent(&tmEvent));
				}
			}
			else
			{
				if (hItemCurHover)
				{
					m_hItemCurHover = nullptr;
					ui::DrawTreeItemGlow(m_hWnd, hItemCurHover);
				}
			}
			break;
		}
		case WM_MOUSELEAVE:
			if (hItemCurHover)
			{
				m_hItemCurHover = nullptr;
				ui::DrawTreeItemGlow(m_hWnd, hItemCurHover);
			}

			return NULL;
		}

		return CTreeCtrl::WindowProc(message, wParam, lParam);
	}

	void StyleTreeCtrl::OnCustomDraw(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
	{
		ui::CustomDrawTree(pNMHDR, pResult, m_HoverButton, m_hItemCurHover);
	}

	// Removes any existing node data and replaces it with lpData
	void StyleTreeCtrl::SetNodeData(HWND hWnd, HTREEITEM hItem, const LPARAM lpData) const
	{
		if (lpData)
		{
			TVITEM tvItem = {0};
			tvItem.hItem = hItem;
			tvItem.mask = TVIF_PARAM;
			if (TreeView_GetItem(hWnd, &tvItem) && tvItem.lParam)
			{
				output::DebugPrintEx(DBGHierarchy, CLASS, L"SetNodeData", L"Node %p, replacing data\n", hItem);
				FreeNodeData(tvItem.lParam);
			}
			else
			{
				output::DebugPrintEx(DBGHierarchy, CLASS, L"SetNodeData", L"Node %p, first data\n", hItem);
			}

			tvItem.lParam = lpData;
			TreeView_SetItem(hWnd, &tvItem);
			// The tree now owns our lpData
		}
	}

	void StyleTreeCtrl::OnSelChanged(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
	{
		const auto pNMTV = reinterpret_cast<LPNMTREEVIEW>(pNMHDR);

		if (pNMTV && pNMTV->itemNew.hItem)
		{
			m_bItemSelected = true;

			OnItemSelected(pNMTV->itemNew.hItem);
		}
		else
		{
			m_bItemSelected = false;
		}

		if (pResult) *pResult = 0;
	}

	// Assert that we want all keyboard input (including ENTER!)
	_Check_return_ UINT StyleTreeCtrl::OnGetDlgCode()
	{
		auto iDlgCode = CTreeCtrl::OnGetDlgCode() | DLGC_WANTMESSAGE;

		// to make sure that the control key is not pressed
		if (GetKeyState(VK_CONTROL) >= 0 && m_hWnd == ::GetFocus())
		{
			// to make sure that the Tab key is pressed
			if (GetKeyState(VK_TAB) < 0) iDlgCode &= ~(DLGC_WANTALLKEYS | DLGC_WANTMESSAGE | DLGC_WANTTAB);
		}

		return iDlgCode;
	}

	void StyleTreeCtrl::OnRightClick(_In_ NMHDR* /*pNMHDR*/, _In_ LRESULT* pResult)
	{
		// Send WM_CONTEXTMENU to self
		(void) SendMessage(WM_CONTEXTMENU, reinterpret_cast<WPARAM>(m_hWnd), GetMessagePos());

		// Mark message as handled and suppress default handling
		*pResult = 1;
	}

	void StyleTreeCtrl::OnContextMenu(_In_ CWnd* pWnd, CPoint pos)
	{
		// If we don't have a position, this may be keyboard initiated context menu. Use the highlighted/selected item to find a position.
		if (pWnd && -1 == pos.x && -1 == pos.y)
		{
			// Find the highlighted item
			const auto item = GetSelectedItem();

			if (item)
			{
				RECT rc = {0};
				GetItemRect(item, &rc, true);
				pos.x = rc.left;
				pos.y = rc.top;
				::ClientToScreen(pWnd->m_hWnd, &pos);
			}
		}
		// If we have a position, make sure the item at that position is selected.
		else
		{
			// Select the item that is at the point pos.
			UINT uFlags = NULL;
			auto ptTree = pos;
			ScreenToClient(&ptTree);
			const auto hClickedItem = HitTest(ptTree, &uFlags);

			if (hClickedItem != nullptr && TVHT_ONITEM & uFlags)
			{
				Select(hClickedItem, TVGN_CARET);
			}
		}

		HandleContextMenu(pos.x, pos.y);
	}
} // namespace controls