#include <StdAfx.h>
#include <UI/Controls/StyleTree/StyleTreeCtrl.h>
#include <UI/UIFunctions.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>

namespace controls
{
	static std::wstring CLASS = L"StyleTreeCtrl";

	void StyleTreeCtrl::Create(_In_ CWnd* pCreateParent, const bool bReadOnly)
	{
		m_bReadOnly = bReadOnly;

		auto style = TVS_HASBUTTONS | TVS_LINESATROOT | TVS_DISABLEDRAGDROP | TVS_SHOWSELALWAYS | TVS_FULLROWSELECT |
					 WS_CHILD | WS_TABSTOP | WS_CLIPSIBLINGS | WS_VISIBLE;
		if (!bReadOnly) style |= TVS_EDITLABELS;
		CTreeCtrl::Create(style, CRect(0, 0, 0, 0), pCreateParent, IDC_FOLDER_TREE);
		TreeView_SetBkColor(m_hWnd, ui::MyGetSysColor(ui::uiColor::Background));
		TreeView_SetTextColor(m_hWnd, ui::MyGetSysColor(ui::uiColor::Text));
		::SendMessageA(m_hWnd, WM_SETFONT, reinterpret_cast<WPARAM>(ui::GetSegoeFont()), false);
	}

	BEGIN_MESSAGE_MAP(StyleTreeCtrl, CTreeCtrl)
#pragma warning(push)
#pragma warning( \
	disable : 26454) // Warning C26454 Arithmetic overflow: 'operator' operation produces a negative unsigned result at compile time
	ON_NOTIFY_REFLECT(NM_CUSTOMDRAW, OnCustomDraw)
	ON_NOTIFY_REFLECT(NM_RCLICK, OnRightClick)
	ON_NOTIFY_REFLECT(TVN_GETDISPINFO, OnGetDispInfo)
	ON_NOTIFY_REFLECT(TVN_ITEMEXPANDING, OnItemExpanding)
	ON_NOTIFY_REFLECT(TVN_SELCHANGED, OnSelChanged)
	ON_NOTIFY_REFLECT(TVN_ENDLABELEDIT, OnEndLabelEdit)
	ON_NOTIFY_REFLECT(TVN_DELETEITEM, OnDeleteItem)
	ON_NOTIFY_REFLECT(NM_DBLCLK, OnDblclk)
#pragma warning(pop)
	END_MESSAGE_MAP()

	LRESULT StyleTreeCtrl::WindowProc(const UINT message, const WPARAM wParam, const LPARAM lParam)
	{
		// Read the current hover local, since we need to clear it before we do any drawing
		switch (message)
		{
		case WM_GETDLGCODE:
			return OnGetDlgCode();
		case WM_CONTEXTMENU:
			OnContextMenu(reinterpret_cast<HWND>(wParam), GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam));
			return 0;
		case WM_KEYDOWN:
			OnKeyDown(static_cast<UINT>(wParam), LOWORD(lParam), HIWORD(lParam));
			return 0;
		case WM_MOUSEMOVE:
		{
			const auto hItemCurHover = m_hItemCurHover;
			auto tvHitTestInfo = TVHITTESTINFO{};
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

					auto tmEvent = TRACKMOUSEEVENT{};
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
			if (m_hItemCurHover)
			{
				const auto hItemCurHover = m_hItemCurHover;
				m_hItemCurHover = nullptr;
				ui::DrawTreeItemGlow(m_hWnd, hItemCurHover);
			}

			return NULL;
		}

		return CTreeCtrl::WindowProc(message, wParam, lParam);
	}

	BOOL StyleTreeCtrl::PreTranslateMessage(MSG* pMsg)
	{
		// If edit control is visible in tree view control, when you send a
		// WM_KEYDOWN message to the edit control it will dismiss the edit
		// control. When the ENTER key was sent to the edit control, the
		// parent window of the tree view control is responsible for updating
		// the item's label in TVN_ENDLABELEDIT notification code.
		if (pMsg && pMsg->message == WM_KEYDOWN && (pMsg->wParam == VK_RETURN || pMsg->wParam == VK_ESCAPE))
		{
			const auto edit = GetEditControl();
			if (edit)
			{
				edit->SendMessage(WM_KEYDOWN, pMsg->wParam, pMsg->lParam);
				return true;
			}
		}

		return CTreeCtrl::PreTranslateMessage(pMsg);
	}

	void StyleTreeCtrl::OnCustomDraw(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
	{
		ui::CustomDrawTree(pNMHDR, pResult, m_HoverButton, m_hItemCurHover);
		if (OnCustomDrawCallback) OnCustomDrawCallback(pNMHDR, pResult, m_hItemCurHover);
	}

	// Removes any existing node data and replaces it with lpData
	void StyleTreeCtrl::SetNodeData(HWND hWnd, HTREEITEM hItem, const LPVOID lpData) const
	{
		if (lpData)
		{
			auto tvItem = TVITEM{};
			tvItem.hItem = hItem;
			tvItem.mask = TVIF_PARAM;
			if (TreeView_GetItem(hWnd, &tvItem) && tvItem.lParam)
			{
				output::DebugPrintEx(
					output::dbgLevel::Hierarchy, CLASS, L"SetNodeData", L"Node %p, replacing data\n", hItem);
				if (FreeNodeDataCallback) FreeNodeDataCallback(tvItem.lParam);
			}
			else
			{
				output::DebugPrintEx(
					output::dbgLevel::Hierarchy, CLASS, L"SetNodeData", L"Node %p, first data\n", hItem);
			}

			tvItem.lParam = reinterpret_cast<LPARAM>(lpData);
			TreeView_SetItem(hWnd, &tvItem);
			// The tree now owns our lpData
		}
	}

	void StyleTreeCtrl::OnSelChanged(_In_opt_ NMHDR* pNMHDR, _In_opt_ LRESULT* pResult)
	{
		const auto pNMTV = reinterpret_cast<LPNMTREEVIEW>(pNMHDR);

		if (pNMTV && pNMTV->itemNew.hItem)
		{
			m_bItemSelected = true;

			if (ItemSelectedCallback) ItemSelectedCallback(pNMTV->itemNew.hItem);
		}
		else
		{
			m_bItemSelected = false;
		}

		if (pResult) *pResult = 0;
	}

	std::wstring StyleTreeCtrl::GetItemTextW(HTREEITEM hItem) const
	{
		return strings::LPCTSTRToWstring(GetItemText(hItem));
	}

	HTREEITEM
	StyleTreeCtrl::AddChildNode(
		_In_ const std::wstring& szName,
		HTREEITEM hParent,
		const LPVOID lpData,
		const std::function<void(HTREEITEM hItem)>& itemAddedCallback) const
	{
		output::DebugPrintEx(
			output::dbgLevel::Hierarchy,
			CLASS,
			L"AddNode",
			L"Adding Node \"%ws\" under node %p, callback = %ws\n",
			szName.c_str(),
			hParent,
			itemAddedCallback != nullptr ? L"true" : L"false");
		auto tvInsert = TVINSERTSTRUCTW{};

		tvInsert.hParent = hParent;
		tvInsert.hInsertAfter = m_bSortNodes ? TVI_SORT : hParent;
		tvInsert.item.mask = TVIF_CHILDREN | TVIF_TEXT;
		tvInsert.item.cChildren = HasChildrenCallback ? I_CHILDRENCALLBACK : I_CHILDRENAUTO;
		tvInsert.item.pszText = const_cast<LPWSTR>(szName.c_str());
		const auto hItem =
			reinterpret_cast<HTREEITEM>(::SendMessage(m_hWnd, TVM_INSERTITEMW, 0, reinterpret_cast<LPARAM>(&tvInsert)));

		SetNodeData(m_hWnd, hItem, lpData);

		if (itemAddedCallback)
		{
			itemAddedCallback(hItem);
		}

		return hItem;
	}

	void StyleTreeCtrl::OnGetDispInfo(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
	{
		const auto lpDispInfo = reinterpret_cast<LPNMTVDISPINFO>(pNMHDR);

		if (lpDispInfo && lpDispInfo->item.mask & TVIF_CHILDREN && HasChildrenCallback)
		{
			lpDispInfo->item.cChildren = HasChildrenCallback(reinterpret_cast<HTREEITEM>(lpDispInfo->item.lParam));
		}

		*pResult = 0;
	}

	// Assert that we want all keyboard input (including ENTER!)
	_Check_return_ UINT StyleTreeCtrl::OnGetDlgCode()
	{
		auto iDlgCode = CTreeCtrl::OnGetDlgCode() | DLGC_WANTMESSAGE;

		// Make sure that the control key is not pressed
		if (GetKeyState(VK_CONTROL) >= 0 && m_hWnd == ::GetFocus())
		{
			// Make sure that the Tab key is pressed
			if (GetKeyState(VK_TAB) < 0) iDlgCode &= ~(DLGC_WANTALLKEYS | DLGC_WANTMESSAGE | DLGC_WANTTAB);
		}

		return iDlgCode;
	}

	void StyleTreeCtrl::OnRightClick(_In_ NMHDR* /*pNMHDR*/, _In_ LRESULT* pResult)
	{
		// Send WM_CONTEXTMENU to self
		static_cast<void>(SendMessage(WM_CONTEXTMENU, reinterpret_cast<WPARAM>(m_hWnd), GetMessagePos()));

		// Mark message as handled and suppress default handling
		*pResult = 1;
	}

	void StyleTreeCtrl::OnContextMenu(_In_ HWND hwnd, const int x, const int y)
	{
		auto pos = POINT{x, y};
		// If we don't have a position, this may be keyboard initiated context menu. Use the highlighted/selected item to find a position.
		if (hwnd && -1 == pos.x && -1 == pos.y)
		{
			// Find the highlighted item
			const auto item = GetSelectedItem();

			if (item)
			{
				auto rc = RECT{};
				static_cast<void>(GetItemRect(item, &rc, true));
				pos.x = rc.left;
				pos.y = rc.top;
				::ClientToScreen(hwnd, &pos);
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

		if (HandleContextMenuCallback) HandleContextMenuCallback(pos.x, pos.y);
	}

	void StyleTreeCtrl::OnItemExpanding(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
	{
		*pResult = 0;

		const auto pNMTreeView = reinterpret_cast<NM_TREEVIEW*>(pNMHDR);
		if (pNMTreeView)
		{
			output::DebugPrintEx(
				output::dbgLevel::Hierarchy,
				CLASS,
				L"OnItemExpanding",
				L"Expanding item %p \"%ws\" action = 0x%08X state = 0x%08X\n",
				pNMTreeView->itemNew.hItem,
				GetItemTextW(pNMTreeView->itemOld.hItem).c_str(),
				pNMTreeView->action,
				pNMTreeView->itemNew.state);
			if (pNMTreeView->action & TVE_EXPAND)
			{
				if (!(pNMTreeView->itemNew.state & TVIS_EXPANDEDONCE))
				{
					if (ExpandNodeCallback) ExpandNodeCallback(pNMTreeView->itemNew.hItem);
				}
			}
		}
	}

	void StyleTreeCtrl::Refresh()
	{
		// Turn off redraw while we work on the window
		SetRedraw(false);

		OnSelChanged(nullptr, nullptr);

		EC_B_S(DeleteAllItems());
		if (OnRefreshCallback) OnRefreshCallback();

		// Turn redraw back on to update our view
		SetRedraw(true);
	}

	// This function will be called when we edit a node so we can attempt to commit the changes
	// TODO: In non-unicode builds, this gives us ANSI strings - need to figure out how to change that
	void StyleTreeCtrl::OnEndLabelEdit(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
	{
		const auto pTVDispInfo = reinterpret_cast<TV_DISPINFO*>(pNMHDR);
		*pResult = 0;

		if (!pTVDispInfo || !pTVDispInfo->item.pszText) return;

		if (OnLabelEditCallback) OnLabelEditCallback(pTVDispInfo->item.hItem, pTVDispInfo->item.pszText);
	}

	void StyleTreeCtrl::OnDblclk(_In_ NMHDR* /*pNMHDR*/, _In_ LRESULT* pResult)
	{
		if (OnDisplaySelectedItemCallback) OnDisplaySelectedItemCallback();

		// Don't do default behavior for double-click (We only want '+' sign expansion.
		// Double click should display the item, not expand the tree.)
		*pResult = 1;
	}

	void StyleTreeCtrl::OnKeyDown(const UINT nChar, const UINT nRepCnt, const UINT nFlags)
	{
		output::DebugPrintEx(output::dbgLevel::Menu, CLASS, L"OnKeyDown", L"0x%X\n", nChar);

		const auto bCtrlPressed = GetKeyState(VK_CONTROL) < 0;
		const auto bShiftPressed = GetKeyState(VK_SHIFT) < 0;
		const auto bMenuPressed = GetKeyState(VK_MENU) < 0;

		if (!bMenuPressed)
		{
			if (!KeyDownCallback || !KeyDownCallback(nChar, bShiftPressed, bCtrlPressed, bMenuPressed))
			{
				CTreeCtrl::OnKeyDown(nChar, nRepCnt, nFlags);
			}
		}
	}

	// Tree control will call this for every node it deletes.
	void StyleTreeCtrl::OnDeleteItem(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
	{
		const auto pNMTreeView = reinterpret_cast<LPNMTREEVIEW>(pNMHDR);
		if (pNMTreeView)
		{
			output::DebugPrintEx(
				output::dbgLevel::Hierarchy,
				CLASS,
				L"OnDeleteItem",
				L"Deleting item %p \"%ws\"\n",
				pNMTreeView->itemOld.hItem,
				GetItemTextW(pNMTreeView->itemOld.hItem).c_str());

			if (FreeNodeDataCallback) FreeNodeDataCallback(pNMTreeView->itemOld.lParam);

			if (!m_bShuttingDown)
			{
				// Collapse the parent if this is the last child
				const auto hPrev = TreeView_GetPrevSibling(m_hWnd, pNMTreeView->itemOld.hItem);
				const auto hNext = TreeView_GetNextSibling(m_hWnd, pNMTreeView->itemOld.hItem);

				if (!(hPrev || hNext))
				{
					output::DebugPrintEx(
						output::dbgLevel::Hierarchy,
						CLASS,
						L"OnDeleteItem",
						L"%p has no siblings\n",
						pNMTreeView->itemOld.hItem);
					const auto hParent = TreeView_GetParent(m_hWnd, pNMTreeView->itemOld.hItem);
					TreeView_SetItemState(m_hWnd, hParent, 0, TVIS_EXPANDED | TVIS_EXPANDEDONCE);
					auto tvItem = TVITEM{};
					tvItem.hItem = hParent;
					tvItem.mask = TVIF_PARAM;
					if (TreeView_GetItem(m_hWnd, &tvItem) && tvItem.lParam)
					{
						if (OnLastChildDeletedCallback) OnLastChildDeletedCallback(tvItem.lParam);
					}
				}
			}
		}

		*pResult = 0;
	}
} // namespace controls