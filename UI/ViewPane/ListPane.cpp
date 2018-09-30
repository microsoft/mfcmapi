#include <StdAfx.h>
#include <UI/ViewPane/ListPane.h>
#include <UI/UIFunctions.h>
#include <utility>

namespace viewpane
{
	static std::wstring CLASS = L"ListPane";

	ListPane* ListPane::Create(
		const int paneID,
		const UINT uidLabel,
		const bool bAllowSort,
		const bool bReadOnly,
		DoListEditCallback callback)
	{
		auto pane = new (std::nothrow) ListPane();
		if (pane)
		{
			pane->Setup(bAllowSort, std::move(callback));
			pane->SetLabel(uidLabel);
			pane->SetReadOnly(bReadOnly);
			pane->m_paneID = paneID;
		}

		return pane;
	}

	ListPane* ListPane::CreateCollapsibleListPane(
		const int paneID,
		const UINT uidLabel,
		const bool bAllowSort,
		const bool bReadOnly,
		DoListEditCallback callback)
	{
		auto pane = new (std::nothrow) ListPane();
		if (pane)
		{
			pane->Setup(bAllowSort, std::move(callback));
			pane->SetLabel(uidLabel);
			pane->SetReadOnly(bReadOnly);
			pane->m_bCollapsible = true;
			pane->m_paneID = paneID;
		}

		return pane;
	}

	void ListPane::Setup(bool bAllowSort, DoListEditCallback callback)
	{
		m_List.AllowEscapeClose();
		m_bAllowSort = bAllowSort;
		m_callback = std::move(callback);
	}

	bool ListPane::IsDirty() { return m_bDirty; }

	void ListPane::Initialize(_In_ CWnd* pParent, _In_ HDC hdc)
	{
		ViewPane::Initialize(pParent, hdc);

		DWORD dwListStyle = LVS_SINGLESEL | WS_BORDER;
		if (!m_bAllowSort) dwListStyle |= LVS_NOSORTHEADER;
		m_List.Create(pParent, dwListStyle, m_nID, false);

		// read only lists don't need buttons
		if (!m_bReadOnly)
		{
			for (auto iButton = 0; iButton < NUMLISTBUTTONS; iButton++)
			{
				const auto szButtonText = strings::loadstring(ListButtons[iButton]);

				EC_B_S(m_ButtonArray[iButton].Create(
					strings::wstringTotstring(szButtonText).c_str(),
					WS_TABSTOP | WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE,
					CRect(0, 0, 0, 0),
					pParent,
					ListButtons[iButton]));

				const auto sizeText = ui::GetTextExtentPoint32(hdc, szButtonText);
				m_iButtonWidth = max(m_iButtonWidth, sizeText.cx);
			}
		}

		m_iButtonWidth += m_iMargin;
		m_bInitialized = true;
	}

	int ListPane::GetMinWidth(_In_ HDC hdc)
	{
		return max(
			ViewPane::GetMinWidth(hdc), (int) (NUMLISTBUTTONS * m_iButtonWidth + m_iMargin * (NUMLISTBUTTONS - 1)));
	}

	int ListPane::GetFixedHeight()
	{
		auto iHeight = 0;
		if (0 != m_paneID) iHeight += m_iSmallHeightMargin; // Top margin

		iHeight += GetLabelHeight();

		if (!m_bCollapsed)
		{
			iHeight += m_iSmallHeightMargin;

			if (!m_bReadOnly)
			{
				iHeight += m_iLargeHeightMargin + m_iButtonHeight;
			}
		}

		return iHeight;
	}

	int ListPane::GetLines()
	{
		if (m_bCollapsed)
		{
			return 0;
		}

		return LINES_LIST;
	}

	ULONG ListPane::HandleChange(UINT nID)
	{
		switch (nID)
		{
		case IDD_LISTMOVEDOWN:
			OnMoveListEntryDown();
			break;
		case IDD_LISTADD:
			OnAddListEntry();
			break;
		case IDD_LISTEDIT:
			(void) OnEditListEntry();
			break;
		case IDD_LISTDELETE:
			OnDeleteListEntry(true);
			break;
		case IDD_LISTMOVEUP:
			OnMoveListEntryUp();
			break;
		case IDD_LISTMOVETOBOTTOM:
			OnMoveListEntryToBottom();
			break;
		case IDD_LISTMOVETOTOP:
			OnMoveListEntryToTop();
			break;
		}

		return ViewPane::HandleChange(nID);
	}

	void ListPane::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		_In_ const int x,
		_In_ const int y,
		_In_ const int width,
		_In_ const int height)
	{
		auto curY = y;
		const auto labelHeight = GetLabelHeight();
		if (0 != m_paneID)
		{
			curY += m_iSmallHeightMargin;
		}

		// Layout our label
		ViewPane::DeferWindowPos(hWinPosInfo, x, curY, width, height - (curY - y));
		curY += labelHeight + m_iSmallHeightMargin;

		const auto cmdShow = m_bCollapsed ? SW_HIDE : SW_SHOW;
		EC_B_S(m_List.ShowWindow(cmdShow));
		auto listHeight = height - (curY - y);
		if (!m_bReadOnly) listHeight -= (m_iLargeHeightMargin + m_iButtonHeight);
		EC_B_S(::DeferWindowPos(hWinPosInfo, m_List.GetSafeHwnd(), nullptr, x, curY, width, listHeight, SWP_NOZORDER));

		if (!m_bReadOnly)
		{
			// buttons go below the list:
			curY += listHeight + m_iLargeHeightMargin;

			const auto iSlotWidth = m_iButtonWidth + m_iMargin;
			const auto iOffset = width + m_iSideMargin + m_iMargin;

			for (auto iButton = 0; iButton < NUMLISTBUTTONS; iButton++)
			{
				EC_B_S(m_ButtonArray[iButton].ShowWindow(cmdShow));
				EC_B_S(::DeferWindowPos(
					hWinPosInfo,
					m_ButtonArray[iButton].GetSafeHwnd(),
					nullptr,
					iOffset - iSlotWidth * (NUMLISTBUTTONS - iButton),
					curY,
					m_iButtonWidth,
					m_iButtonHeight,
					SWP_NOZORDER));
			}
		}
	}

	void ListPane::CommitUIValues() {}

	void ListPane::SetListString(ULONG iListRow, ULONG iListCol, const std::wstring& szListString)
	{
		m_List.SetItemText(iListRow, iListCol, szListString);
	}

	_Check_return_ controls::sortlistdata::SortListData* ListPane::InsertRow(int iRow, const std::wstring& szText) const
	{
		return m_List.InsertRow(iRow, szText);
	}

	void ListPane::ClearList()
	{
		m_List.DeleteAllColumns();
		EC_B_S(m_List.DeleteAllItems());
	}

	void ListPane::ResizeList(bool bSort)
	{
		if (bSort)
		{
			m_List.SortClickedColumn();
		}

		m_List.AutoSizeColumns(false);
	}

	_Check_return_ ULONG ListPane::GetItemCount() const { return m_List.GetItemCount(); }

	_Check_return_ controls::sortlistdata::SortListData* ListPane::GetItemData(int iRow) const
	{
		return reinterpret_cast<controls::sortlistdata::SortListData*>(m_List.GetItemData(iRow));
	}

	_Check_return_ controls::sortlistdata::SortListData* ListPane::GetSelectedListRowData() const
	{
		const auto iItem = m_List.GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED);

		if (-1 != iItem)
		{
			return GetItemData(iItem);
		}

		return nullptr;
	}

	void ListPane::InsertColumn(int nCol, UINT uidText) { m_List.InsertColumnW(nCol, strings::loadstring(uidText)); }

	void ListPane::SetColumnType(int nCol, ULONG ulPropType) const
	{
		HDITEM hdItem = {};
		auto lpMyHeader = m_List.GetHeaderCtrl();

		if (lpMyHeader)
		{
			hdItem.mask = HDI_LPARAM;
			auto lpHeaderData = new (std::nothrow)
				controls::sortlistctrl::HeaderData; // Will be deleted in CSortListCtrl::DeleteAllColumns
			if (lpHeaderData)
			{
				lpHeaderData->ulTagArrayRow = NULL;
				lpHeaderData->ulPropTag = PROP_TAG(ulPropType, PROP_ID_NULL);
				lpHeaderData->bIsAB = false;
				lpHeaderData->szTipString[0] = NULL;
				hdItem.lParam = reinterpret_cast<LPARAM>(lpHeaderData);
				EC_B_S(lpMyHeader->SetItem(nCol, &hdItem));
			}
		}
	}

	void ListPane::UpdateButtons()
	{
		const ULONG ulNumItems = m_List.GetItemCount();

		for (auto iButton = 0; iButton < NUMLISTBUTTONS; iButton++)
		{
			switch (ListButtons[iButton])
			{
			case IDD_LISTMOVETOBOTTOM:
			case IDD_LISTMOVEDOWN:
			case IDD_LISTMOVETOTOP:
			case IDD_LISTMOVEUP:
				EC_B_S(m_ButtonArray[iButton].EnableWindow(ulNumItems >= 2));
				break;
			case IDD_LISTDELETE:
				EC_B_S(m_ButtonArray[iButton].EnableWindow(ulNumItems >= 1));
				break;
			case IDD_LISTEDIT:
				EC_B_S(m_ButtonArray[iButton].EnableWindow(ulNumItems >= 1));
				break;
			}
		}
	}

	void ListPane::SwapListItems(ULONG ulFirstItem, ULONG ulSecondItem)
	{
		auto lpData1 = GetItemData(ulFirstItem);
		auto lpData2 = GetItemData(ulSecondItem);

		// swap the data
		EC_B_S(m_List.SetItemData(ulFirstItem, reinterpret_cast<DWORD_PTR>(lpData2)));
		EC_B_S(m_List.SetItemData(ulSecondItem, reinterpret_cast<DWORD_PTR>(lpData1)));

		// swap the text (skip the first column!)
		const auto lpMyHeader = m_List.GetHeaderCtrl();
		if (lpMyHeader)
		{
			for (auto i = 1; i < lpMyHeader->GetItemCount(); i++)
			{
				const auto szText1 = GetItemText(ulFirstItem, i);
				const auto szText2 = GetItemText(ulSecondItem, i);
				m_List.SetItemText(ulFirstItem, i, szText2);
				m_List.SetItemText(ulSecondItem, i, szText1);
			}
		}
	}

	void ListPane::OnMoveListEntryUp()
	{
		const auto iItem = m_List.GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED);
		output::DebugPrintEx(DBGGeneric, CLASS, L"OnMoveListEntryUp", L"This item was selected: 0x%08X\n", iItem);

		if (-1 == iItem) return;
		if (0 == iItem) return;

		SwapListItems(iItem - 1, iItem);
		m_List.SetSelectedItem(iItem - 1);
		m_bDirty = true;
	}

	void ListPane::OnMoveListEntryDown()
	{
		const auto iItem = m_List.GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED);
		output::DebugPrintEx(DBGGeneric, CLASS, L"OnMoveListEntryDown", L"This item was selected: 0x%08X\n", iItem);

		if (-1 == iItem) return;
		if (m_List.GetItemCount() == iItem + 1) return;

		SwapListItems(iItem, iItem + 1);
		m_List.SetSelectedItem(iItem + 1);
		m_bDirty = true;
	}

	void ListPane::OnMoveListEntryToTop()
	{
		const auto iItem = m_List.GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED);
		output::DebugPrintEx(DBGGeneric, CLASS, L"OnMoveListEntryToTop", L"This item was selected: 0x%08X\n", iItem);

		if (-1 == iItem) return;
		if (0 == iItem) return;

		for (auto i = iItem; i > 0; i--)
		{
			SwapListItems(i, i - 1);
		}

		m_List.SetSelectedItem(iItem);
		m_bDirty = true;
	}

	void ListPane::OnMoveListEntryToBottom()
	{
		const auto iItem = m_List.GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED);
		output::DebugPrintEx(DBGGeneric, CLASS, L"OnMoveListEntryDown", L"This item was selected: 0x%08X\n", iItem);

		if (-1 == iItem) return;
		if (m_List.GetItemCount() == iItem + 1) return;

		for (auto i = iItem; i < m_List.GetItemCount() - 1; i++)
		{
			SwapListItems(i, i + 1);
		}

		m_List.SetSelectedItem(iItem);
		m_bDirty = true;
	}

	void ListPane::OnAddListEntry()
	{
		const auto iItem = m_List.GetItemCount();

		(void) InsertRow(iItem, std::to_wstring(iItem));

		m_List.SetSelectedItem(iItem);

		const auto bDidEdit = OnEditListEntry();

		// if we didn't do anything in the edit, undo the add
		// pass false to make sure we don't mark the list dirty if it wasn't already
		if (!bDidEdit) OnDeleteListEntry(false);

		UpdateButtons();
	}

	void ListPane::OnDeleteListEntry(bool bDoDirty)
	{
		const auto iItem = m_List.GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED);
		output::DebugPrintEx(DBGGeneric, CLASS, L"OnDeleteListEntry", L"This item was selected: 0x%08X\n", iItem);

		if (iItem == -1) return;

		const auto hRes = EC_B(m_List.DeleteItem(iItem));
		m_List.SetSelectedItem(iItem);

		if (hRes == S_OK && bDoDirty)
		{
			m_bDirty = true;
		}

		UpdateButtons();
	}

	_Check_return_ bool ListPane::OnEditListEntry()
	{
		const auto iItem = m_List.GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED);
		output::DebugPrintEx(DBGGeneric, CLASS, L"OnEditListEntry", L"This item was selected: 0x%08X\n", iItem);

		if (iItem == -1) return false;

		const auto lpData = GetItemData(iItem);
		if (!lpData) return false;

		auto bDidEdit = false;
		if (m_callback)
		{
			bDidEdit = m_callback(m_paneID, iItem, lpData);
		}

		// the list is dirty now if the edit succeeded or it was already dirty
		m_bDirty = bDidEdit || m_bDirty;
		return bDidEdit;
	}

	std::wstring ListPane::GetItemText(_In_ int nItem, _In_ int nSubItem) const
	{
		return m_List.GetItemText(nItem, nSubItem);
	}
} // namespace viewpane