#include "stdafx.h"
#include "ListPane.h"
#include "Editor.h"

ViewPane* CreateListPane(UINT uidLabel, bool bAllowSort, bool bReadOnly, LPVOID lpEdit)
{
	return new ListPane(uidLabel, bReadOnly, bAllowSort, static_cast<CEditor*>(lpEdit));
}

static wstring CLASS = L"ListPane";

__ListButtons ListButtons[NUMLISTBUTTONS] = {
 { IDD_LISTMOVEDOWN },
 { IDD_LISTMOVETOBOTTOM },
 { IDD_LISTADD },
 { IDD_LISTEDIT },
 { IDD_LISTDELETE },
 { IDD_LISTMOVETOTOP },
 { IDD_LISTMOVEUP },
};

#define LINES_LIST 6

ListPane::ListPane(UINT uidLabel, bool bReadOnly, bool bAllowSort, CEditor* lpEdit) :ViewPane(uidLabel, bReadOnly), m_bDirty(false)
{
	m_bAllowSort = bAllowSort;
	m_lpEdit = lpEdit;
	m_iButtonWidth = 50;
	m_List.AllowEscapeClose();
}

bool ListPane::IsType(__ViewTypes vType)
{
	return CTRL_LISTPANE == vType;
}

ULONG ListPane::GetFlags()
{
	ULONG ulFlags = vpNone;
	if (m_bDirty) ulFlags |= vpDirty;
	if (m_bReadOnly) ulFlags |= vpReadonly;
	return ulFlags;
}

int ListPane::GetMinWidth(_In_ HDC hdc)
{
	return max(ViewPane::GetMinWidth(hdc), (int)(NUMLISTBUTTONS * m_iButtonWidth + m_iMargin * (NUMLISTBUTTONS - 1)));
}

int ListPane::GetFixedHeight()
{
	auto iHeight = 0;
	if (0 != m_iControl) iHeight += m_iSmallHeightMargin; // Top margin

	if (m_bUseLabelControl)
	{
		iHeight += m_iLabelHeight;
	}

	if (!m_bReadOnly)
	{
		iHeight += m_iLargeHeightMargin + m_iButtonHeight;
	}

	return iHeight;
}

int ListPane::GetLines()
{
	return LINES_LIST;
}

ULONG ListPane::HandleChange(UINT nID)
{
	switch (nID)
	{
	case IDD_LISTMOVEDOWN: OnMoveListEntryDown(); break;
	case IDD_LISTADD: OnAddListEntry(); break;
	case IDD_LISTEDIT: (void)OnEditListEntry(); break;
	case IDD_LISTDELETE: OnDeleteListEntry(true); break;
	case IDD_LISTMOVEUP: OnMoveListEntryUp(); break;
	case IDD_LISTMOVETOBOTTOM: OnMoveListEntryToBottom(); break;
	case IDD_LISTMOVETOTOP: OnMoveListEntryToTop(); break;
	default: return static_cast<ULONG>(-1);
	}

	return m_iControl;
}

void ListPane::SetWindowPos(int x, int y, int width, int height)
{
	auto hRes = S_OK;

	auto variableHeight = height - GetFixedHeight();

	if (0 != m_iControl)
	{
		y += m_iSmallHeightMargin;
		// height -= m_iSmallHeightMargin;
	}

	if (m_bUseLabelControl)
	{
		EC_B(m_Label.SetWindowPos(
			nullptr,
			x,
			y,
			width,
			m_iLabelHeight,
			SWP_NOZORDER));
		y += m_iLabelHeight;
		// height -= m_iLabelHeight;
	}

	EC_B(m_List.SetWindowPos(
		NULL,
		x,
		y,
		width,
		variableHeight,
		SWP_NOZORDER));
	y += variableHeight;

	if (!m_bReadOnly)
	{
		// buttons go below the list:
		y += m_iLargeHeightMargin;

		auto iSlotWidth = m_iButtonWidth + m_iMargin;
		auto iOffset = width + m_iSideMargin + m_iMargin;

		for (auto iButton = 0; iButton < NUMLISTBUTTONS; iButton++)
		{
			EC_B(m_ButtonArray[iButton].SetWindowPos(
				nullptr,
				iOffset - iSlotWidth * (NUMLISTBUTTONS - iButton),
				y,
				m_iButtonWidth,
				m_iButtonHeight,
				SWP_NOZORDER));
		}
	}
}

void ListPane::Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc)
{
	ViewPane::Initialize(iControl, pParent, nullptr);

	auto hRes = S_OK;

	SIZE sizeText = { 0 };
	DWORD dwListStyle = LVS_SINGLESEL | WS_BORDER;
	if (!m_bAllowSort)
		dwListStyle |= LVS_NOSORTHEADER;
	EC_H(m_List.Create(pParent, dwListStyle, m_nID, false));

	// read only lists don't need buttons
	if (!m_bReadOnly)
	{
		for (auto iButton = 0; iButton < NUMLISTBUTTONS; iButton++)
		{
			CString szButtonText;
			EC_B(szButtonText.LoadString(ListButtons[iButton].uiButtonID));

			EC_B(m_ButtonArray[iButton].Create(
				szButtonText,
				WS_TABSTOP
				| WS_CHILD
				| WS_CLIPSIBLINGS
				| WS_VISIBLE,
				CRect(0, 0, 0, 0),
				pParent,
				ListButtons[iButton].uiButtonID));

			::GetTextExtentPoint32(hdc, szButtonText, szButtonText.GetLength(), &sizeText);
			m_iButtonWidth = max(m_iButtonWidth, sizeText.cx);
		}
	}

	m_iButtonWidth += m_iMargin;
	m_bInitialized = true;
}

void ListPane::SetListString(ULONG iListRow, ULONG iListCol, wstring szListString)
{
	m_List.SetItemText(iListRow, iListCol, szListString);
}

_Check_return_ SortListData* ListPane::InsertRow(int iRow, wstring szText) const
{
	return m_List.InsertRow(iRow, szText);
}

void ListPane::ClearList()
{
	auto hRes = S_OK;

	m_List.DeleteAllColumns();
	EC_B(m_List.DeleteAllItems());
}

void ListPane::ResizeList(bool bSort)
{
	if (bSort)
	{
		m_List.SortClickedColumn();
	}

	m_List.AutoSizeColumns(false);
}

_Check_return_ ULONG ListPane::GetItemCount() const
{
	return m_List.GetItemCount();
}

_Check_return_ SortListData* ListPane::GetItemData(int iRow) const
{
	return reinterpret_cast<SortListData*>(m_List.GetItemData(iRow));
}

_Check_return_ SortListData* ListPane::GetSelectedListRowData() const
{
	auto iItem = m_List.GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED);

	if (-1 != iItem)
	{
		return GetItemData(iItem);
	}

	return nullptr;
}

void ListPane::InsertColumn(int nCol, UINT uidText)
{
	CString szText;
	auto hRes = S_OK;
	EC_B(szText.LoadString(uidText));

	m_List.InsertColumn(nCol, szText);
}

void ListPane::SetColumnType(int nCol, ULONG ulPropType) const
{
	auto hRes = S_OK;
	HDITEM hdItem = { 0 };
	auto lpMyHeader = m_List.GetHeaderCtrl();

	if (lpMyHeader)
	{
		hdItem.mask = HDI_LPARAM;
		auto lpHeaderData = new HeaderData; // Will be deleted in CSortListCtrl::DeleteAllColumns
		if (lpHeaderData)
		{
			lpHeaderData->ulTagArrayRow = NULL;
			lpHeaderData->ulPropTag = PROP_TAG(ulPropType, PROP_ID_NULL);
			lpHeaderData->bIsAB = false;
			lpHeaderData->szTipString[0] = NULL;
			hdItem.lParam = reinterpret_cast<LPARAM>(lpHeaderData);
			EC_B(lpMyHeader->SetItem(nCol, &hdItem));
		}
	}
}

void ListPane::UpdateListButtons()
{
	ULONG ulNumItems = m_List.GetItemCount();

	auto hRes = S_OK;
	for (auto iButton = 0; iButton < NUMLISTBUTTONS; iButton++)
	{
		switch (ListButtons[iButton].uiButtonID)
		{
		case IDD_LISTMOVETOBOTTOM:
		case IDD_LISTMOVEDOWN:
		case IDD_LISTMOVETOTOP:
		case IDD_LISTMOVEUP:
			EC_B(m_ButtonArray[iButton].EnableWindow(ulNumItems >= 2 ? true : false));
			break;
		case IDD_LISTDELETE:
			EC_B(m_ButtonArray[iButton].EnableWindow(ulNumItems >= 1 ? true : false));
			break;
		case IDD_LISTEDIT:
			EC_B(m_ButtonArray[iButton].EnableWindow(ulNumItems >= 1 ? true : false));
			break;
		}
	}
}

void ListPane::SwapListItems(ULONG ulFirstItem, ULONG ulSecondItem)
{
	auto hRes = S_OK;
	auto lpData1 = GetItemData(ulFirstItem);
	auto lpData2 = GetItemData(ulSecondItem);

	// swap the data
	EC_B(m_List.SetItemData(ulFirstItem, reinterpret_cast<DWORD_PTR>(lpData2)));
	EC_B(m_List.SetItemData(ulSecondItem, reinterpret_cast<DWORD_PTR>(lpData1)));

	// swap the text (skip the first column!)
	auto lpMyHeader = m_List.GetHeaderCtrl();
	if (lpMyHeader)
	{
		for (auto i = 1; i < lpMyHeader->GetItemCount(); i++)
		{
			auto szText1 = GetItemText(ulFirstItem, i);
			auto szText2 = GetItemText(ulSecondItem, i);
			m_List.SetItemText(ulFirstItem, i, szText2);
			m_List.SetItemText(ulSecondItem, i, szText1);
		}
	}
}

void ListPane::OnMoveListEntryUp()
{
	auto iItem = m_List.GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED);
	DebugPrintEx(DBGGeneric, CLASS, L"OnMoveListEntryUp", L"This item was selected: 0x%08X\n", iItem);

	if (-1 == iItem) return;
	if (0 == iItem) return;

	SwapListItems(iItem - 1, iItem);
	m_List.SetSelectedItem(iItem - 1);
	m_bDirty = true;
}

void ListPane::OnMoveListEntryDown()
{
	auto iItem = m_List.GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED);
	DebugPrintEx(DBGGeneric, CLASS, L"OnMoveListEntryDown", L"This item was selected: 0x%08X\n", iItem);

	if (-1 == iItem) return;
	if (m_List.GetItemCount() == iItem + 1) return;

	SwapListItems(iItem, iItem + 1);
	m_List.SetSelectedItem(iItem + 1);
	m_bDirty = true;
}

void ListPane::OnMoveListEntryToTop()
{
	auto iItem = m_List.GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED);
	DebugPrintEx(DBGGeneric, CLASS, L"OnMoveListEntryToTop", L"This item was selected: 0x%08X\n", iItem);

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
	auto iItem = m_List.GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED);
	DebugPrintEx(DBGGeneric, CLASS, L"OnMoveListEntryDown", L"This item was selected: 0x%08X\n", iItem);

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
	auto iItem = m_List.GetItemCount();

	auto lpData = InsertRow(iItem, format(L"%d", iItem)); // STRING_OK
	if (lpData)
	{
		lpData->InitializePropList(0);
	}

	m_List.SetSelectedItem(iItem);

	auto bDidEdit = OnEditListEntry();

	// if we didn't do anything in the edit, undo the add
	// pass false to make sure we don't mark the list dirty if it wasn't already
	if (!bDidEdit) OnDeleteListEntry(false);

	UpdateListButtons();
}

void ListPane::OnDeleteListEntry(bool bDoDirty)
{
	auto iItem = m_List.GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED);
	DebugPrintEx(DBGGeneric, CLASS, L"OnDeleteListEntry", L"This item was selected: 0x%08X\n", iItem);

	if (iItem == -1) return;

	auto hRes = S_OK;

	EC_B(m_List.DeleteItem(iItem));
	m_List.SetSelectedItem(iItem);

	if (S_OK == hRes && bDoDirty)
	{
		m_bDirty = true;
	}

	UpdateListButtons();
}

_Check_return_ bool ListPane::OnEditListEntry()
{
	auto iItem = m_List.GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED);
	DebugPrintEx(DBGGeneric, CLASS, L"OnEditListEntry", L"This item was selected: 0x%08X\n", iItem);

	if (iItem == -1) return false;

	auto lpData = GetItemData(iItem);
	if (!lpData) return false;

	// TODO: Figure out a way to do this that doesn't involve caching the edit control
	auto bDidEdit = m_lpEdit->DoListEdit(m_iControl, iItem, lpData);

	// the list is dirty now if the edit succeeded or it was already dirty
	m_bDirty = bDidEdit || m_bDirty;
	return bDidEdit;
}

wstring ListPane::GetItemText(_In_ int nItem, _In_ int nSubItem) const
{
	return m_List.GetItemText(nItem, nSubItem);
}
