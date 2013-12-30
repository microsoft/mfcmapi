#include "stdafx.h"
#include "..\stdafx.h"
#include "ListPane.h"
#include "..\Editor.h"

ViewPane* CreateListPane(UINT uidLabel, bool bAllowSort, bool bReadOnly, LPVOID lpEdit)
{
	return new ListPane(uidLabel, bReadOnly, bAllowSort, (CEditor*) lpEdit);
}

static TCHAR* CLASS = _T("ListPane");

__ListButtons ListButtons[NUMLISTBUTTONS] = {
	{IDD_LISTMOVEDOWN},
	{IDD_LISTMOVETOBOTTOM},
	{IDD_LISTADD},
	{IDD_LISTEDIT},
	{IDD_LISTDELETE},
	{IDD_LISTMOVETOTOP},
	{IDD_LISTMOVEUP},
};

#define LINES_LIST 6

ListPane::ListPane(UINT uidLabel, bool bReadOnly, bool bAllowSort, CEditor* lpEdit):ViewPane(uidLabel, bReadOnly)
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
	int iHeight = 0;
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
	switch(nID)
	{
	case IDD_LISTMOVEDOWN:		OnMoveListEntryDown(); break;
	case IDD_LISTADD:			OnAddListEntry(); break;
	case IDD_LISTEDIT:			(void) OnEditListEntry(); break;
	case IDD_LISTDELETE:		OnDeleteListEntry(true); break;
	case IDD_LISTMOVEUP:		OnMoveListEntryUp(); break;
	case IDD_LISTMOVETOBOTTOM:	OnMoveListEntryToBottom(); break;
	case IDD_LISTMOVETOTOP:		OnMoveListEntryToTop(); break;
	default: return (ULONG) -1;
	}
	return m_iControl;
}

void ListPane::SetWindowPos(int x, int y, int width, int height)
{
	HRESULT hRes = S_OK;

	int variableHeight = height - GetFixedHeight();

	if (0 != m_iControl)
	{
		y += m_iSmallHeightMargin;
		height -= m_iSmallHeightMargin;
	}

	if (m_bUseLabelControl)
	{
		EC_B(m_Label.SetWindowPos(
			0,
			x,
			y,
			width,
			m_iLabelHeight,
			SWP_NOZORDER));
		y += m_iLabelHeight;
		height -= m_iLabelHeight;
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

		int iSlotWidth = m_iButtonWidth + m_iMargin;
		int iOffset = width + m_iSideMargin + m_iMargin;
		int iButton = 0;

		for (iButton = 0 ; iButton < NUMLISTBUTTONS ; iButton++)
		{
			EC_B(m_ButtonArray[iButton].SetWindowPos(
				0,
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
	ViewPane::Initialize(iControl, pParent, NULL);

	HRESULT hRes = S_OK;

	SIZE sizeText = {0};
	DWORD dwListStyle = LVS_SINGLESEL | WS_BORDER;
	if (!m_bAllowSort)
		dwListStyle |= LVS_NOSORTHEADER;
	EC_H(m_List.Create(pParent, dwListStyle, m_nID, false));

	// read only lists don't need buttons
	if (!m_bReadOnly)
	{
		int iButton = 0;

		for (iButton = 0; iButton < NUMLISTBUTTONS; iButton++)
		{
			CString szButtonText;
			EC_B(szButtonText.LoadString(ListButtons[iButton].uiButtonID));

			EC_B(m_ButtonArray[iButton].Create(
				szButtonText,
				WS_TABSTOP
				| WS_CHILD
				| WS_CLIPSIBLINGS
				| WS_VISIBLE,
				CRect(0,0,0,0),
				pParent,
				ListButtons[iButton].uiButtonID));

			::GetTextExtentPoint32(hdc, szButtonText, szButtonText.GetLength(), &sizeText);
			m_iButtonWidth = max(m_iButtonWidth,sizeText.cx);
		}
	}
	m_iButtonWidth += m_iMargin;
	m_bInitialized = true;
}

void ListPane::SetListStringA(ULONG iListRow, ULONG iListCol, _In_opt_z_ LPCSTR szListString)
{
	m_List.SetItemTextA(iListRow, iListCol, szListString?szListString:"");
}

void ListPane::SetListStringW(ULONG iListRow, ULONG iListCol, _In_opt_z_ LPCWSTR szListString)
{
	m_List.SetItemTextW(iListRow, iListCol, szListString?szListString:L"");
}

_Check_return_ SortListData* ListPane::InsertRow(int iRow, _In_z_ LPCTSTR szText)
{
	return m_List.InsertRow(iRow, (LPTSTR) szText);
}

void ListPane::ClearList()
{
	HRESULT hRes = S_OK;

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

_Check_return_ ULONG ListPane::GetItemCount()
{
	return m_List.GetItemCount();
}

_Check_return_ SortListData* ListPane::GetItemData(int iRow)
{
	return (SortListData*) m_List.GetItemData(iRow);
}

_Check_return_ SortListData* ListPane::GetSelectedListRowData()
{
	int	iItem = m_List.GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED);

	if (-1 != iItem)
	{
		return GetItemData(iItem);
	}
	return NULL;
}

void ListPane::InsertColumn(int nCol, UINT uidText)
{
	CString szText;
	HRESULT hRes = S_OK;
	EC_B(szText.LoadString(uidText));

	m_List.InsertColumn(nCol, szText);
}

void ListPane::SetColumnType(int nCol, ULONG ulPropType)
{
	HRESULT hRes = S_OK;
	CHeaderCtrl* lpMyHeader = NULL;
	HDITEM hdItem = {0};
	lpMyHeader = m_List.GetHeaderCtrl();

	if (lpMyHeader)
	{
		hdItem.mask = HDI_LPARAM;
		LPHEADERDATA lpHeaderData = NULL;

		lpHeaderData = new HeaderData; // Will be deleted in CSortListCtrl::DeleteAllColumns
		if (lpHeaderData)
		{
			lpHeaderData->ulTagArrayRow = NULL;
			lpHeaderData->ulPropTag = PROP_TAG(ulPropType, PROP_ID_NULL);
			lpHeaderData->bIsAB = false;
			lpHeaderData->szTipString[0] = NULL;
			hdItem.lParam = (LPARAM) lpHeaderData;
			EC_B(lpMyHeader->SetItem(nCol, &hdItem));
		}
	}
}

void ListPane::UpdateListButtons()
{
	ULONG ulNumItems = m_List.GetItemCount();

	HRESULT hRes = S_OK;
	int iButton = 0;

	for (iButton = 0; iButton <NUMLISTBUTTONS; iButton++)
	{
		switch(ListButtons[iButton].uiButtonID)
		{
		case IDD_LISTMOVETOBOTTOM:
		case IDD_LISTMOVEDOWN:
		case IDD_LISTMOVETOTOP:
		case IDD_LISTMOVEUP:
			{
				EC_B(m_ButtonArray[iButton].EnableWindow(ulNumItems >= 2?true:false));
				break;
			}
		case IDD_LISTDELETE:
			{
				EC_B(m_ButtonArray[iButton].EnableWindow(ulNumItems >= 1?true:false));
				break;
			}
		case IDD_LISTEDIT:
			{
				EC_B(m_ButtonArray[iButton].EnableWindow(ulNumItems >= 1?true:false));
				break;
			}
		}
	}
}

void ListPane::SwapListItems(ULONG ulFirstItem, ULONG ulSecondItem)
{
	HRESULT hRes = S_OK;
	SortListData* lpData1 = GetItemData(ulFirstItem);
	SortListData* lpData2 = GetItemData(ulSecondItem);

	// swap the data
	EC_B(m_List.SetItemData(ulFirstItem, (DWORD_PTR) lpData2));
	EC_B(m_List.SetItemData(ulSecondItem, (DWORD_PTR) lpData1));

	// swap the text (skip the first column!)
	CHeaderCtrl* lpMyHeader = NULL;
	lpMyHeader = m_List.GetHeaderCtrl();
	if (lpMyHeader)
	{
		for (int i = 1;i<lpMyHeader->GetItemCount();i++)
		{
			CString szText1;
			CString szText2;

			szText1	= m_List.GetItemText(ulFirstItem, i);
			szText2	= m_List.GetItemText(ulSecondItem, i);
			m_List.SetItemText(ulFirstItem, i, szText2);
			m_List.SetItemText(ulSecondItem, i, szText1);
		}
	}
}

void ListPane::OnMoveListEntryUp()
{
	int iItem = NULL;

	iItem = m_List.GetNextItem(-1,LVNI_FOCUSED | LVNI_SELECTED);
	DebugPrintEx(DBGGeneric, CLASS, _T("OnMoveListEntryUp"),_T("This item was selected: 0x%08X\n"), iItem);

	if (-1 == iItem) return;
	if (0 == iItem) return;

	SwapListItems(iItem - 1, iItem);
	m_List.SetSelectedItem(iItem - 1);
	m_bDirty = true;
}

void ListPane::OnMoveListEntryDown()
{
	int iItem = NULL;

	iItem = m_List.GetNextItem(-1,LVNI_FOCUSED | LVNI_SELECTED);
	DebugPrintEx(DBGGeneric,CLASS, _T("OnMoveListEntryDown"),_T("This item was selected: 0x%08X\n"), iItem);

	if (-1 == iItem) return;
	if (m_List.GetItemCount() == iItem+1) return;

	SwapListItems(iItem, iItem + 1);
	m_List.SetSelectedItem(iItem + 1);
	m_bDirty = true;
}

void ListPane::OnMoveListEntryToTop()
{
	int iItem = NULL;

	iItem = m_List.GetNextItem(-1,LVNI_FOCUSED | LVNI_SELECTED);
	DebugPrintEx(DBGGeneric,CLASS, _T("OnMoveListEntryToTop"),_T("This item was selected: 0x%08X\n"), iItem);

	if (-1 == iItem) return;
	if (0 == iItem) return;

	int i = 0;
	for (i = iItem ; i >0 ; i--)
	{
		SwapListItems(i, i - 1);
	}
	m_List.SetSelectedItem(iItem);
	m_bDirty = true;
}

void ListPane::OnMoveListEntryToBottom()
{
	int iItem = NULL;

	iItem = m_List.GetNextItem(-1,LVNI_FOCUSED | LVNI_SELECTED);
	DebugPrintEx(DBGGeneric,CLASS, _T("OnMoveListEntryDown"),_T("This item was selected: 0x%08X\n"), iItem);

	if (-1 == iItem) return;
	if (m_List.GetItemCount() == iItem+1) return;

	int i = 0;
	for (i = iItem ; i < m_List.GetItemCount() - 1 ; i++)
	{
		SwapListItems(i, i + 1);
	}
	m_List.SetSelectedItem(iItem);
	m_bDirty = true;
}

void ListPane::OnAddListEntry()
{
	int iItem = m_List.GetItemCount();

	CString szTmp;
	szTmp.Format(_T("%d"), iItem); // STRING_OK

	SortListData* lpData = InsertRow(iItem, szTmp);
	if (lpData) lpData->ulSortDataType = SORTLIST_MVPROP;

	m_List.SetSelectedItem(iItem);

	bool bDidEdit = OnEditListEntry();

	// if we didn't do anything in the edit, undo the add
	// pass false to make sure we don't mark the list dirty if it wasn't already
	if (!bDidEdit) OnDeleteListEntry(false);

	UpdateListButtons();
}

void ListPane::OnDeleteListEntry(bool bDoDirty)
{
	int iItem = NULL;

	iItem = m_List.GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED);
	DebugPrintEx(DBGGeneric, CLASS, _T("OnDeleteListEntry"), _T("This item was selected: 0x%08X\n"), iItem);

	if (iItem == -1) return;

	HRESULT hRes = S_OK;

	EC_B(m_List.DeleteItem(iItem));
	m_List.SetSelectedItem(iItem);

	if (S_OK == hRes && bDoDirty)
	{
		m_bDirty = true;
	}
	UpdateListButtons();
} // ListPane::OnDeleteListEntry

_Check_return_ bool ListPane::OnEditListEntry()
{
	int iItem = NULL;

	iItem = m_List.GetNextItem(-1, LVNI_FOCUSED | LVNI_SELECTED);
	DebugPrintEx(DBGGeneric,CLASS, _T("OnEditListEntry"), _T("This item was selected: 0x%08X\n"), iItem);

	if (iItem == -1) return false;

	SortListData* lpData = GetItemData(iItem);

	if (!lpData) return false;

	// TODO: Figure out a way to do this that doesn't involve caching the edit control
	bool bDidEdit = m_lpEdit->DoListEdit(m_iControl, iItem, lpData);

	// the list is dirty now if the edit succeeded or it was already dirty
	m_bDirty = bDidEdit || m_bDirty;
	return bDidEdit;
}

CString ListPane::GetItemText(_In_ int nItem, _In_ int nSubItem)
{
	return m_List.GetItemText(nItem, nSubItem);
}
