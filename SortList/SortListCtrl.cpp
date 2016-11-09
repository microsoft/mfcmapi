#include "stdafx.h"
#include "SortListCtrl.h"
#include "SortHeader.h"
#include "String.h"
#include "UIFunctions.h"

static wstring CLASS = L"CSortListCtrl";

CSortListCtrl::CSortListCtrl()
{
	TRACE_CONSTRUCTOR(CLASS);
	m_cRef = 1;
	m_iRedrawCount = 0;
	m_iClickedColumn = 0;
	m_bSortUp = false;
	m_bHaveSorted = false;
	m_bHeaderSubclassed = false;
	m_iItemCurHover = -1;
	m_bAllowEscapeClose = false;
}

CSortListCtrl::~CSortListCtrl()
{
	TRACE_DESTRUCTOR(CLASS);
	CWnd::DestroyWindow();
}

STDMETHODIMP_(ULONG) CSortListCtrl::AddRef()
{
	auto lCount(InterlockedIncrement(&m_cRef));
	TRACE_ADDREF(CLASS, lCount);
	return lCount;
}

STDMETHODIMP_(ULONG) CSortListCtrl::Release()
{
	auto lCount(InterlockedDecrement(&m_cRef));
	TRACE_RELEASE(CLASS, lCount);
	if (!lCount) delete this;
	return lCount;
}

BEGIN_MESSAGE_MAP(CSortListCtrl, CListCtrl)
	ON_WM_KEYDOWN()
	ON_WM_GETDLGCODE()
	ON_WM_DRAWITEM()
	ON_NOTIFY_REFLECT(LVN_DELETEALLITEMS, OnDeleteAllItems)
	ON_NOTIFY_REFLECT(LVN_DELETEITEM, OnDeleteItem)
	ON_NOTIFY_REFLECT(NM_CUSTOMDRAW, OnCustomDraw)
END_MESSAGE_MAP()

_Check_return_ HRESULT CSortListCtrl::Create(_In_ CWnd* pCreateParent, ULONG ulFlags, UINT nID, bool bImages)
{
	auto hRes = S_OK;
	EC_B(CListCtrl::Create(
		ulFlags
		| LVS_REPORT
		| LVS_SHOWSELALWAYS
		| WS_TABSTOP
		| WS_CHILD
		| WS_CLIPCHILDREN
		| WS_CLIPSIBLINGS
		| WS_VISIBLE,
		CRect(0, 0, 0, 0), // size doesn't matter
		pCreateParent,
		nID));
	ListView_SetBkColor(m_hWnd, MyGetSysColor(cBackground));
	ListView_SetTextBkColor(m_hWnd, MyGetSysColor(cBackground));
	ListView_SetTextColor(m_hWnd, MyGetSysColor(cText));
	::SendMessageA(m_hWnd, WM_SETFONT, reinterpret_cast<WPARAM>(GetSegoeFont()), false);

	SetExtendedStyle(GetExtendedStyle() | LVS_EX_FULLROWSELECT | LVS_EX_HEADERDRAGDROP | LVS_EX_INFOTIP | LVS_EX_DOUBLEBUFFER);

	if (bImages)
	{
		auto hImageList = ImageList_Create(16, 16, ILC_COLOR32 | ILC_MASK, 1, 1);
		m_ImageList.Attach(hImageList);

		CBitmap myBitmap;
		myBitmap.LoadBitmap(IDB_ICONS);
		m_ImageList.Add(&myBitmap, MyGetSysColor(cBitmapTransBack));

		SetImageList(&m_ImageList, LVSIL_SMALL);
	}

	return hRes;
}

static bool s_bInTrack = false;
static int s_iTrack = 0;
static int s_iHeaderHeight = 0;

void OnBeginTrack(_In_ NMHDR* pNMHDR, _In_ HWND hWndParent)
{
	RECT rcHeader = { 0 };
	if (!pNMHDR) return;
	LPNMHEADER pHdr = reinterpret_cast<LPNMHEADER>(pNMHDR);
	Header_GetItemRect(pHdr->hdr.hwndFrom, pHdr->iItem, &rcHeader);
	s_bInTrack = true;
	s_iTrack = rcHeader.right;
	s_iHeaderHeight = rcHeader.bottom - rcHeader.top;
	DrawTrackingBar(pHdr->hdr.hwndFrom, hWndParent, s_iTrack, s_iHeaderHeight, false);
}

void OnEndTrack(_In_ NMHDR* pNMHDR, _In_ HWND hWndParent)
{
	if (s_bInTrack && pNMHDR)
	{
		DrawTrackingBar(pNMHDR->hwndFrom, hWndParent, s_iTrack, s_iHeaderHeight, true);
	}
	s_bInTrack = false;
}

void OnTrack(_In_ NMHDR* pNMHDR, _In_ HWND hWndParent)
{
	if (s_bInTrack && pNMHDR)
	{
		RECT rcHeader = { 0 };
		LPNMHEADER pHdr = reinterpret_cast<LPNMHEADER>(pNMHDR);
		Header_GetItemRect(pHdr->hdr.hwndFrom, pHdr->iItem, &rcHeader);
		if (s_iTrack != rcHeader.right)
		{
			DrawTrackingBar(pHdr->hdr.hwndFrom, hWndParent, s_iTrack, s_iHeaderHeight, true);
			s_iTrack = rcHeader.right;
			DrawTrackingBar(pHdr->hdr.hwndFrom, hWndParent, s_iTrack, s_iHeaderHeight, false);
		}
	}
}

LRESULT CSortListCtrl::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	auto iItemCur = m_iItemCurHover;
	auto hRes = S_OK;

	switch (message)
	{
		// I can handle notify messages for my child header control since I am the parent window
		// This makes it easy for me to customize the child control to do what I want
		// I cannot handle notify heading to my parent though - have to depend on reflection for that
	case WM_NOTIFY:
	{
		auto pHdr = reinterpret_cast<LPNMHDR>(lParam);

		switch (pHdr->code)
		{
		case HDN_ITEMCLICKA:
		case HDN_ITEMCLICKW:
			OnColumnClick(reinterpret_cast<LPNMHEADERW>(pHdr)->iItem);
			return NULL;
		case HDN_DIVIDERDBLCLICKA:
		case HDN_DIVIDERDBLCLICKW:
			AutoSizeColumn(reinterpret_cast<LPNMHEADERW>(pHdr)->iItem, 0, 0);
			return NULL;
		case HDN_BEGINTRACKA:
		case HDN_BEGINTRACKW:
			OnBeginTrack(pHdr, m_hWnd);
			break;
		case HDN_ENDTRACKA:
		case HDN_ENDTRACKW:
			OnEndTrack(pHdr, m_hWnd);
			break;
		case HDN_ITEMCHANGEDA:
		case HDN_ITEMCHANGEDW:
			// Let the control handle the resized columns before we redraw our tracking bar
			auto lRet = CListCtrl::WindowProc(message, wParam, lParam);
			OnTrack(pHdr, m_hWnd);
			return lRet;
		}
		break; // WM_NOTIFY
	}
	case WM_SHOWWINDOW:
		// subclass the header
		if (!m_bHeaderSubclassed)
		{
			m_bHeaderSubclassed = m_cSortHeader.Init(GetHeaderCtrl(), GetSafeHwnd());
		}
		break;
	case WM_DESTROY:
	{
		DeleteAllColumns(true);
		break;
	}
	case WM_MOUSEMOVE:
	{
		LVHITTESTINFO lvHitTestInfo = { 0 };
		lvHitTestInfo.pt.x = GET_X_LPARAM(lParam);
		lvHitTestInfo.pt.y = GET_Y_LPARAM(lParam);

		WC_B(::SendMessage(m_hWnd, LVM_HITTEST, 0, reinterpret_cast<LPARAM>(&lvHitTestInfo)));
		// Hover highlight
		if (lvHitTestInfo.flags & LVHT_ONITEM)
		{
			if (iItemCur != lvHitTestInfo.iItem)
			{
				// 'Unglow' the previous line
				if (-1 != iItemCur)
				{
					m_iItemCurHover = -1;
					DrawListItemGlow(m_hWnd, iItemCur);
				}

				// Glow the current line - it's important that m_iItemCurHover be set before we draw the glow
				m_iItemCurHover = lvHitTestInfo.iItem;
				DrawListItemGlow(m_hWnd, lvHitTestInfo.iItem);

				TRACKMOUSEEVENT tmEvent = { 0 };
				tmEvent.cbSize = sizeof(TRACKMOUSEEVENT);
				tmEvent.dwFlags = TME_LEAVE;
				tmEvent.hwndTrack = m_hWnd;

				EC_B(TrackMouseEvent(&tmEvent));
			}
		}
		else
		{
			if (-1 != iItemCur)
			{
				m_iItemCurHover = -1;
				DrawListItemGlow(m_hWnd, iItemCur);
			}
		}
		break;
	}
	case WM_MOUSELEAVE:
		// Turn off any hot highlighting
		if (-1 != iItemCur)
		{
			m_iItemCurHover = -1;
			DrawListItemGlow(m_hWnd, iItemCur);
		}
		break;
	}
	return CListCtrl::WindowProc(message, wParam, lParam);
}

// Override for list item painting
void CSortListCtrl::OnCustomDraw(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
{
	CustomDrawList(pNMHDR, pResult, m_iItemCurHover);
}

void CSortListCtrl::OnDeleteAllItems(_In_ NMHDR* /*pNMHDR*/, _In_ LRESULT* pResult)
{
	*pResult = false; // make sure we get LVN_DELETEITEM for all items
}

void CSortListCtrl::OnDeleteItem(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
{
	auto pNMV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);

	if (pNMV)
	{
		auto lpData = reinterpret_cast<SortListData*>(GetItemData(pNMV->iItem));
		delete lpData;
	}
	*pResult = 0;
}

_Check_return_ SortListData* CSortListCtrl::InsertRow(int iRow, const wstring& szText) const
{
	return InsertRow(iRow, szText, 0, 0);
}

_Check_return_ SortListData* CSortListCtrl::InsertRow(int iRow, const wstring& szText, int iIndent, int iImage) const
{
	auto lpData = new SortListData();

	LVITEMW lvItem = { 0 };
	lvItem.iItem = iRow;
	lvItem.iSubItem = 0;
	lvItem.mask = LVIF_TEXT | LVIF_PARAM | LVIF_INDENT | LVIF_IMAGE;
	lvItem.pszText = const_cast<LPWSTR>(szText.c_str());
	lvItem.iIndent = iIndent;
	lvItem.iImage = iImage;
	lvItem.lParam = reinterpret_cast<LPARAM>(lpData);
	::SendMessage(m_hWnd, LVM_INSERTITEMW, 0, reinterpret_cast<LPARAM>(&lvItem));

	return lpData;
}

void CSortListCtrl::MySetRedraw(bool bRedraw)
{
	if (bRedraw)
	{
		m_iRedrawCount--;
		if (0 >= m_iRedrawCount)
		{
			SetRedraw(true);
		}
	}
	else
	{
		if (0 == m_iRedrawCount)
		{
			SetRedraw(false);
		}
		m_iRedrawCount++;
	}
}

enum __SortStyle
{
	SORTSTYLE_STRING = 0,
	SORTSTYLE_NUMERIC,
	SORTSTYLE_HEX,
};

struct SortInfo
{
	bool bSortUp;
	__SortStyle sortstyle;
};

#define sortEqual 0
#define sort1First -1
#define sort2First 1
// Sort the item in alphabetical order.
// Simplistic algorithm that only looks at the text. This pays no attention to the underlying MAPI properties.
// This will sort dates and numbers badly. :)
_Check_return_ int CALLBACK CSortListCtrl::MyCompareProc(_In_ LPARAM lParam1, _In_ LPARAM lParam2, _In_ LPARAM lParamSort)
{
	if (!lParamSort) return sortEqual;
	auto iRet = 0;
	auto lpSortInfo = reinterpret_cast<SortInfo*>(lParamSort);
	auto lpData1 = reinterpret_cast<SortListData*>(lParam1);
	auto lpData2 = reinterpret_cast<SortListData*>(lParam2);

	if (!lpData1 && !lpData2) return sortEqual; // item which don't exist must be equal
	if (!lpData1) return sort2First; // sort null items to the end - this makes lParam2>lParam1
	if (!lpData2) return sort1First; // sort null items to the end - this makes lParam1>lParam2

	// Don't sort items which aren't fully loaded
	if (!lpData1->bItemFullyLoaded) return sort2First;
	if (!lpData2->bItemFullyLoaded) return sort1First;

	switch (lpSortInfo->sortstyle)
	{
	case SORTSTYLE_STRING:
		// Empty strings should always sort after non-empty strings
		if (lpData1->szSortText.empty()) return sort2First;
		if (lpData2->szSortText.empty()) return sort1First;
		iRet = lpData1->szSortText.compare(lpData2->szSortText);

		return lpSortInfo->bSortUp ? -iRet : iRet;
	case SORTSTYLE_HEX:
		// Empty strings should always sort after non-empty strings
		if (lpData1->szSortText.empty()) return sort2First;
		if (lpData2->szSortText.empty()) return sort1First;

		if (lpData1->ulSortValue.LowPart == lpData2->ulSortValue.LowPart)
		{
			iRet = lpData1->szSortText.compare(lpData2->szSortText);
		}
		else
		{
			int lCheck = max(lpData1->ulSortValue.LowPart, lpData2->ulSortValue.LowPart);

			for (auto i = 0; i < lCheck; i++)
			{
				if (lpData1->szSortText[i] != lpData2->szSortText[i])
				{
					iRet = lpData1->szSortText[i] < lpData2->szSortText[i] ? -1 : 1;
					break;
				}
			}
		}

		return lpSortInfo->bSortUp ? -iRet : iRet;
	case SORTSTYLE_NUMERIC:
	{
		auto ul1 = lpData1->ulSortValue;
		auto ul2 = lpData2->ulSortValue;
		return lpSortInfo->bSortUp ? ul2.QuadPart > ul1.QuadPart:ul1.QuadPart >= ul2.QuadPart;
	}
	default:
		break;
	}

	return 0;
}

#ifndef HDF_SORTUP
#define HDF_SORTUP 0x0400
#endif
#ifndef HDF_SORTDOWN
#define HDF_SORTDOWN 0x0200
#endif

void CSortListCtrl::SortClickedColumn()
{
	auto hRes = S_OK;
	HDITEM hdItem = { 0 };
	ULONG ulPropTag = NULL;
	SortInfo sortinfo = { 0 };

	// szText will be filled out by our LVM_GETITEMW calls
	// There's little point in getting more than 128 characters for sorting
	WCHAR szText[128];

	m_bHaveSorted = true;
	auto lpMyHeader = GetHeaderCtrl();
	if (lpMyHeader)
	{
		// Clear previous sorts
		for (auto i = 0; i < lpMyHeader->GetItemCount(); i++)
		{
			hdItem.mask = HDI_FORMAT;
			EC_B(lpMyHeader->GetItem(i, &hdItem));
			hdItem.fmt &= ~(HDF_SORTUP | HDF_SORTDOWN);
			lpMyHeader->SetItem(i, &hdItem);
		}

		hdItem.mask = HDI_FORMAT;
		lpMyHeader->GetItem(m_iClickedColumn, &hdItem);
		hdItem.fmt &= ~(HDF_SORTUP | HDF_SORTDOWN);
		hdItem.fmt |= m_bSortUp ? HDF_SORTUP : HDF_SORTDOWN;
		lpMyHeader->SetItem(m_iClickedColumn, &hdItem);

		hdItem.mask = HDI_LPARAM;
		EC_B(lpMyHeader->GetItem(m_iClickedColumn, &hdItem));
		if (hdItem.lParam)
		{
			ulPropTag = reinterpret_cast<LPHEADERDATA>(hdItem.lParam)->ulPropTag;
		}
	}

	// Set our sort text
	LVITEMW lvi = { 0 };
	lvi.mask = LVIF_PARAM | LVIF_TEXT;
	lvi.iSubItem = m_iClickedColumn;
	lvi.cchTextMax = _countof(szText);
	lvi.pszText = szText;

	switch (PROP_TYPE(ulPropTag))
	{
	case PT_I2:
	case PT_LONG:
	case PT_R4:
	case PT_DOUBLE:
	case PT_APPTIME:
	case PT_CURRENCY:
	case PT_I8:
		sortinfo.sortstyle = SORTSTYLE_NUMERIC;
		for (auto i = 0; i < GetItemCount(); i++)
		{
			lvi.iItem = i;
			lvi.lParam = 0;
			szText[0] = NULL;
			::SendMessage(m_hWnd, LVM_GETITEMW, static_cast<WPARAM>(0), reinterpret_cast<LPARAM>(&lvi));
			auto lpData = reinterpret_cast<SortListData*>(lvi.lParam);
			if (lpData)
			{
				lpData->szSortText.clear();
				lpData->ulSortValue.QuadPart = wstringToUlong(szText, 10, false);
			}
		}
		break;
	case PT_SYSTIME:
	{
		ULONG ulSourceCol = 0;
		if (hdItem.lParam)
		{
			ulSourceCol = reinterpret_cast<LPHEADERDATA>(hdItem.lParam)->ulTagArrayRow;
		}

		sortinfo.sortstyle = SORTSTYLE_NUMERIC;
		for (auto i = 0; i < GetItemCount(); i++)
		{
			auto lpData = reinterpret_cast<SortListData*>(GetItemData(i));
			if (lpData)
			{
				lpData->szSortText.clear();
				lpData->ulSortValue.QuadPart = 0;
				if (ulSourceCol < lpData->cSourceProps && PROP_TYPE(lpData->lpSourceProps[ulSourceCol].ulPropTag) == PT_SYSTIME)
				{
					lpData->ulSortValue.LowPart = lpData->lpSourceProps[ulSourceCol].Value.ft.dwLowDateTime;
					lpData->ulSortValue.HighPart = lpData->lpSourceProps[ulSourceCol].Value.ft.dwHighDateTime;
				}
			}
		}
		break;
	}
	case PT_BINARY:
		sortinfo.sortstyle = SORTSTYLE_HEX;
		for (auto i = 0; i < GetItemCount(); i++)
		{
			lvi.iItem = i;
			lvi.lParam = 0;
			szText[0] = NULL;
			::SendMessage(m_hWnd, LVM_GETITEMW, static_cast<WPARAM>(0), reinterpret_cast<LPARAM>(&lvi));
			auto lpData = reinterpret_cast<SortListData*>(lvi.lParam);
			if (lpData)
			{
				// Remove the lpb prefix
				lpData->szSortText = wstring(szText);
				auto pos = lpData->szSortText.find(L"lpb: "); // STRING_OK
				if (pos != string::npos)
				{
					lpData->szSortText = lpData->szSortText.substr(pos);
				}

				lpData->ulSortValue.LowPart = static_cast<DWORD>(lpData->szSortText.length());
			}
		}
		break;
	default:
		sortinfo.sortstyle = SORTSTYLE_STRING;
		for (auto i = 0; i < GetItemCount(); i++)
		{
			lvi.iItem = i;
			lvi.lParam = 0;
			szText[0] = NULL;
			::SendMessage(m_hWnd, LVM_GETITEMW, static_cast<WPARAM>(0), reinterpret_cast<LPARAM>(&lvi));
			auto lpData = reinterpret_cast<SortListData*>(lvi.lParam);
			if (lpData)
			{
				lpData->szSortText = szText;
				lpData->ulSortValue.QuadPart = NULL;
			}
		}
		break;
	}

	sortinfo.bSortUp = m_bSortUp;
	EC_B(SortItems(MyCompareProc, reinterpret_cast<LPARAM>(&sortinfo)));
}

// Leverage in support for sorting columns.
// Basically a call to CListCtrl::SortItems with MyCompareProc
void CSortListCtrl::OnColumnClick(int iColumn)
{
	if (m_bHaveSorted && m_iClickedColumn == iColumn)
	{
		m_bSortUp = !m_bSortUp;
	}
	else
	{
		if (!m_bHaveSorted) m_bSortUp = false; // init this to down arrow on first pass
		m_iClickedColumn = iColumn;
	}
	SortClickedColumn();
}

// Used by child classes to force a sort order on a column
void CSortListCtrl::FakeClickColumn(int iColumn, bool bSortUp)
{
	m_iClickedColumn = iColumn;
	m_bSortUp = bSortUp;
	SortClickedColumn();
}

void CSortListCtrl::AutoSizeColumn(int iColumn, int iMaxWidth, int iMinWidth)
{
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	MySetRedraw(false);
	SetColumnWidth(iColumn, LVSCW_AUTOSIZE_USEHEADER);
	auto width = GetColumnWidth(iColumn);
	if (iMaxWidth && width > iMaxWidth) SetColumnWidth(iColumn, iMaxWidth);
	else if (width < iMinWidth) SetColumnWidth(iColumn, iMinWidth);
	MySetRedraw(true);
}

void CSortListCtrl::AutoSizeColumns(bool bMinWidth)
{
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric, CLASS, L"AutoSizeColumns", L"Sizing columns\n");
	auto lpMyHeader = GetHeaderCtrl();
	if (lpMyHeader)
	{
		MySetRedraw(false);
		for (auto i = 0; i < lpMyHeader->GetItemCount(); i++)
		{
			AutoSizeColumn(i, 150, bMinWidth ? 150 : 0);
		}

		MySetRedraw(true);
	}
}

void CSortListCtrl::DeleteAllColumns(bool bShutdown)
{
	auto hRes = S_OK;
	HDITEM hdItem = { 0 };

	DebugPrintEx(DBGGeneric, CLASS, L"DeleteAllColumns", L"Deleting existing columns\n");
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	auto lpMyHeader = GetHeaderCtrl();
	if (lpMyHeader)
	{
		// Delete all of the old column headers
		auto iColCount = lpMyHeader->GetItemCount();
		// Delete from right to left, which is faster than left to right
		if (iColCount)
		{
			if (!bShutdown) MySetRedraw(false);
			for (auto iCol = iColCount - 1; iCol >= 0; iCol--)
			{
				hdItem.mask = HDI_LPARAM;
				EC_B(lpMyHeader->GetItem(iCol, &hdItem));

				// This will be a HeaderData, created in CContentsTableListCtrl::AddColumn
				if (SUCCEEDED(hRes))
					delete reinterpret_cast<HeaderData*>(hdItem.lParam);

				if (!bShutdown) EC_B(DeleteColumn(iCol));
			}

			if (!bShutdown) MySetRedraw(true);
		}
	}
}

void CSortListCtrl::AllowEscapeClose()
{
	m_bAllowEscapeClose = true;
}

// Assert that we want all keyboard input (including ENTER!)
// In the case of TAB though, let it through
_Check_return_ UINT CSortListCtrl::OnGetDlgCode()
{
	auto iDlgCode = CListCtrl::OnGetDlgCode();

	iDlgCode |= DLGC_WANTMESSAGE;

	// to make sure that the control key is not pressed
	if (GetKeyState(VK_CONTROL) >= 0 && m_hWnd == ::GetFocus())
	{
		// to make sure that the Tab key is pressed
		if (GetKeyState(VK_TAB) < 0)
			iDlgCode &= ~(DLGC_WANTALLKEYS | DLGC_WANTMESSAGE | DLGC_WANTTAB);

		if (m_bAllowEscapeClose)
		{
			// to make sure that the Escape key is pressed
			if (GetKeyState(VK_ESCAPE) < 0)
				iDlgCode &= ~(DLGC_WANTALLKEYS | DLGC_WANTMESSAGE | DLGC_WANTTAB);
		}
	}

	return iDlgCode;
}

void CSortListCtrl::SetItemText(int nItem, int nSubItem, const wstring& lpszText)
{
	// Remove any whitespace before setting in the list
	auto szWhitespace = const_cast<LPWSTR>(wcspbrk(lpszText.c_str(), L"\r\n\t")); // STRING_OK
	while (szWhitespace != nullptr)
	{
		szWhitespace[0] = L' ';
		szWhitespace = static_cast<LPWSTR>(wcspbrk(szWhitespace, L"\r\n\t")); // STRING_OK
	}

	(void)CListCtrl::SetItemText(nItem, nSubItem, wstringTotstring(lpszText).c_str());
}

wstring CSortListCtrl::GetItemText(_In_ int nItem, _In_ int nSubItem) const
{
	return LPCTSTRToWstring(CListCtrl::GetItemText(nItem, nSubItem));
}

// if asked to select the item after the last item - will select the last item.
void CSortListCtrl::SetSelectedItem(int iItem)
{
	auto hRes = S_OK;
	DebugPrintEx(DBGGeneric, CLASS, L"SetSelectedItem", L"selecting iItem = %d\n", iItem);
	auto bSet = SetItemState(iItem, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);

	if (bSet)
	{
		EnsureVisible(iItem, false);
	}
	else if (iItem > 0)
	{
		EC_B(SetItemState(iItem - 1, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED));
		EnsureVisible(iItem - 1, false);
	}
}