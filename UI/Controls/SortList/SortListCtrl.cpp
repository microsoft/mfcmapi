#include <StdAfx.h>
#include <UI/Controls/SortList/SortListCtrl.h>
#include <UI/Controls/SortList/SortHeader.h>
#include <core/utility/strings.h>
#include <UI/UIFunctions.h>
#include <core/utility/output.h>

namespace controls::sortlistctrl
{
	static std::wstring CLASS = L"CSortListCtrl";

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
		const auto lCount(InterlockedIncrement(&m_cRef));
		TRACE_ADDREF(CLASS, lCount);
		return lCount;
	}

	STDMETHODIMP_(ULONG) CSortListCtrl::Release()
	{
		const auto lCount(InterlockedDecrement(&m_cRef));
		TRACE_RELEASE(CLASS, lCount);
		if (!lCount) delete this;
		return lCount;
	}

	BEGIN_MESSAGE_MAP(CSortListCtrl, CListCtrl)
	ON_WM_KEYDOWN()
	ON_WM_GETDLGCODE()
	ON_WM_DRAWITEM()
#pragma warning(push)
#pragma warning( \
	disable : 26454) // Warning C26454 Arithmetic overflow: 'operator' operation produces a negative unsigned result at compile time
	ON_NOTIFY_REFLECT(LVN_DELETEALLITEMS, OnDeleteAllItems)
	ON_NOTIFY_REFLECT(LVN_DELETEITEM, OnDeleteItem)
	ON_NOTIFY_REFLECT(NM_CUSTOMDRAW, OnCustomDraw)
#pragma warning(pop)
	END_MESSAGE_MAP()

	void CSortListCtrl::Create(_In_ CWnd* pCreateParent, ULONG ulFlags, UINT nID, bool bImages)
	{
		EC_B_S(CListCtrl::Create(
			ulFlags | LVS_REPORT | LVS_SHOWSELALWAYS | WS_TABSTOP | WS_CHILD | WS_CLIPCHILDREN | WS_CLIPSIBLINGS |
				WS_VISIBLE,
			CRect(0, 0, 0, 0), // size doesn't matter
			pCreateParent,
			nID));
		ListView_SetBkColor(m_hWnd, MyGetSysColor(ui::uiColor::Background));
		ListView_SetTextBkColor(m_hWnd, MyGetSysColor(ui::uiColor::Background));
		ListView_SetTextColor(m_hWnd, ui::MyGetSysColor(ui::uiColor::Text));
		::SendMessageA(m_hWnd, WM_SETFONT, reinterpret_cast<WPARAM>(ui::GetSegoeFont()), false);

		SetExtendedStyle(
			GetExtendedStyle() | LVS_EX_FULLROWSELECT | LVS_EX_HEADERDRAGDROP | LVS_EX_INFOTIP | LVS_EX_DOUBLEBUFFER);

		if (bImages)
		{
			// At 100% DPI, ideal ImageList height is 16 pixels
			// At 200%, it's 32, 200% of 16.
			// So we can use our DPI scale to find out what size ImageList to create.
			auto scale = ui::GetDPIScale();
			const auto imageX = 16 * scale.x / scale.denominator;
			const auto imageY = 16 * scale.y / scale.denominator;
			const auto hImageList = ImageList_Create(imageX, imageY, ILC_COLOR32 | ILC_MASK, 1, 1);

			m_ImageList.Attach(hImageList);

			// Now we load our bitmap and scale it the same as our ImageList before loading
			const auto hBitmap = ::LoadBitmap(GetModuleHandle(nullptr), MAKEINTRESOURCE(IDB_ICONS));
			// Our bitmap is 32*32, which is already 2x scaled up
			// So double the denominator to compensate before scaling the bitmap
			scale.denominator *= 2;
			const auto hBitmapScaled = ScaleBitmap(hBitmap, scale);
			m_ImageList.Add(CBitmap::FromHandle(hBitmapScaled), MyGetSysColor(ui::uiColor::BitmapTransBack));

			SetImageList(&m_ImageList, LVSIL_SMALL);
			if (hBitmapScaled) DeleteObject(hBitmapScaled);
			if (hBitmap) DeleteObject(hBitmap);
		}
	}

	static bool s_bInTrack = false;
	static int s_iTrack = 0;
	static int s_iHeaderHeight = 0;

	void OnBeginTrack(_In_ NMHDR* pNMHDR, _In_ HWND hWndParent) noexcept
	{
		RECT rcHeader = {0};
		if (!pNMHDR) return;
		const auto pHdr = reinterpret_cast<LPNMHEADER>(pNMHDR);
		if (Header_GetItemRect(pHdr->hdr.hwndFrom, pHdr->iItem, &rcHeader))
		{
			s_bInTrack = true;
			s_iTrack = rcHeader.right;
			s_iHeaderHeight = rcHeader.bottom - rcHeader.top;
			ui::DrawTrackingBar(pHdr->hdr.hwndFrom, hWndParent, s_iTrack, s_iHeaderHeight, false);
		}
	}

	void OnEndTrack(_In_ NMHDR* pNMHDR, _In_ HWND hWndParent) noexcept
	{
		if (s_bInTrack && pNMHDR)
		{
			ui::DrawTrackingBar(pNMHDR->hwndFrom, hWndParent, s_iTrack, s_iHeaderHeight, true);
		}
		s_bInTrack = false;
	}

	void OnTrack(_In_ NMHDR* pNMHDR, _In_ HWND hWndParent) noexcept
	{
		if (s_bInTrack && pNMHDR)
		{
			RECT rcHeader = {0};
			const auto pHdr = reinterpret_cast<LPNMHEADER>(pNMHDR);
			if (Header_GetItemRect(pHdr->hdr.hwndFrom, pHdr->iItem, &rcHeader) && s_iTrack != rcHeader.right)
			{
				ui::DrawTrackingBar(pHdr->hdr.hwndFrom, hWndParent, s_iTrack, s_iHeaderHeight, true);
				s_iTrack = rcHeader.right;
				ui::DrawTrackingBar(pHdr->hdr.hwndFrom, hWndParent, s_iTrack, s_iHeaderHeight, false);
			}
		}
	}

	LRESULT CSortListCtrl::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
	{
		const auto iItemCur = m_iItemCurHover;

		switch (message)
		{
			// I can handle notify messages for my child header control since I am the parent window
			// This makes it easy for me to customize the child control to do what I want
			// I cannot handle notify heading to my parent though - have to depend on reflection for that
		case WM_NOTIFY:
		{
			const auto pHdr = reinterpret_cast<LPNMHDR>(lParam);

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
				const auto lRet = CListCtrl::WindowProc(message, wParam, lParam);
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
			LVHITTESTINFO lvHitTestInfo = {0};
			lvHitTestInfo.pt.x = GET_X_LPARAM(lParam);
			lvHitTestInfo.pt.y = GET_Y_LPARAM(lParam);

			WC_B_S(::SendMessage(m_hWnd, LVM_HITTEST, 0, reinterpret_cast<LPARAM>(&lvHitTestInfo)));
			// Hover highlight
			if (lvHitTestInfo.flags & LVHT_ONITEM)
			{
				if (iItemCur != lvHitTestInfo.iItem)
				{
					// 'Unglow' the previous line
					if (-1 != iItemCur)
					{
						m_iItemCurHover = -1;
						ui::DrawListItemGlow(m_hWnd, iItemCur);
					}

					// Glow the current line - it's important that m_iItemCurHover be set before we draw the glow
					m_iItemCurHover = lvHitTestInfo.iItem;
					ui::DrawListItemGlow(m_hWnd, lvHitTestInfo.iItem);

					TRACKMOUSEEVENT tmEvent = {0};
					tmEvent.cbSize = sizeof(TRACKMOUSEEVENT);
					tmEvent.dwFlags = TME_LEAVE;
					tmEvent.hwndTrack = m_hWnd;

					EC_B_S(TrackMouseEvent(&tmEvent));
				}
			}
			else
			{
				if (-1 != iItemCur)
				{
					m_iItemCurHover = -1;
					ui::DrawListItemGlow(m_hWnd, iItemCur);
				}
			}
			break;
		}
		case WM_MOUSELEAVE:
			// Turn off any hot highlighting
			if (-1 != iItemCur)
			{
				m_iItemCurHover = -1;
				ui::DrawListItemGlow(m_hWnd, iItemCur);
			}
			break;
		}
		return CListCtrl::WindowProc(message, wParam, lParam);
	}

	// Override for list item painting
	void CSortListCtrl::OnCustomDraw(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult) noexcept
	{
		ui::CustomDrawList(reinterpret_cast<LPNMLVCUSTOMDRAW>(pNMHDR), pResult, m_iItemCurHover);
	}

	void CSortListCtrl::OnDeleteAllItems(_In_ NMHDR* /*pNMHDR*/, _In_ LRESULT* pResult) noexcept
	{
		*pResult = false; // make sure we get LVN_DELETEITEM for all items
	}

	void CSortListCtrl::OnDeleteItem(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
	{
		const auto pNMV = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);

		if (pNMV)
		{
			const auto lpData = reinterpret_cast<sortlistdata::sortListData*>(GetItemData(pNMV->iItem));
			delete lpData;
		}
		*pResult = 0;
	}

	_Check_return_ sortlistdata::sortListData* CSortListCtrl::InsertRow(int iRow, const std::wstring& szText) const
	{
		return InsertRow(iRow, szText, 0, sortIcon::siDefault);
	}

	_Check_return_ sortlistdata::sortListData*
	CSortListCtrl::InsertRow(int iRow, const std::wstring& szText, int iIndent, sortIcon iImage) const
	{
		auto lpData = new (std::nothrow) sortlistdata::sortListData();

		LVITEMW lvItem = {0};
		lvItem.iItem = iRow;
		lvItem.iSubItem = 0;
		lvItem.mask = LVIF_TEXT | LVIF_PARAM | LVIF_INDENT | LVIF_IMAGE;
		lvItem.pszText = const_cast<LPWSTR>(szText.c_str());
		lvItem.iIndent = iIndent;
		lvItem.iImage = static_cast<int>(iImage);
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
#define sort1First (-1)
#define sort2First 1
	// Sort the item in alphabetical order.
	// Simplistic algorithm that only looks at the text. This pays no attention to the underlying MAPI properties.
	// This will sort dates and numbers badly. :)
	_Check_return_ int CALLBACK
	CSortListCtrl::MyCompareProc(_In_ LPARAM lParam1, _In_ LPARAM lParam2, _In_ LPARAM lParamSort) noexcept
	{
		if (!lParamSort) return sortEqual;
		auto iRet = 0;
		const auto lpSortInfo = reinterpret_cast<SortInfo*>(lParamSort);
		auto lpData1 = reinterpret_cast<sortlistdata::sortListData*>(lParam1);
		auto lpData2 = reinterpret_cast<sortlistdata::sortListData*>(lParam2);

		if (!lpData1 && !lpData2) return sortEqual; // item which don't exist must be equal
		if (!lpData1) return sort2First; // sort null items to the end - this makes lParam2>lParam1
		if (!lpData2) return sort1First; // sort null items to the end - this makes lParam1>lParam2

		// Don't sort items which aren't fully loaded
		if (!lpData1->getFullyLoaded()) return sort2First;
		if (!lpData2->getFullyLoaded()) return sort1First;

		switch (lpSortInfo->sortstyle)
		{
		case SORTSTYLE_STRING:
			// Empty strings should always sort after non-empty strings
			if (lpData1->getSortText().empty()) return sort2First;
			if (lpData2->getSortText().empty()) return sort1First;
			iRet = lpData1->getSortText().compare(lpData2->getSortText());

			return lpSortInfo->bSortUp ? -iRet : iRet;
		case SORTSTYLE_HEX:
			// Empty strings should always sort after non-empty strings
			if (lpData1->getSortText().empty()) return sort2First;
			if (lpData2->getSortText().empty()) return sort1First;

			if (lpData1->getSortValue().LowPart == lpData2->getSortValue().LowPart)
			{
				iRet = lpData1->getSortText().compare(lpData2->getSortText());
			}
			else
			{
				const int lCheck = max(lpData1->getSortValue().LowPart, lpData2->getSortValue().LowPart);

				for (auto i = 0; i < lCheck; i++)
				{
					if (lpData1->getSortText()[i] != lpData2->getSortText()[i])
					{
						iRet = lpData1->getSortText()[i] < lpData2->getSortText()[i] ? -1 : 1;
						break;
					}
				}
			}

			return lpSortInfo->bSortUp ? -iRet : iRet;
		case SORTSTYLE_NUMERIC:
		{
			const auto ul1 = lpData1->getSortValue();
			const auto ul2 = lpData2->getSortValue();
			return lpSortInfo->bSortUp ? ul2.QuadPart > ul1.QuadPart : ul1.QuadPart >= ul2.QuadPart;
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
		HDITEM hdItem = {0};
		ULONG ulPropTag = NULL;
		SortInfo sortinfo = {false};

		// szText will be filled out by our LVM_GETITEMW calls
		// There's little point in getting more than 128 characters for sorting
		WCHAR szText[128] = {};

		m_bHaveSorted = true;
		auto lpMyHeader = GetHeaderCtrl();
		if (lpMyHeader)
		{
			// Clear previous sorts
			for (auto i = 0; i < lpMyHeader->GetItemCount(); i++)
			{
				hdItem.mask = HDI_FORMAT;
				EC_B_S(lpMyHeader->GetItem(i, &hdItem));
				hdItem.fmt &= ~(HDF_SORTUP | HDF_SORTDOWN);
				lpMyHeader->SetItem(i, &hdItem);
			}

			hdItem.mask = HDI_FORMAT;
			lpMyHeader->GetItem(m_iClickedColumn, &hdItem);
			hdItem.fmt &= ~(HDF_SORTUP | HDF_SORTDOWN);
			hdItem.fmt |= m_bSortUp ? HDF_SORTUP : HDF_SORTDOWN;
			lpMyHeader->SetItem(m_iClickedColumn, &hdItem);

			hdItem.mask = HDI_LPARAM;
			EC_B_S(lpMyHeader->GetItem(m_iClickedColumn, &hdItem));
			if (hdItem.lParam)
			{
				ulPropTag = reinterpret_cast<LPHEADERDATA>(hdItem.lParam)->ulPropTag;
			}
		}

		// Set our sort text
		LVITEMW lvi = {0};
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
				auto lpData = reinterpret_cast<sortlistdata::sortListData*>(lvi.lParam);
				if (lpData)
				{
					lpData->clearSortValues();
					lpData->setSortValue({strings::wstringToUlong(szText, 10, false), 0});
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
				auto lpData = reinterpret_cast<sortlistdata::sortListData*>(GetItemData(i));
				if (lpData)
				{
					auto row = lpData->getRow();
					lpData->clearSortValues();
					if (ulSourceCol < row.cValues && PROP_TYPE(row.lpProps[ulSourceCol].ulPropTag) == PT_SYSTIME)
					{
						lpData->setSortValue(
							{row.lpProps[ulSourceCol].Value.ft.dwLowDateTime,
							 row.lpProps[ulSourceCol].Value.ft.dwHighDateTime});
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
				auto lpData = reinterpret_cast<sortlistdata::sortListData*>(lvi.lParam);
				if (lpData)
				{
					// Remove the lpb prefix
					auto szSortText = std::wstring(szText);
					const auto pos = szSortText.find(L"lpb: "); // STRING_OK
					if (pos != std::string::npos)
					{
						szSortText = szSortText.substr(pos);
					}

					lpData->setSortText(szSortText);
					lpData->setSortValue({static_cast<DWORD>(szSortText.length()), 0});
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
				auto lpData = reinterpret_cast<sortlistdata::sortListData*>(lvi.lParam);
				if (lpData)
				{
					lpData->clearSortValues();
					lpData->setSortText(szText);
				}
			}
			break;
		}

		sortinfo.bSortUp = m_bSortUp;
		EC_B_S(SortItems(MyCompareProc, reinterpret_cast<LPARAM>(&sortinfo)));
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
		const auto width = GetColumnWidth(iColumn);
		if (iMaxWidth && width > iMaxWidth)
			SetColumnWidth(iColumn, iMaxWidth);
		else if (width < iMinWidth)
			SetColumnWidth(iColumn, iMinWidth);
		MySetRedraw(true);
	}

	void CSortListCtrl::AutoSizeColumns(bool bMinWidth)
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"AutoSizeColumns", L"Sizing columns\n");
		const auto lpMyHeader = GetHeaderCtrl();
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
		HDITEM hdItem = {0};

		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"DeleteAllColumns", L"Deleting existing columns\n");
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		const auto lpMyHeader = GetHeaderCtrl();
		if (lpMyHeader)
		{
			// Delete all of the old column headers
			const auto iColCount = lpMyHeader->GetItemCount();
			// Delete from right to left, which is faster than left to right
			if (iColCount)
			{
				if (!bShutdown) MySetRedraw(false);
				for (auto iCol = iColCount - 1; iCol >= 0; iCol--)
				{
					hdItem.mask = HDI_LPARAM;
					const auto hRes = EC_B(lpMyHeader->GetItem(iCol, &hdItem));

					// This will be a HeaderData, created in CContentsTableListCtrl::AddColumn
					if (SUCCEEDED(hRes)) delete reinterpret_cast<HeaderData*>(hdItem.lParam);

					if (!bShutdown) EC_B_S(DeleteColumn(iCol));
				}

				if (!bShutdown) MySetRedraw(true);
			}
		}
	}

	void CSortListCtrl::AllowEscapeClose() noexcept { m_bAllowEscapeClose = true; }

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
			if (GetKeyState(VK_TAB) < 0) iDlgCode &= ~(DLGC_WANTALLKEYS | DLGC_WANTMESSAGE | DLGC_WANTTAB);

			if (m_bAllowEscapeClose)
			{
				// to make sure that the Escape key is pressed
				if (GetKeyState(VK_ESCAPE) < 0) iDlgCode &= ~(DLGC_WANTALLKEYS | DLGC_WANTMESSAGE | DLGC_WANTTAB);
			}
		}

		return iDlgCode;
	}

	void CSortListCtrl::SetItemText(int nItem, int nSubItem, const std::wstring& lpszText)
	{
		// Remove any whitespace before setting in the list
		auto trimmedText = strings::collapseTree(lpszText);

		auto lvi = LVITEMW();
		lvi.iSubItem = nSubItem;
		lvi.pszText = const_cast<LPWSTR>(trimmedText.c_str());
		static_cast<void>(::SendMessage(m_hWnd, LVM_SETITEMTEXTW, nItem, reinterpret_cast<LPARAM>(&lvi)));
	}

	std::wstring CSortListCtrl::GetItemText(_In_ int nItem, _In_ int nSubItem) const
	{
		return strings::LPCTSTRToWstring(CListCtrl::GetItemText(nItem, nSubItem));
	}

	// if asked to select the item after the last item - will select the last item.
	void CSortListCtrl::SetSelectedItem(int iItem)
	{
		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"SetSelectedItem", L"selecting iItem = %d\n", iItem);
		const auto bSet = SetItemState(iItem, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED);

		if (bSet)
		{
			EnsureVisible(iItem, false);
		}
		else if (iItem > 0)
		{
			EC_B_S(SetItemState(iItem - 1, LVIS_SELECTED | LVIS_FOCUSED, LVIS_SELECTED | LVIS_FOCUSED));
			EnsureVisible(iItem - 1, false);
		}
	}

	int CSortListCtrl::InsertColumnW(_In_ int nCol, const std::wstring& columnHeading) noexcept
	{
		auto column = LVCOLUMNW();
		column.mask = LVCF_TEXT | LVCF_FMT;
		column.pszText = const_cast<LPWSTR>(columnHeading.c_str());
		column.fmt = LVCFMT_LEFT;

		return static_cast<int>(::SendMessage(m_hWnd, LVM_INSERTCOLUMNW, nCol, reinterpret_cast<LPARAM>(&column)));
	}
} // namespace controls::sortlistctrl