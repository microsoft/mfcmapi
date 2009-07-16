// SortListCtrl.cpp : implementation file
//

#include "stdafx.h"
#include "SortListCtrl.h"
#include "MapiObjects.h"
#include "ContentsTableDlg.h"
#include "MAPIFunctions.h"
#include "MFCUtilityFunctions.h"
#include "InterpretProp.h"
#include "AboutDlg.h"
#include "SortHeader.h"
#include "AdviseSink.h"

static TCHAR* CLASS = _T("CSortListCtrl");

void FreeSortListData(SortListData* lpData)
{
	if (!lpData) return;
	switch (lpData->ulSortDataType)
	{
	case SORTLIST_CONTENTS: // _ContentsData
	case SORTLIST_PROP: // _PropListData
	case SORTLIST_MVPROP: // _MVPropData
	case SORTLIST_TAGARRAY: // _TagData
	case SORTLIST_BINARY: // _BinaryData
	case SORTLIST_RES: // _ResData
	case SORTLIST_COMMENT: // _CommentData
		// Nothing to do
		break;
	case SORTLIST_TREENODE: // _NodeData
		if (lpData->data.Node.lpAdviseSink)
		{
			// unadvise before releasing our sink
			if (lpData->data.Node.ulAdviseConnection && lpData->data.Node.lpHierarchyTable)
				lpData->data.Node.lpHierarchyTable->Unadvise(lpData->data.Node.ulAdviseConnection);
			lpData->data.Node.lpAdviseSink->Release();
		}
		if (lpData->data.Node.lpHierarchyTable) lpData->data.Node.lpHierarchyTable->Release();
		break;
	}
	MAPIFreeBuffer(lpData->szSortText);
	MAPIFreeBuffer(lpData->lpSourceProps);
	MAPIFreeBuffer(lpData);
}

/////////////////////////////////////////////////////////////////////////////
// CSortListCtrl

CSortListCtrl::CSortListCtrl()
{
	TRACE_CONSTRUCTOR(CLASS);
	m_cRef = 1;
	m_iRedrawCount = 0;
	m_iClickedColumn = 0;
	m_bSortUp = false;
	m_bHaveSorted = FALSE;
	m_bHeaderSubclassed = false;
}

CSortListCtrl::~CSortListCtrl()
{
	TRACE_DESTRUCTOR(CLASS);
	DestroyWindow();
}

STDMETHODIMP_(ULONG) CSortListCtrl::AddRef()
{
	LONG lCount(InterlockedIncrement(&m_cRef));
	TRACE_ADDREF(CLASS,lCount);
	return lCount;
}

STDMETHODIMP_(ULONG) CSortListCtrl::Release()
{
	LONG lCount(InterlockedDecrement(&m_cRef));
	TRACE_RELEASE(CLASS,lCount);
	if (!lCount)  delete this;
	return lCount;
}

BEGIN_MESSAGE_MAP(CSortListCtrl, CListCtrl)
	ON_WM_KEYDOWN()
	ON_WM_GETDLGCODE()
	ON_NOTIFY_REFLECT(LVN_DELETEALLITEMS, OnDeleteAllItems)
	ON_NOTIFY_REFLECT(LVN_DELETEITEM, OnDeleteItem)
END_MESSAGE_MAP()


HRESULT CSortListCtrl::Create(CWnd* pCreateParent, ULONG ulFlags, UINT nID, BOOL bImages)
{
	HRESULT hRes = S_OK;
	EC_B(CListCtrl::Create(
		ulFlags
		| LVS_REPORT
		| LVS_SHOWSELALWAYS
//		| LVS_NOSORTHEADER // Put this back in to kill the 'clickable' headers
		| WS_TABSTOP
		| WS_CHILD
//		| WS_BORDER
//		| WS_CLIPCHILDREN // if this is passed, the header control doesn't get updated properly
		| WS_CLIPSIBLINGS
		| WS_VISIBLE,
		CRect(0,0,0,0), // size doesn't matter
		pCreateParent,
		nID));

	SetExtendedStyle(GetExtendedStyle() | LVS_EX_FULLROWSELECT | LVS_EX_HEADERDRAGDROP | LVS_EX_INFOTIP);

	if (bImages)
	{
		HIMAGELIST hImageList=ImageList_Create(16, 16, ILC_COLOR32 | ILC_MASK, 1, 1);
		m_ImageList.Attach(hImageList);

		CBitmap myBitmap;
		myBitmap.LoadBitmap(IDB_ICONS);
		m_ImageList.Add(&myBitmap, RGB(255, 0, 255));

		SetImageList(&m_ImageList,LVSIL_SMALL);
	}

	return hRes;
}

LRESULT CSortListCtrl::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	LRESULT lResult = NULL;

	switch (message)
	{
	// I can handle notify messages for my child header control since I am the parent window
	// This makes it easy for me to customize the child control to do what I want
	// I cannot handle notify's heading to my parent though - have to depend on reflection for that
	case WM_NOTIFY:
		{
			LPNMHDR pHdr = (LPNMHDR) lParam;

			switch(pHdr->code)
			{
			case HDN_ITEMCLICKA:
			case HDN_ITEMCLICKW:
				OnColumnClick(((LPNMHEADERW)pHdr)->iItem);
				return lResult;
				break;
			case HDN_DIVIDERDBLCLICKA:
			case HDN_DIVIDERDBLCLICKW:
				AutoSizeColumn(((LPNMHEADERW)pHdr)->iItem,0,0);
				return lResult;
				break;
			}
			break; // WM_NOTIFY
		}
	case WM_PAINT: // TODO: Find a MUCH better message to host this on
		// subclass the header
		if (!m_bHeaderSubclassed)
		{
			m_bHeaderSubclassed = m_cSortHeader.Init(GetHeaderCtrl(),GetSafeHwnd());
		}
		break;
	case WM_ERASEBKGND:
		{
			CListCtrl::OnEraseBkgnd((CDC*) wParam);
			return TRUE;
			break;
		}
	case WM_DESTROY:
		{
			DeleteAllColumns();
			break;
		}
	} // end switch
	return CListCtrl::WindowProc(message,wParam,lParam);
}

void CSortListCtrl::OnDeleteAllItems(NMHDR* /*pNMHDR*/, LRESULT* pResult)
{
	*pResult = FALSE; // make sure we get LVN_DELETEITEM for all items
}

void CSortListCtrl::OnDeleteItem(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMLISTVIEW pNMV = (LPNMLISTVIEW) pNMHDR;

	if (pNMV)
	{
		SortListData* lpData;
		lpData = (SortListData*) GetItemData(pNMV->iItem);
		FreeSortListData(lpData);
	}
	*pResult = 0;
}

SortListData* CSortListCtrl::InsertRow(int iRow, LPTSTR szText)
{
	return InsertRow(iRow,szText,0,0);
}

SortListData* CSortListCtrl::InsertRow(int iRow, LPTSTR szText, int iIndent, int iImage)
{
	HRESULT			hRes = S_OK;
	SortListData*	lpData = NULL;

	EC_H(MAPIAllocateBuffer(
		(ULONG)sizeof(SortListData),
		(LPVOID*) &lpData));
	if (lpData)
	{
		memset(lpData, 0, sizeof(SortListData));
	}

	LVITEM lvItem = {0};
	lvItem.iItem = iRow;
	lvItem.iSubItem = 0;
	lvItem.mask = LVIF_TEXT | LVIF_PARAM | LVIF_INDENT | LVIF_IMAGE;
	lvItem.pszText = szText;
	lvItem.iIndent = iIndent;
	lvItem.iImage = iImage;
	lvItem.lParam = (LPARAM) lpData;
	iRow = InsertItem(&lvItem); // Assign result to iRow in case it changes

	return lpData;
}

void CSortListCtrl::MySetRedraw(BOOL bRedraw)
{
	if (bRedraw)
	{
		m_iRedrawCount--;
		if (0 >= m_iRedrawCount)
		{
			SetRedraw(TRUE);
		}
	}
	else
	{
		if (0 == m_iRedrawCount)
		{
			SetRedraw(FALSE);
		}
		m_iRedrawCount++;
	}
} // CSortListCtrl::MySetRedraw


enum __SortStyle
{
	SORTSTYLE_STRING = 0,
	SORTSTYLE_NUMERIC
};

typedef struct _SortInfo
{
	BOOL		bSortUp;
	__SortStyle	sortstyle;
} SortInfo;


#define sortEqual 0
#define sort1First -1
#define sort2First 1
// Sort the item in alphabetical order.
// Simplistic algorithm that only looks at the text. This pays no attention to the underlying MAPI properties.
// This will sort dates and numbers badly. :)
int CALLBACK CSortListCtrl::MyCompareProc(LPARAM lParam1, LPARAM lParam2, LPARAM lParamSort)
{
	if (!lParamSort) return sortEqual;
	SortInfo* lpSortInfo = (SortInfo*) lParamSort;
	SortListData* lpData1 = (SortListData*) lParam1;
	SortListData* lpData2 = (SortListData*) lParam2;

	if (!lpData1 && !lpData2) return sortEqual; // item which don't exist must be equal
	if (!lpData1) return sort2First; // sort null items to the end - this makes lParam2>lParam1
	if (!lpData2) return sort1First; // sort null items to the end - this makes lParam1>lParam2

	// Don't sort items which aren't fully loaded
	if (!lpData1->bItemFullyLoaded) return sort2First;
	if (!lpData2->bItemFullyLoaded) return sort1First;

	switch (lpSortInfo->sortstyle)
	{
	case SORTSTYLE_STRING:
		{
			TCHAR* sz1 = lpData1->szSortText;
			TCHAR* sz2 = lpData2->szSortText;
			// Empty strings should always sort after non-empty strings
			if (!sz1) return sort2First;
			if (!sz1[0]) return sort2First;
			if (!sz2) return sort1First;
			if (!sz2[0]) return sort1First;
			int iRet = lstrcmpi(
				sz1,
				sz2);

			return (lpSortInfo->bSortUp?-iRet:iRet);
		}
		break;
	case SORTSTYLE_NUMERIC:
		{
			ULARGE_INTEGER ul1 = lpData1->ulSortValue;
			ULARGE_INTEGER ul2 = lpData2->ulSortValue;
			return (lpSortInfo->bSortUp?ul2.QuadPart > ul1.QuadPart:ul1.QuadPart >= ul2.QuadPart);
		}
		break;
	default:
		break;
	}
	return 0;
}

#ifndef HDF_SORTUP
#define HDF_SORTUP              0x0400
#endif
#ifndef HDF_SORTDOWN
#define HDF_SORTDOWN            0x0200
#endif

void CSortListCtrl::SortClickedColumn()
{
	HRESULT			hRes = S_OK;
	HDITEM			hdItem = {0};
	ULONG			ulPropTag = NULL;
	CHeaderCtrl*	lpMyHeader = NULL;
	SortInfo		sortinfo = {0};

	// There's little point in getting more than 128 characters for sorting
	TCHAR			szText[128];

	m_bHaveSorted = true;
	lpMyHeader = GetHeaderCtrl();
	if (lpMyHeader)
	{
		static BOOL bIsXP = false;
		static BOOL bVersionCheck = false;

		if (!bVersionCheck)
		{
			// check windows XP
			OSVERSIONINFOEX info;
			ZeroMemory(&info, sizeof(OSVERSIONINFOEX));
			info.dwOSVersionInfoSize = sizeof(OSVERSIONINFOEX);
			if (GetVersionEx((OSVERSIONINFO*)&info))
				bIsXP = (info.dwMajorVersion > 5) |
				(info.dwMajorVersion == 5 && info.dwMinorVersion >= 1);
			bVersionCheck = true;
		}

		// Just in case we got any old bitmaps to clear
		for (int i = 0 ; i < lpMyHeader->GetItemCount(); i++)
		{
			hdItem.mask = HDI_FORMAT | HDI_BITMAP;
			EC_B(lpMyHeader->GetItem(i,&hdItem));

			hdItem.fmt &= ~(HDF_BITMAP | HDF_BITMAP_ON_RIGHT);
			if (bIsXP)
				hdItem.fmt &= ~(HDF_SORTUP | HDF_SORTDOWN);
			if (hdItem.hbm != 0)
			{
				DeleteObject(hdItem.hbm);
				hdItem.hbm = 0;
			}
			lpMyHeader->SetItem(i, &hdItem);
		}

		hdItem.mask = HDI_FORMAT | HDI_BITMAP;
		lpMyHeader->GetItem(m_iClickedColumn, &hdItem);
		if (hdItem.hbm != 0)
		{
			DeleteObject(hdItem.hbm);
			hdItem.hbm = 0;
		}
		hdItem.fmt |= HDF_BITMAP | HDF_BITMAP_ON_RIGHT ;
		if (!bIsXP)
		{
			EC_D(hdItem.hbm, (HBITMAP) LoadImage(AfxGetInstanceHandle(), MAKEINTRESOURCE(m_bSortUp ? IDB_UP : IDB_DOWN), IMAGE_BITMAP, 0, 0, LR_LOADMAP3DCOLORS));
		}
		else
		{
			hdItem.fmt &= ~(HDF_SORTUP | HDF_SORTDOWN);
			hdItem.fmt |= m_bSortUp ? HDF_SORTUP : HDF_SORTDOWN;
		}
		lpMyHeader->SetItem(m_iClickedColumn, &hdItem);

		hdItem.mask = HDI_LPARAM;
		EC_B(lpMyHeader->GetItem(m_iClickedColumn,&hdItem));
		if (hdItem.lParam)
		{
			ulPropTag = ((LPHEADERDATA)hdItem.lParam)->ulPropTag;
		}
	}

// #define PT_UNSPECIFIED	((ULONG)  0)	/* (Reserved for interface use) type doesn't matter to caller */
// #define PT_NULL			((ULONG)  1)	/* NULL property value */
// #define	PT_I2			((ULONG)  2)	/* Signed 16-bit value */
// #define PT_LONG			((ULONG)  3)	/* Signed 32-bit value */
// #define	PT_R4			((ULONG)  4)	/* 4-byte floating point */
// #define PT_DOUBLE		((ULONG)  5)	/* Floating point double */
// #define PT_CURRENCY		((ULONG)  6)	/* Signed 64-bit int (decimal w/	4 digits right of decimal pt) */
// #define	PT_APPTIME		((ULONG)  7)	/* Application time */
// #define PT_ERROR		((ULONG) 10)	/* 32-bit error value */
// #define PT_BOOLEAN		((ULONG) 11)	/* 16-bit boolean (non-zero true) */
// #define PT_OBJECT		((ULONG) 13)	/* Embedded object in a property */
// #define	PT_I8			((ULONG) 20)	/* 8-byte signed integer */
// #define PT_STRING8		((ULONG) 30)	/* Null terminated 8-bit character string */
// #define PT_UNICODE		((ULONG) 31)	/* Null terminated Unicode string */
// #define PT_SYSTIME		((ULONG) 64)	/* FILETIME 64-bit int w/ number of 100ns periods since Jan 1,1601 */
// #define	PT_CLSID		((ULONG) 72)	/* OLE GUID */
// #define PT_BINARY		((ULONG) 258)	/* Uninterpreted (counted byte array) */
	// Set our sort text
	LVITEM lvi = {0};
	lvi.mask = LVIF_PARAM | LVIF_TEXT;
	lvi.iSubItem = m_iClickedColumn;
	lvi.cchTextMax = sizeof(szText);
	lvi.pszText = szText;

	switch(PROP_TYPE(ulPropTag))
	{
	case (PT_I2):
	case (PT_LONG):
	case (PT_R4):
	case (PT_DOUBLE):
	case (PT_APPTIME):
	case (PT_CURRENCY):
	case (PT_I8):
		{
			sortinfo.sortstyle = SORTSTYLE_NUMERIC;
			for (int i = 0;i<GetItemCount();i++)
			{
				lvi.iItem = i;
				lvi.lParam = 0;
				szText[0] = NULL;
				::SendMessage(m_hWnd, LVM_GETITEM, (WPARAM)0, (LPARAM)&lvi);
				SortListData *lpData = (SortListData*) lvi.lParam;
				if (lpData)
				{
					MAPIFreeBuffer(lpData->szSortText);
					lpData->szSortText = NULL;
					lpData->ulSortValue.QuadPart = _tcstoul(szText,NULL,10);
				}
			}
			break;
		}
	case (PT_SYSTIME):
		{
			ULONG ulSourceCol = 0;
			if (hdItem.lParam)
			{
				ulSourceCol = ((LPHEADERDATA)hdItem.lParam)->ulTagArrayRow;
			}

			sortinfo.sortstyle = SORTSTYLE_NUMERIC;
			for (int i = 0;i<GetItemCount();i++)
			{
				SortListData *lpData = (SortListData*) GetItemData(i);
				if (lpData)
				{
					MAPIFreeBuffer(lpData->szSortText);
					lpData->szSortText = NULL;
					lpData->ulSortValue.QuadPart = 0;
					if (ulSourceCol < lpData->cSourceProps && PROP_TYPE(lpData->lpSourceProps[ulSourceCol].ulPropTag) == PT_SYSTIME)
					{
						lpData->ulSortValue.LowPart= lpData->lpSourceProps[ulSourceCol].Value.ft.dwLowDateTime;
						lpData->ulSortValue.HighPart = lpData->lpSourceProps[ulSourceCol].Value.ft.dwHighDateTime;
					}
				}
			}
			break;
		}
	default:
		{
			sortinfo.sortstyle = SORTSTYLE_STRING;
			for (int i = 0;i<GetItemCount();i++)
			{
				lvi.iItem = i;
				lvi.lParam = 0;
				szText[0] = NULL;
				::SendMessage(m_hWnd, LVM_GETITEM, (WPARAM)0, (LPARAM)&lvi);
				SortListData *lpData = (SortListData*) lvi.lParam;
				if (lpData)
				{
					MAPIFreeBuffer(lpData->szSortText);
					EC_H(CopyString(
						&lpData->szSortText,
						szText,
						NULL)); // Do not allocate off of lpData - If we do that we'll 'leak' memory every time we sort until we close the window
					lpData->ulSortValue.QuadPart = NULL;
				}
			}
			break;
		}
	}

	sortinfo.bSortUp = m_bSortUp;
	EC_B(SortItems(MyCompareProc, (LPARAM) &sortinfo));
} // CSortListCtrl::SortClickedColumn

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
} // CSortListCtrl::OnColumnClick

// Used by child classes to force a sort order on a column
void CSortListCtrl::FakeClickColumn(int iColumn, BOOL bSortUp)
{
	m_iClickedColumn = iColumn;
	m_bSortUp = bSortUp;
	SortClickedColumn();
} // CSortListCtrl::FakeClickColumn

void CSortListCtrl::AutoSizeColumn(int iColumn, int iMinWidth, int iMaxWidth)
{
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	MySetRedraw(FALSE);
	SetColumnWidth(iColumn,LVSCW_AUTOSIZE_USEHEADER);
	int width = GetColumnWidth(iColumn);
	if (iMaxWidth && width > iMaxWidth) SetColumnWidth(iColumn,iMaxWidth);
	else if (width < iMinWidth) SetColumnWidth(iColumn,iMinWidth);
	MySetRedraw(TRUE);
}

void CSortListCtrl::AutoSizeColumns()
{
	CHeaderCtrl*	lpMyHeader = NULL;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGGeneric,CLASS,_T("AutoSizeColumns"),_T("Sizing columns\n"));
	lpMyHeader = GetHeaderCtrl();
	if (lpMyHeader)
	{
		MySetRedraw(FALSE);
		for (int i = 0;i<lpMyHeader->GetItemCount();i++)
		{
			AutoSizeColumn(i,100,150);
		}
		MySetRedraw(TRUE);
	}
}

void CSortListCtrl::DeleteAllColumns()
{
	HRESULT			hRes = S_OK;
	CHeaderCtrl*	lpMyHeader = NULL;
	HDITEM			hdItem = {0};

	DebugPrintEx(DBGGeneric,CLASS,_T("DeleteAllColumns"),_T("Deleting existing columns\n"));
	CWaitCursor		Wait; // Change the mouse to an hourglass while we work.

	lpMyHeader = GetHeaderCtrl();
	if (lpMyHeader)
	{
		MySetRedraw(FALSE);

		// Delete all of the old column headers
		int iColCount = lpMyHeader->GetItemCount();
		// Delete from right to left, which is faster than left to right
		int iCol = NULL;
		if (iColCount)
		{
			for (iCol = iColCount-1;iCol >= 0 ; iCol--)
			{
				hdItem.mask = HDI_LPARAM;
				EC_B(lpMyHeader->GetItem(iCol,&hdItem));

				// This will be a HeaderData, created in CContentsTableListCtrl::AddColumn
				if (SUCCEEDED(hRes))
					delete (HeaderData*) hdItem.lParam;

				EC_B(DeleteColumn(iCol));
			}
		}
		MySetRedraw(TRUE);
	}
}

// Assert that we want all keyboard input (including ENTER!)
// In the case of TAB though, let it through
UINT CSortListCtrl::OnGetDlgCode()
{
	UINT iDlgCode = CListCtrl::OnGetDlgCode();

	iDlgCode |= DLGC_WANTMESSAGE;

	// to make sure that the control key is not pressed
	if ((GetKeyState(VK_CONTROL) >= 0)  && (m_hWnd == ::GetFocus()))
	{
		// to make sure that the Tab key is pressed
		if (GetKeyState(VK_TAB) < 0)
			iDlgCode &= ~(DLGC_WANTALLKEYS | DLGC_WANTMESSAGE |	DLGC_WANTTAB);
	}

	return iDlgCode;
} // CSortListCtrl::OnGetDlgCode

BOOL CSortListCtrl::SetItemText(int nItem, int nSubItem, LPCTSTR lpszText)
{
	// Remove any whitespace before setting in the list
	LPTSTR szWhitespace = (LPTSTR) _tcspbrk(lpszText,_T("\r\n\t")); // STRING_OK
	while (szWhitespace != NULL)
	{
		szWhitespace[0] = _T(' ');
		szWhitespace = (LPTSTR) _tcspbrk(szWhitespace,_T("\r\n\t")); // STRING_OK
	}
	return CListCtrl::SetItemText(nItem,nSubItem,lpszText);
} // CSortListCtrl::SetItemText

void CSortListCtrl::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	DebugPrintEx(DBGMenu,CLASS,_T("OnKeyDown"),_T("0x%X\n"),nChar);

	ULONG bCtrlPressed = GetKeyState(VK_CONTROL) < 0;
	ULONG bShiftPressed = GetKeyState(VK_SHIFT) < 0;
	ULONG bMenuPressed = GetKeyState(VK_MENU) < 0;

	if (!bMenuPressed)
	{
		if ('D' == nChar && bCtrlPressed)
		{
			if (bShiftPressed)
			{
				SetDebugLevel(DBGNoDebug);
			}
			else
				SetDebugLevel(DBGAll);
		}
		else if ('H' == nChar && bCtrlPressed)
		{
			DisplayAboutDlg(this);
		}
		else
		{
			CListCtrl::OnKeyDown(nChar,nRepCnt,nFlags);
		}
	}
}

// if asked to select the item after the last item - will select the last item.
void CSortListCtrl::SetSelectedItem(int iItem)
{
	HRESULT hRes = S_OK;
	DebugPrintEx(DBGGeneric,CLASS,_T("SetSelectedItem"),_T("selecting iItem = %d\n"),iItem);
	BOOL bSet = false;

	bSet = SetItemState(iItem,LVIS_SELECTED | LVIS_FOCUSED,LVIS_SELECTED | LVIS_FOCUSED);

	if (bSet)
	{
		EnsureVisible(iItem,false);
	}
	else if (iItem > 0)
	{
		EC_B(SetItemState(iItem-1,LVIS_SELECTED | LVIS_FOCUSED,LVIS_SELECTED | LVIS_FOCUSED));
		EnsureVisible(iItem-1,false);
	}
}