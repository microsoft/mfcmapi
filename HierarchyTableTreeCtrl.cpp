#include "stdafx.h"
#include "HierarchyTableTreeCtrl.h"
#include "BaseDialog.h"
#include "HierarchyTableDlg.h"
#include "MapiObjects.h"
#include "MAPIFunctions.h"
#include "UIFunctions.h"
#include "AdviseSink.h"
#include "SingleMAPIPropListCtrl.h"

enum
{
	htPR_DISPLAY_NAME,
	htPR_ENTRYID,
	htPR_INSTANCE_KEY,
	htPR_SUBFOLDERS,
	htPR_CONTAINER_FLAGS,
	htNUMCOLS
};

static const SizedSPropTagArray(htNUMCOLS, sptHTCols) =
{
 htNUMCOLS,
 PR_DISPLAY_NAME,
 PR_ENTRYID,
 PR_INSTANCE_KEY,
 PR_SUBFOLDERS,
 PR_CONTAINER_FLAGS
};

enum
{
	htcPR_CONTENT_COUNT,
	htcPR_ASSOC_CONTENT_COUNT,
	htcPR_DELETED_FOLDER_COUNT,
	htcPR_DELETED_MSG_COUNT,
	htcPR_DELETED_ASSOC_MSG_COUNT,
	htcNUMCOLS
};
static const SizedSPropTagArray(htNUMCOLS, sptHTCountCols) =
{
 htcNUMCOLS,
 PR_CONTENT_COUNT,
 PR_ASSOC_CONTENT_COUNT,
 PR_DELETED_FOLDER_COUNT,
 PR_DELETED_MSG_COUNT,
 PR_DELETED_ASSOC_MSG_COUNT
};

static wstring CLASS = L"CHierarchyTableTreeCtrl";

/////////////////////////////////////////////////////////////////////////////
// CHierarchyTableTreeCtrl

CHierarchyTableTreeCtrl::CHierarchyTableTreeCtrl(
	_In_ CWnd* pCreateParent,
	_In_ CMapiObjects* lpMapiObjects,
	_In_ CHierarchyTableDlg *lpHostDlg,
	ULONG ulDisplayFlags,
	UINT nIDContextMenu)
{
	TRACE_CONSTRUCTOR(CLASS);
	CRect pRect;

	m_cRef = 1;
	m_bShuttingDown = false;

	// We borrow our parent's Mapi objects
	m_lpMapiObjects = lpMapiObjects;
	if (m_lpMapiObjects) m_lpMapiObjects->AddRef();

	m_lpHostDlg = lpHostDlg;
	if (m_lpHostDlg)
	{
		m_lpHostDlg->AddRef();
		m_lpHostDlg->GetClientRect(pRect);
	}

	m_lpContainer = NULL;
	m_ulContainerType = NULL;

	m_ulDisplayFlags = ulDisplayFlags;

	m_bItemSelected = false;

	m_nIDContextMenu = nIDContextMenu;

	m_hItemCurHover = NULL;
	m_HoverButton = false;

	Create(
		TVS_HASBUTTONS
		| TVS_LINESATROOT
		| TVS_EDITLABELS
		| TVS_DISABLEDRAGDROP
		| TVS_SHOWSELALWAYS
		| TVS_FULLROWSELECT
		| WS_CHILD
		| WS_TABSTOP
		//| WS_CLIPCHILDREN
		| WS_CLIPSIBLINGS
		| WS_VISIBLE,
		pRect,
		pCreateParent,
		IDC_FOLDER_TREE);
	TreeView_SetBkColor(m_hWnd, MyGetSysColor(cBackground));
	TreeView_SetTextColor(m_hWnd, MyGetSysColor(cText));
	::SendMessageA(m_hWnd, WM_SETFONT, (WPARAM)GetSegoeFont(), false);
}

CHierarchyTableTreeCtrl::~CHierarchyTableTreeCtrl()
{
	TRACE_DESTRUCTOR(CLASS);
	m_bShuttingDown = true;
	DestroyWindow();

	if (m_lpHostDlg) m_lpHostDlg->Release();
	if (m_lpMapiObjects) m_lpMapiObjects->Release();
}

STDMETHODIMP_(ULONG) CHierarchyTableTreeCtrl::AddRef()
{
	LONG lCount = InterlockedIncrement(&m_cRef);
	TRACE_ADDREF(CLASS, lCount);
	return lCount;
}

STDMETHODIMP_(ULONG) CHierarchyTableTreeCtrl::Release()
{
	LONG lCount = InterlockedDecrement(&m_cRef);
	TRACE_RELEASE(CLASS, lCount);
	if (!lCount) delete this;
	return lCount;
}

BEGIN_MESSAGE_MAP(CHierarchyTableTreeCtrl, CTreeCtrl)
	ON_NOTIFY_REFLECT(TVN_GETDISPINFO, OnGetDispInfo)
	ON_NOTIFY_REFLECT(TVN_SELCHANGED, OnSelChanged)
	ON_NOTIFY_REFLECT(TVN_ENDLABELEDIT, OnEndLabelEdit)
	ON_NOTIFY_REFLECT(TVN_ITEMEXPANDING, OnItemExpanding)
	ON_NOTIFY_REFLECT(TVN_DELETEITEM, OnDeleteItem)
	ON_NOTIFY_REFLECT(NM_DBLCLK, OnDblclk)
	ON_NOTIFY_REFLECT(NM_RCLICK, OnRightClick)
	ON_NOTIFY_REFLECT(NM_CUSTOMDRAW, OnCustomDraw)
	ON_WM_KEYDOWN()
	ON_WM_GETDLGCODE()
	ON_WM_CONTEXTMENU()
	ON_MESSAGE(WM_MFCMAPI_ADDITEM, msgOnAddItem)
	ON_MESSAGE(WM_MFCMAPI_DELETEITEM, msgOnDeleteItem)
	ON_MESSAGE(WM_MFCMAPI_MODIFYITEM, msgOnModifyItem)
	ON_MESSAGE(WM_MFCMAPI_REFRESHTABLE, msgOnRefreshTable)
END_MESSAGE_MAP()

LRESULT CHierarchyTableTreeCtrl::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	// Read the current hover local, since we need to clear it before we do any drawing
	HTREEITEM hItemCurHover = m_hItemCurHover;
	switch (message)
	{
	case WM_MOUSEMOVE:
	{
		HRESULT hRes = S_OK;

		TVHITTESTINFO tvHitTestInfo = { 0 };
		tvHitTestInfo.pt.x = GET_X_LPARAM(lParam);
		tvHitTestInfo.pt.y = GET_Y_LPARAM(lParam);

		WC_B(::SendMessage(m_hWnd, TVM_HITTEST, 0, (LPARAM)&tvHitTestInfo));
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
				DrawTreeItemGlow(m_hWnd, tvHitTestInfo.hItem);

				if (hItemCurHover)
				{
					m_hItemCurHover = NULL;
					DrawTreeItemGlow(m_hWnd, hItemCurHover);
				}
				m_hItemCurHover = tvHitTestInfo.hItem;

				TRACKMOUSEEVENT tmEvent = { 0 };
				tmEvent.cbSize = sizeof(TRACKMOUSEEVENT);
				tmEvent.dwFlags = TME_LEAVE;
				tmEvent.hwndTrack = m_hWnd;

				WC_B(TrackMouseEvent(&tmEvent));
			}
		}
		else
		{
			if (hItemCurHover)
			{
				m_hItemCurHover = NULL;
				DrawTreeItemGlow(m_hWnd, hItemCurHover);
			}
		}
		break;
	}
	case WM_MOUSELEAVE:
	{
		if (hItemCurHover)
		{
			m_hItemCurHover = NULL;
			DrawTreeItemGlow(m_hWnd, hItemCurHover);
		}
		return NULL;
		break;
	}
	} // end switch
	return CTreeCtrl::WindowProc(message, wParam, lParam);
}

/////////////////////////////////////////////////////////////////////////////
// CHierarchyTableTreeCtrl message handlers

_Check_return_ HRESULT CHierarchyTableTreeCtrl::RefreshHierarchyTable()
{
	HRESULT hRes = S_OK;

	// Turn off redraw while we work on the window
	SetRedraw(false);

	m_bItemSelected = false; // clear this just in case

	EC_B(DeleteItem(GetRootItem()));

	if (m_lpContainer)
		EC_H(AddRootNode(m_lpContainer));

	if (m_lpHostDlg) m_lpHostDlg->OnUpdateSingleMAPIPropListCtrl(NULL, NULL);

	// Turn redraw back on to update our view
	SetRedraw(true);
	return hRes;
}

_Check_return_ HRESULT CHierarchyTableTreeCtrl::LoadHierarchyTable(_In_ LPMAPICONTAINER lpMAPIContainer)
{
	HRESULT hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (m_lpContainer) m_lpContainer->Release();
	m_lpContainer = lpMAPIContainer;
	m_ulContainerType = NULL;

	// If we weren't passed a container to load, give up
	if (!m_lpContainer) return hRes;
	m_lpContainer->AddRef();

	m_ulContainerType = GetMAPIObjectType(lpMAPIContainer);

	WC_H(RefreshHierarchyTable());
	if (MAPI_E_NOT_FOUND == hRes)
	{
		WARNHRESMSG(hRes, IDS_HIERARCHNOTFOUND);
	}
	else CHECKHRESMSG(hRes, IDS_REFRESHHIERARCHYFAILED);

	return hRes;
}

// TODO: Move code to copy props here and normalize calling code so it doesn't appear to leak
SortListData* BuildNodeData(
	ULONG cProps,
	_In_opt_ LPSPropValue lpProps,
	_In_opt_ LPSBinary lpEntryID,
	_In_opt_ LPSBinary lpInstanceKey,
	ULONG bSubfolders,
	ULONG ulContainerFlags)
{
	HRESULT hRes = S_OK;
	SortListData* lpData = NULL;
	// We're gonna set up a data item to pass off to the tree
	// Allocate some space
	WC_H(MAPIAllocateBuffer(
		(ULONG)sizeof(SortListData),
		(LPVOID *)&lpData));

	if (lpData)
	{
		memset(lpData, 0, sizeof(SortListData));
		lpData->ulSortDataType = SORTLIST_TREENODE;

		if (lpEntryID)
		{
			WC_H(MAPIAllocateMore(
				(ULONG)sizeof(SBinary),
				lpData,
				(LPVOID*)&lpData->data.Node.lpEntryID));

			// Copy the data over
			WC_H(CopySBinary(
				lpData->data.Node.lpEntryID,
				lpEntryID,
				lpData));
		}

		if (lpInstanceKey)
		{
			WC_H(MAPIAllocateMore(
				(ULONG)sizeof(SBinary),
				lpData,
				(LPVOID*)&lpData->data.Node.lpInstanceKey));
			WC_H(CopySBinary(
				lpData->data.Node.lpInstanceKey,
				lpInstanceKey,
				lpData));
		}

		lpData->data.Node.cSubfolders = -1;
		if (bSubfolders != MAPI_E_NOT_FOUND)
		{
			lpData->data.Node.cSubfolders = (bSubfolders != 0);
		}
		else if (ulContainerFlags != MAPI_E_NOT_FOUND)
		{
			lpData->data.Node.cSubfolders = (ulContainerFlags & AB_SUBCONTAINERS) ? 1 : 0;
		}
		lpData->cSourceProps = cProps;
		lpData->lpSourceProps = lpProps;

		return lpData;
	}
	return NULL;
} // BuildNodeData

SortListData* BuildNodeData(_In_ LPSRow lpsRow)
{
	if (!lpsRow) return NULL;

	LPSPropValue lpEID = NULL; // don't free
	LPSPropValue lpInstance = NULL; // don't free
	LPSBinary lpEIDBin = NULL; // don't free
	LPSBinary lpInstanceBin = NULL; // don't free
	LPSPropValue lpSubfolders = NULL; // don't free
	LPSPropValue lpContainerFlags = NULL; // don't free

	lpEID = PpropFindProp(
		lpsRow->lpProps,
		lpsRow->cValues,
		PR_ENTRYID);
	if (lpEID) lpEIDBin = &lpEID->Value.bin;
	lpInstance = PpropFindProp(
		lpsRow->lpProps,
		lpsRow->cValues,
		PR_INSTANCE_KEY);
	if (lpInstance) lpInstanceBin = &lpInstance->Value.bin;

	lpSubfolders = PpropFindProp(
		lpsRow->lpProps,
		lpsRow->cValues,
		PR_SUBFOLDERS);
	lpContainerFlags = PpropFindProp(
		lpsRow->lpProps,
		lpsRow->cValues,
		PR_CONTAINER_FLAGS);

	return BuildNodeData(
		lpsRow->cValues,
		lpsRow->lpProps, // pass on props to be archived in node
		lpEIDBin,
		lpInstanceBin,
		lpSubfolders ? (ULONG)lpSubfolders->Value.b : MAPI_E_NOT_FOUND,
		lpContainerFlags ? lpContainerFlags->Value.ul : MAPI_E_NOT_FOUND);
} // BuildNodeData

// Removes any existing node data and replaces it with lpData
void SetNodeData(HWND hWnd, HTREEITEM hItem, SortListData* lpData)
{
	if (lpData)
	{
		TVITEM tvItem = { 0 };
		tvItem.hItem = hItem;
		tvItem.mask = TVIF_PARAM;
		if (TreeView_GetItem(hWnd, &tvItem) && tvItem.lParam)
		{
			DebugPrintEx(DBGHierarchy, CLASS, L"SetNodeData", L"Node %p, replacing data\n", hItem);
			FreeSortListData((SortListData*)tvItem.lParam);
		}
		else
		{
			DebugPrintEx(DBGHierarchy, CLASS, L"SetNodeData", L"Node %p, first data\n", hItem);
		}

		tvItem.lParam = (LPARAM)lpData;
		TreeView_SetItem(hWnd, &tvItem);
		// The tree now owns our lpData
	}
} // SetNodeData

_Check_return_ HRESULT CHierarchyTableTreeCtrl::AddRootNode(_In_ LPMAPICONTAINER lpMAPIContainer)
{
	HRESULT hRes = S_OK;
	LPSPropValue lpProps = NULL;
	LPSPropValue lpRootName = NULL; // don't free
	LPSBinary lpEIDBin = NULL; // don't free

	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_hWnd) return S_OK;

	if (!lpMAPIContainer) return MAPI_E_INVALID_PARAMETER;

	ULONG cVals = 0;

	WC_H_GETPROPS(lpMAPIContainer->GetProps(
		(LPSPropTagArray)&sptHTCols,
		fMapiUnicode,
		&cVals,
		&lpProps));
	hRes = S_OK;

	// Get the entry ID for the Root Container
	if (!lpProps || PT_ERROR == PROP_TYPE(lpProps[htPR_ENTRYID].ulPropTag))
	{
		DebugPrint(DBGHierarchy, L"Could not find EntryID for Root Container. This is benign. Assuming NULL.\n");
		lpEIDBin = NULL;
	}
	else lpEIDBin = &lpProps[htPR_ENTRYID].Value.bin;

	// Get the Display Name for the Root Container
	if (!lpProps || PT_ERROR == PROP_TYPE(lpProps[htPR_DISPLAY_NAME].ulPropTag))
	{
		DebugPrint(DBGHierarchy, L"Could not find Display Name for Root Container. This is benign. Assuming NULL.\n");
		lpRootName = NULL;
	}
	else lpRootName = &lpProps[htPR_DISPLAY_NAME];

	wstring szName;

	// Shouldn't have to check lpRootName for non-NULL since CheckString does it, but prefast is complaining
	if (lpRootName && CheckStringProp(lpRootName, PT_TSTRING))
	{
		szName = LPCTSTRToWstring(lpRootName->Value.LPSZ);
	}
	else
	{
		szName = loadstring(IDS_ROOTCONTAINER);
	}

	SortListData* lpData = BuildNodeData(
		cVals,
		lpProps, // pass our lpProps to be archived
		lpEIDBin,
		NULL,
		true, // Always assume root nodes have children so we always paint an expanding icon
		lpProps ? lpProps[htPR_CONTAINER_FLAGS].Value.ul : MAPI_E_NOT_FOUND);

	AddNode(
		szName,
		TVI_ROOT,
		lpData,
		true);

	// Node owns the lpProps memory now, so we don't free it
	return hRes;
}

void CHierarchyTableTreeCtrl::AddNode(
	_In_ wstring szName,
	HTREEITEM hParent,
	SortListData* lpData,
	bool bGetTable)
{
	HTREEITEM hItem = NULL;

	DebugPrintEx(DBGHierarchy, CLASS, L"AddNode", L"Adding Node \"%ws\" under node %p, bGetTable = 0x%X\n", szName.c_str(), hParent, bGetTable);
	TVINSERTSTRUCT tvInsert = { 0 };

	tvInsert.hParent = hParent;
	tvInsert.hInsertAfter = TVI_SORT;
	tvInsert.item.mask = TVIF_CHILDREN | TVIF_TEXT;
	tvInsert.item.cChildren = I_CHILDRENCALLBACK;
	tvInsert.item.pszText = wstringToLPTSTR(szName);

	hItem = TreeView_InsertItem(m_hWnd, &tvInsert);

	SetNodeData(m_hWnd, hItem, lpData);

	if (bGetTable &&
		(RegKeys[regkeyHIER_ROOT_NOTIFS].ulCurDWORD || hParent != TVI_ROOT))
	{
		(void)GetHierarchyTable(hItem, NULL, true);
	}
}

void CHierarchyTableTreeCtrl::AddNode(_In_ LPSRow lpsRow, HTREEITEM hParent, bool bGetTable)
{
	if (!lpsRow) return;

	wstring szName;
	LPSPropValue lpName = NULL; // don't free

	lpName = PpropFindProp(
		lpsRow->lpProps,
		lpsRow->cValues,
		PR_DISPLAY_NAME);
	if (CheckStringProp(lpName, PT_TSTRING))
	{
		szName = LPCTSTRToWstring(lpName->Value.LPSZ);
	}
	else
	{
		szName = loadstring(IDS_UNKNOWNNAME);
	}
	DebugPrintEx(DBGHierarchy, CLASS, L"AddNode", L"Adding to %p: %ws\n", hParent, szName.c_str());

	SortListData* lpData = BuildNodeData(lpsRow);

	AddNode(
		szName,
		hParent,
		lpData,
		bGetTable);
}

_Check_return_ LPMAPITABLE CHierarchyTableTreeCtrl::GetHierarchyTable(HTREEITEM hItem, _In_opt_ LPMAPICONTAINER lpMAPIContainer, bool bRegNotifs)
{
	HRESULT hRes = S_OK;
	SortListData* lpData = NULL;

	lpData = (SortListData*)GetItemData(hItem);

	if (!lpData) return NULL;

	if (!lpData->data.Node.lpHierarchyTable)
	{
		if (lpMAPIContainer)
		{
			lpMAPIContainer->AddRef();
		}
		else
		{
			GetContainer(hItem, mfcmapiDO_NOT_REQUEST_MODIFY, &lpMAPIContainer);
		}

		if (lpMAPIContainer)
		{
			// Get the hierarchy table for the node and shove it into the data
			LPMAPITABLE lpHierarchyTable = NULL;

			// on the AB, something about this call triggers table reloads on the parent hierarchy table
			// no idea why they're triggered - doesn't happen for all AB providers
			WC_MAPI(lpMAPIContainer->GetHierarchyTable(
				(m_ulDisplayFlags & dfDeleted ? SHOW_SOFT_DELETES : NULL) | fMapiUnicode,
				&lpHierarchyTable));

			if (lpHierarchyTable)
			{
				EC_MAPI(lpHierarchyTable->SetColumns(
					(LPSPropTagArray)&sptHTCols,
					TBL_BATCH));
			}

			lpData->data.Node.lpHierarchyTable = lpHierarchyTable;
			lpMAPIContainer->Release();
		}
	}

	if (lpData->data.Node.lpHierarchyTable && !lpData->data.Node.lpAdviseSink)
	{
		// set up our advise sink
		if (bRegNotifs &&
			(RegKeys[regkeyHIER_ROOT_NOTIFS].ulCurDWORD || GetRootItem() != hItem))
		{
			DebugPrintEx(DBGNotify, CLASS, L"GetHierarchyTable", L"Advise sink for \"%ws\" = %p\n", LPCTSTRToWstring(GetItemText(hItem)).c_str(), hItem);
			lpData->data.Node.lpAdviseSink = new CAdviseSink(m_hWnd, hItem);

			if (lpData->data.Node.lpAdviseSink)
			{
				WC_MAPI(lpData->data.Node.lpHierarchyTable->Advise(
					fnevTableModified,
					(IMAPIAdviseSink *)lpData->data.Node.lpAdviseSink,
					&lpData->data.Node.ulAdviseConnection));
				if (MAPI_E_NO_SUPPORT == hRes) // Some tables don't support this!
				{
					if (lpData->data.Node.lpAdviseSink) lpData->data.Node.lpAdviseSink->Release();
					lpData->data.Node.lpAdviseSink = NULL;
					DebugPrint(DBGNotify, L"This table doesn't support notifications\n");
					hRes = S_OK; // mask the error
				}
				else if (S_OK == hRes)
				{
					LPMDB lpMDB = m_lpMapiObjects->GetMDB(); // do not release
					if (lpMDB)
					{
						lpData->data.Node.lpAdviseSink->SetAdviseTarget(lpMDB);

						LPSPropValue lpProp = NULL;
						// Try to trigger some RPC to get the notifications going
						WC_MAPI(HrGetOneProp(
							lpMDB,
							PR_TEST_LINE_SPEED,
							&lpProp));
						if (MAPI_E_NOT_FOUND == hRes)
						{
							// We're not on an Exchange server. We don't need to generate RPC after all.
							hRes = S_OK;
						}
						MAPIFreeBuffer(lpProp);
					}
				}
				DebugPrintEx(DBGNotify, CLASS, L"GetHierarchyTable", L"Advise sink %p, ulAdviseConnection = 0x%08X\n", lpData->data.Node.lpAdviseSink, (int)lpData->data.Node.ulAdviseConnection);
			}
		}
	}
	return lpData->data.Node.lpHierarchyTable;
}

// Add the first level contents of lpMAPIContainer under the Parent node
_Check_return_ HRESULT CHierarchyTableTreeCtrl::ExpandNode(HTREEITEM hParent)
{
	LPMAPITABLE lpHierarchyTable = NULL;
	LPSRowSet pRows = NULL;
	HRESULT hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_hWnd) return S_OK;
	if (!hParent) return MAPI_E_INVALID_PARAMETER;
	DebugPrintEx(DBGHierarchy, CLASS, L"ExpandNode", L"Expanding %p\n", hParent);

	lpHierarchyTable = GetHierarchyTable(hParent, NULL, (0 != RegKeys[regkeyHIER_EXPAND_NOTIFS].ulCurDWORD));

	if (lpHierarchyTable)
	{
		// go to the first row
		EC_MAPI(lpHierarchyTable->SeekRow(
			BOOKMARK_BEGINNING,
			0,
			NULL));
		hRes = S_OK; // don't let failure here fail the whole load

		ULONG i = 0;
		// get each row in turn and add it to the list
		// TODO: Query several rows at once
		if (!FAILED(hRes)) for (;;)
		{
			hRes = S_OK;
			// Note - we're saving the rows off in AddNode, so we don't FreeProws this...we just MAPIFreeBuffer the array
			if (pRows) MAPIFreeBuffer(pRows);
			pRows = NULL;
			EC_MAPI(lpHierarchyTable->QueryRows(
				1,
				NULL,
				&pRows));
			if (FAILED(hRes) || !pRows || !pRows->cRows) break;
			// Now we can process the row!

			AddNode(
				pRows->aRow,
				hParent,
				false);
			i++;
		}
	}

	// Note - we're saving the props off in AddNode, so we don't FreeProws this...we just MAPIFreeBuffer the array
	if (pRows) MAPIFreeBuffer(pRows);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
// Message Handlers
void CHierarchyTableTreeCtrl::OnGetDispInfo(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
{
	NMTVDISPINFO* lpDispInfo = (LPNMTVDISPINFO)pNMHDR;

	if (lpDispInfo &&
		lpDispInfo->item.mask & TVIF_CHILDREN)
	{
		SortListData* lpData = NULL;
		lpData = (SortListData*)lpDispInfo->item.lParam;

		if (lpData)
		{
			if (m_ulDisplayFlags & dfDeleted)
			{
				lpDispInfo->item.cChildren = 1;
			}
			else if (lpData->data.Node.cSubfolders >= 0)
			{
				lpDispInfo->item.cChildren = lpData->data.Node.cSubfolders ? 1 : 0;
			}
			else
			{
				LPCTSTR szName = NULL;
				if (PROP_TYPE(lpData->lpSourceProps[0].ulPropTag) == PT_TSTRING) szName = lpData->lpSourceProps[0].Value.LPSZ;
				DebugPrintEx(DBGHierarchy, CLASS, L"OnGetDispInfo", L"Using Hierarchy table %d %p %ws\n", lpData->data.Node.cSubfolders, lpData->data.Node.lpHierarchyTable, LPCTSTRToWstring(szName).c_str());
				// won't force the hierarchy table - just get it if we've already got it
				LPMAPITABLE lpHierarchyTable = lpData->data.Node.lpHierarchyTable;
				if (lpHierarchyTable)
				{
					lpDispInfo->item.cChildren = 1;
					HRESULT hRes = S_OK;
					ULONG ulRowCount = NULL;
					WC_MAPI(lpHierarchyTable->GetRowCount(
						NULL,
						&ulRowCount));
					if (S_OK == hRes && !ulRowCount)
					{
						lpDispInfo->item.cChildren = 0;
					}
				}
			}
		}
	}
	*pResult = 0;
}

void CHierarchyTableTreeCtrl::UpdateSelectionUI(HTREEITEM hItem)
{
	HRESULT hRes = S_OK;
	LPMAPICONTAINER lpMAPIContainer = NULL;
	LPSPropValue lpProps = NULL;
	ULONG cVals = 0;
	UINT uiMsg = IDS_STATUSTEXTNOFOLDER;
	wstring szParam1;
	wstring szParam2;
	wstring szParam3;

	DebugPrintEx(DBGHierarchy, CLASS, L"UpdateSelectionUI", L"%p\n", hItem);

	if (!m_lpHostDlg) return;

	// Have to request modify or this object is read only in the single prop control.
	GetContainer(hItem, mfcmapiREQUEST_MODIFY, &lpMAPIContainer);

	// make sure we've gotten the hierarchy table for this node
	(void)GetHierarchyTable(hItem, lpMAPIContainer, (0 != RegKeys[regkeyHIER_EXPAND_NOTIFS].ulCurDWORD));

	if (SUCCEEDED(hRes) && lpMAPIContainer)
	{
		// Get some props for status bar
		WC_H_GETPROPS(lpMAPIContainer->GetProps(
			(LPSPropTagArray)&sptHTCountCols,
			fMapiUnicode,
			&cVals,
			&lpProps));
		if (lpProps)
		{
			if (!(m_ulDisplayFlags & dfDeleted))
			{
				uiMsg = IDS_STATUSTEXTCONTENTCOUNTS;
				if (PT_ERROR == PROP_TYPE(lpProps[htcPR_CONTENT_COUNT].ulPropTag))
				{
					WARNHRESMSG(lpProps[htcPR_CONTENT_COUNT].Value.err, IDS_NODELACKSCONTENTCOUNT);
					szParam1 = L"?"; // STRING_OK
				}
				else szParam1 = format(L"%u", lpProps[htcPR_CONTENT_COUNT].Value.ul);

				if (PT_ERROR == PROP_TYPE(lpProps[htcPR_ASSOC_CONTENT_COUNT].ulPropTag))
				{
					WARNHRESMSG(lpProps[htcPR_ASSOC_CONTENT_COUNT].Value.err, IDS_NODELACKSASSOCCONTENTCOUNT);
					szParam2 = L"?"; // STRING_OK
				}
				else szParam2 = format(L"%u", lpProps[htcPR_ASSOC_CONTENT_COUNT].Value.ul);
			}
			else
			{
				uiMsg = IDS_STATUSTEXTDELETEDCOUNTS;
				if (PT_ERROR == PROP_TYPE(lpProps[htcPR_DELETED_MSG_COUNT].ulPropTag))
				{
					WARNHRESMSG(lpProps[htcPR_DELETED_MSG_COUNT].Value.err, IDS_NODELACKSDELETEDMESSAGECOUNT);
					szParam1 = L"?"; // STRING_OK
				}
				else szParam1 = format(L"%u", lpProps[htcPR_DELETED_MSG_COUNT].Value.ul);

				if (PT_ERROR == PROP_TYPE(lpProps[htcPR_DELETED_ASSOC_MSG_COUNT].ulPropTag))
				{
					WARNHRESMSG(lpProps[htcPR_DELETED_ASSOC_MSG_COUNT].Value.err, IDS_NODELACKSDELETEDASSOCMESSAGECOUNT);
					szParam2 = L"?"; // STRING_OK
				}
				else szParam2 = format(L"%u", lpProps[htcPR_DELETED_ASSOC_MSG_COUNT].Value.ul);

				if (PT_ERROR == PROP_TYPE(lpProps[htcPR_DELETED_FOLDER_COUNT].ulPropTag))
				{
					WARNHRESMSG(lpProps[htcPR_DELETED_FOLDER_COUNT].Value.err, IDS_NODELACKSDELETEDSUBFOLDERCOUNT);
					szParam3 = L"?"; // STRING_OK
				}
				else szParam3 = format(L"%u", lpProps[htcPR_DELETED_FOLDER_COUNT].Value.ul);
			}
			MAPIFreeBuffer(lpProps);
		}
	}

	m_lpHostDlg->UpdateStatusBarText(
		STATUSDATA1,
		uiMsg,
		szParam1,
		szParam2,
		szParam3);

	m_lpHostDlg->OnUpdateSingleMAPIPropListCtrl(lpMAPIContainer, (SortListData*)GetItemData(hItem));
	auto selectedItem = LPCTSTRToWstring(GetItemText(GetSelectedItem()));
	m_lpHostDlg->UpdateTitleBarText(selectedItem);

	if (lpMAPIContainer) lpMAPIContainer->Release();
}

_Check_return_ bool CHierarchyTableTreeCtrl::IsItemSelected()
{
	return m_bItemSelected;
}

void CHierarchyTableTreeCtrl::OnSelChanged(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
{
	LPNMTREEVIEW pNMTV = (LPNMTREEVIEW)pNMHDR;

	if (pNMTV && pNMTV->itemNew.hItem)
	{
		m_bItemSelected = true;

		UpdateSelectionUI(pNMTV->itemNew.hItem);
	}
	else
	{
		m_bItemSelected = false;
	}

	*pResult = 0;
}

// This function will be called when we edit a node so we can attempt to commit the changes
void CHierarchyTableTreeCtrl::OnEndLabelEdit(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
{
	HRESULT hRes = S_OK;
	LPMAPICONTAINER lpMAPIContainer = NULL;
	SPropValue sDisplayName;
	TV_DISPINFO* pTVDispInfo = (TV_DISPINFO*)pNMHDR;

	*pResult = 0;

	if (!pTVDispInfo || !pTVDispInfo->item.pszText) return;

	GetContainer(pTVDispInfo->item.hItem, mfcmapiREQUEST_MODIFY, &lpMAPIContainer);
	if (!lpMAPIContainer) return;

	sDisplayName.ulPropTag = PR_DISPLAY_NAME;
	sDisplayName.Value.LPSZ = pTVDispInfo->item.pszText;

	EC_MAPI(HrSetOneProp(lpMAPIContainer, &sDisplayName));

	lpMAPIContainer->Release();
}

void CHierarchyTableTreeCtrl::OnDblclk(_In_ NMHDR* /*pNMHDR*/, _In_ LRESULT* pResult)
{
	// Due to problems with focus...we have to post instead of calling OnDisplayItem directly.
	// Post the message to display the item
	if (m_lpHostDlg)
		m_lpHostDlg->PostMessage(WM_COMMAND, ID_DISPLAYSELECTEDITEM, NULL);

	// Don't do default behavior for double-click (We only want '+' sign expansion.
	// Double click should display the item, not expand the tree.)
	*pResult = 1;
}

void CHierarchyTableTreeCtrl::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	DebugPrintEx(DBGMenu, CLASS, L"OnKeyDown", L"0x%X\n", nChar);

	bool bCtrlPressed = GetKeyState(VK_CONTROL) < 0;
	bool bShiftPressed = GetKeyState(VK_SHIFT) < 0;
	bool bMenuPressed = GetKeyState(VK_MENU) < 0;

	if (!bMenuPressed)
	{
		if (VK_RETURN == nChar && bCtrlPressed)
		{
			DebugPrintEx(DBGGeneric, CLASS, L"OnKeyDown", L"calling Display Associated Contents\n");
			if (m_lpHostDlg)
				m_lpHostDlg->PostMessage(WM_COMMAND, ID_DISPLAYASSOCIATEDCONTENTS, NULL);
		}
		else if (!m_lpHostDlg || !m_lpHostDlg->HandleKeyDown(nChar, bShiftPressed, bCtrlPressed, bMenuPressed))
		{
			CTreeCtrl::OnKeyDown(nChar, nRepCnt, nFlags);
		}
	}
}

// Assert that we want all keyboard input (including ENTER!)
_Check_return_ UINT CHierarchyTableTreeCtrl::OnGetDlgCode()
{
	UINT iDlgCode = CTreeCtrl::OnGetDlgCode();

	iDlgCode |= DLGC_WANTMESSAGE;

	// to make sure that the control key is not pressed
	if ((GetKeyState(VK_CONTROL) >= 0) && (m_hWnd == ::GetFocus()))
	{
		// to make sure that the Tab key is pressed
		if (GetKeyState(VK_TAB) < 0)
			iDlgCode &= ~(DLGC_WANTALLKEYS | DLGC_WANTMESSAGE | DLGC_WANTTAB);
	}

	return iDlgCode;
}

void CHierarchyTableTreeCtrl::OnRightClick(_In_ NMHDR* /*pNMHDR*/, _In_ LRESULT* pResult)
{
	// Send WM_CONTEXTMENU to self
	(void)SendMessage(WM_CONTEXTMENU, (WPARAM)m_hWnd, GetMessagePos());

	// Mark message as handled and suppress default handling
	*pResult = 1;
}

void CHierarchyTableTreeCtrl::OnContextMenu(_In_ CWnd* pWnd, CPoint pos)
{
	if (pWnd && -1 == pos.x && -1 == pos.y)
	{
		// Find the highlighted item
		HTREEITEM item = GetSelectedItem();

		if (item)
		{
			RECT rc = { 0 };
			GetItemRect(item, &rc, true);
			pos.x = rc.left;
			pos.y = rc.top;
			::ClientToScreen(pWnd->m_hWnd, &pos);
		}
	}
	else
	{
		// Select the item that is at the point pos.
		UINT uFlags = NULL;
		CPoint ptTree = pos;
		ScreenToClient(&ptTree);
		HTREEITEM hClickedItem = HitTest(ptTree, &uFlags);

		if ((hClickedItem != NULL) && (TVHT_ONITEM & uFlags))
		{
			Select(hClickedItem, TVGN_CARET);
		}
	}

	DisplayContextMenu(m_nIDContextMenu, IDR_MENU_HIERARCHY_TABLE, m_lpHostDlg->m_hWnd, pos.x, pos.y);
}

_Check_return_ SortListData* CHierarchyTableTreeCtrl::GetSelectedItemData()
{
	HTREEITEM Item = NULL;

	// Find the highlighted item
	Item = GetSelectedItem();

	if (Item)
	{
		return (SortListData*)GetItemData(Item);
	}
	return NULL;
}

_Check_return_ LPSBinary CHierarchyTableTreeCtrl::GetSelectedItemEID()
{
	HTREEITEM Item = NULL;

	// Find the highlighted item
	Item = GetSelectedItem();

	// get the EID associated with it
	if (Item)
	{
		SortListData* lpData = NULL;
		lpData = (SortListData*)GetItemData(Item);
		if (lpData)
			return lpData->data.Node.lpEntryID;
	}
	return NULL;
}

_Check_return_ LPMAPICONTAINER CHierarchyTableTreeCtrl::GetSelectedContainer(__mfcmapiModifyEnum bModify)
{
	LPMAPICONTAINER lpSelectedContainer = NULL;

	GetContainer(GetSelectedItem(), bModify, &lpSelectedContainer);

	return lpSelectedContainer;
}

void CHierarchyTableTreeCtrl::GetContainer(
	HTREEITEM Item,
	__mfcmapiModifyEnum bModify,
	_In_ LPMAPICONTAINER* lppContainer)
{
	HRESULT hRes = S_OK;
	ULONG ulObjType = NULL;
	ULONG ulFlags = NULL;
	SortListData* lpData = NULL;
	LPSBinary lpCurBin = NULL;
	SBinary NullBin = { 0 };
	LPMAPICONTAINER lpContainer = NULL;

	*lppContainer = NULL;

	if (!Item) return;

	DebugPrintEx(DBGGeneric, CLASS, L"GetContainer", L"HTREEITEM = %p, bModify = %d, m_ulContainerType = 0x%X\n", Item, bModify, m_ulContainerType);

	lpData = (SortListData*)GetItemData(Item);

	if (!lpData)
	{
		// We didn't get an entryID, so log it and get out of here
		DebugPrintEx(DBGGeneric, CLASS, L"GetContainer", L"GetItemData returned NULL or lpEntryID is NULL\n");
		return;
	}

	ulFlags = (mfcmapiREQUEST_MODIFY == bModify ? MAPI_MODIFY : NULL);

	lpCurBin = lpData->data.Node.lpEntryID;
	if (!lpCurBin) lpCurBin = &NullBin;

	// Check the type of the root container to know whether the MDB or AddrBook object is valid
	// This also allows NULL EIDs to return the root container itself.
	// Use the Root container if we can't decide and log an error
	if (m_lpMapiObjects)
	{
		if (MAPI_ABCONT == m_ulContainerType)
		{
			LPADRBOOK lpAddrBook = m_lpMapiObjects->GetAddrBook(false); // do not release
			if (lpAddrBook)
			{
				DebugPrint(DBGGeneric, L"\tCalling OpenEntry on address book with ulFlags = 0x%X\n", ulFlags);

				WC_H(CallOpenEntry(
					NULL,
					lpAddrBook,
					NULL,
					NULL,
					lpCurBin->cb,
					(LPENTRYID)lpCurBin->lpb,
					NULL,
					ulFlags,
					&ulObjType,
					(LPUNKNOWN*)&lpContainer));
			}
		}
		else if (MAPI_FOLDER == m_ulContainerType)
		{
			LPMDB lpMDB = m_lpMapiObjects->GetMDB(); // do not release
			if (lpMDB)
			{
				ulFlags = (mfcmapiREQUEST_MODIFY == bModify ? MAPI_MODIFY : NULL) |
					(m_ulDisplayFlags & dfDeleted ? (SHOW_SOFT_DELETES | MAPI_NO_CACHE) : NULL);

				WC_H(CallOpenEntry(
					lpMDB,
					NULL,
					NULL,
					NULL,
					lpCurBin->cb,
					(LPENTRYID)lpCurBin->lpb,
					NULL,
					ulFlags,
					&ulObjType,
					(LPUNKNOWN*)&lpContainer));
			}
		}
	}

	// If we didn't get a container above, try to open from m_lpContainer
	if (!lpContainer && m_lpContainer)
	{
		WARNHRESMSG(MAPI_E_CALL_FAILED, IDS_UNKNOWNCONTAINERTYPE);
		hRes = S_OK;
		WC_H(CallOpenEntry(
			NULL,
			NULL,
			m_lpContainer,
			NULL,
			lpCurBin->cb,
			(LPENTRYID)lpCurBin->lpb,
			NULL,
			ulFlags,
			&ulObjType,
			(LPUNKNOWN*)&lpContainer));
	}

	// if we failed because write access was denied, try again if acceptable
	if (!lpContainer && FAILED(hRes) && mfcmapiREQUEST_MODIFY == bModify)
	{
		DebugPrint(DBGGeneric, L"\tOpenEntry failed: 0x%X. Will try again without MAPI_MODIFY\n", hRes);
		// We failed to open the item with MAPI_MODIFY.
		// Let's try to open it with NULL
		GetContainer(
			Item,
			mfcmapiDO_NOT_REQUEST_MODIFY,
			&lpContainer);
	}

	// Ok - we're just out of luck
	if (!lpContainer)
	{
		WARNHRESMSG(hRes, IDS_NOCONTAINER);
		hRes = MAPI_E_NOT_FOUND;
	}

	if (lpContainer) *lppContainer = lpContainer;
	DebugPrintEx(DBGGeneric, CLASS, L"GetContainer", L"returning lpContainer = %p, ulObjType = 0x%X and hRes = 0x%X\n", lpContainer, ulObjType, hRes);
}

///////////////////////////////////////////////////////////////////////////////
// End - Message Handlers

// When + is clicked, add all entries in the table as children
void CHierarchyTableTreeCtrl::OnItemExpanding(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
{
	HRESULT hRes = S_OK;

	*pResult = 0;

	NM_TREEVIEW* pNMTreeView = (NM_TREEVIEW*)pNMHDR;
	if (pNMTreeView)
	{
		DebugPrintEx(DBGHierarchy, CLASS, L"OnItemExpanding", L"Expanding item %p \"%ws\" action = 0x%08X state = 0x%08X\n", pNMTreeView->itemNew.hItem, LPCTSTRToWstring(GetItemText(pNMTreeView->itemOld.hItem)).c_str(), pNMTreeView->action, pNMTreeView->itemNew.state);
		if (pNMTreeView->action & TVE_EXPAND)
		{
			if (!(pNMTreeView->itemNew.state & TVIS_EXPANDEDONCE))
			{
				EC_H(ExpandNode(pNMTreeView->itemNew.hItem));
			}
		}
	}
}

// Tree control will call this for every node it deletes.
void CHierarchyTableTreeCtrl::OnDeleteItem(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
{
	LPNMTREEVIEW pNMTreeView = (LPNMTREEVIEW)pNMHDR;
	if (pNMTreeView)
	{
		SortListData* lpData = (SortListData*)pNMTreeView->itemOld.lParam;
		DebugPrintEx(DBGHierarchy, CLASS, L"OnDeleteItem", L"Deleting item %p \"%ws\"\n", pNMTreeView->itemOld.hItem, LPCTSTRToWstring(GetItemText(pNMTreeView->itemOld.hItem)).c_str());

		if (lpData && lpData->data.Node.lpAdviseSink)
		{
			DebugPrintEx(DBGHierarchy, CLASS, L"OnDeleteItem", L"Unadvising %p, ulAdviseConnection = 0x%08X\n", lpData->data.Node.lpAdviseSink, (int)lpData->data.Node.ulAdviseConnection);
		}
		if (lpData) FreeSortListData(lpData);
		lpData = NULL;

		if (!m_bShuttingDown)
		{
			// Collapse the parent if this is the last child
			HTREEITEM hPrev = TreeView_GetPrevSibling(m_hWnd, pNMTreeView->itemOld.hItem);
			HTREEITEM hNext = TreeView_GetNextSibling(m_hWnd, pNMTreeView->itemOld.hItem);

			if (!(hPrev || hNext))
			{
				DebugPrintEx(DBGHierarchy, CLASS, L"OnDeleteItem", L"%p has no siblings\n", pNMTreeView->itemOld.hItem);
				HTREEITEM hParent = TreeView_GetParent(m_hWnd, pNMTreeView->itemOld.hItem);
				TreeView_SetItemState(m_hWnd, hParent, 0, TVIS_EXPANDED | TVIS_EXPANDEDONCE);
				TVITEM tvItem = { 0 };
				tvItem.hItem = hParent;
				tvItem.mask = TVIF_PARAM;
				if (TreeView_GetItem(m_hWnd, &tvItem) && tvItem.lParam)
				{
					lpData = (SortListData*)tvItem.lParam;
					lpData->data.Node.cSubfolders = 0;
				}
			}
		}
	}

	*pResult = 0;
}

// WM_MFCMAPI_ADDITEM
// If the parent has been expanded once, add the new row
// Otherwise, ditch the notification
_Check_return_ LRESULT CHierarchyTableTreeCtrl::msgOnAddItem(WPARAM wParam, LPARAM lParam)
{
	TABLE_NOTIFICATION* tab = (TABLE_NOTIFICATION*)wParam;
	HTREEITEM hParent = (HTREEITEM)lParam;

	DebugPrintEx(DBGHierarchy, CLASS, L"msgOnAddItem", L"Received message add item under: %p =\"%ws\"\n", hParent, LPCTSTRToWstring(GetItemText(hParent)).c_str());

	// only need to add the node if we're expanded
	int iState = GetItemState(hParent, NULL);
	if (iState & TVIS_EXPANDEDONCE)
	{
		HRESULT hRes = S_OK;
		// We make this copy here and pass it in to AddNode, where it is grabbed by BuildDataItem to be part of the item data
		// The mem will be freed when the item data is cleaned up - do not free here
		SRow NewRow = { 0 };
		NewRow.cValues = tab->row.cValues;
		NewRow.ulAdrEntryPad = tab->row.ulAdrEntryPad;
		WC_MAPI(ScDupPropset(
			tab->row.cValues,
			tab->row.lpProps,
			MAPIAllocateBuffer,
			&NewRow.lpProps));
		AddNode(&NewRow, hParent, true);
	}
	else
	{
		// in case the item doesn't know it has children, let it know
		TVITEM tvItem = { 0 };
		tvItem.hItem = hParent;
		tvItem.mask = TVIF_PARAM;
		if (TreeView_GetItem(m_hWnd, &tvItem) && tvItem.lParam)
		{
			SortListData* lpData = (SortListData*)tvItem.lParam;
			lpData->data.Node.cSubfolders = 1;
		}
	}

	return S_OK;
}

// WM_MFCMAPI_DELETEITEM
// Remove the child node.
_Check_return_ LRESULT CHierarchyTableTreeCtrl::msgOnDeleteItem(WPARAM wParam, LPARAM lParam)
{
	HRESULT hRes = S_OK;
	TABLE_NOTIFICATION* tab = (TABLE_NOTIFICATION*)wParam;
	HTREEITEM hParent = (HTREEITEM)lParam;

	HTREEITEM hItemToDelete = FindNode(
		&tab->propIndex.Value.bin,
		hParent);

	if (hItemToDelete)
	{
		DebugPrintEx(DBGHierarchy, CLASS, L"msgOnDeleteItem", L"Received message delete item: %p =\"%ws\"\n", hItemToDelete, LPCTSTRToWstring(GetItemText(hItemToDelete)).c_str());
		EC_B(DeleteItem(hItemToDelete));
	}

	return hRes;
}

// WM_MFCMAPI_MODIFYITEM
// Update any UI for the node and resort if needed.
_Check_return_ LRESULT CHierarchyTableTreeCtrl::msgOnModifyItem(WPARAM wParam, LPARAM lParam)
{
	HRESULT hRes = S_OK;
	TABLE_NOTIFICATION* tab = (TABLE_NOTIFICATION*)wParam;
	HTREEITEM hParent = (HTREEITEM)lParam;

	HTREEITEM hModifyItem = FindNode(
		&tab->propIndex.Value.bin,
		hParent);

	if (hModifyItem)
	{
		DebugPrintEx(DBGHierarchy, CLASS, L"msgOnModifyItem", L"Received message modify item: %p =\"%ws\"\n", hModifyItem, LPCTSTRToWstring(GetItemText(hModifyItem)).c_str());

		LPSPropValue lpName = NULL; // don't free
		lpName = PpropFindProp(
			tab->row.lpProps,
			tab->row.cValues,
			PR_DISPLAY_NAME);

		if (CheckStringProp(lpName, PT_TSTRING))
		{
			EC_B(SetItemText(hModifyItem, lpName->Value.LPSZ));
		}
		else
		{
			CString szText;
			EC_B(szText.LoadString(IDS_UNKNOWNNAME));
			EC_B(SetItemText(hModifyItem, szText));
		}

		// We make this copy here and pass it in to the node
		// The mem will be freed when the item data is cleaned up - do not free here
		SRow NewRow = { 0 };
		NewRow.cValues = tab->row.cValues;
		NewRow.ulAdrEntryPad = tab->row.ulAdrEntryPad;
		WC_MAPI(ScDupPropset(
			tab->row.cValues,
			tab->row.lpProps,
			MAPIAllocateBuffer,
			&NewRow.lpProps));
		SortListData* lpData = BuildNodeData(&NewRow);
		SetNodeData(m_hWnd, hModifyItem, lpData);
		if (hParent) EC_B(SortChildren(hParent));
	}

	if (hModifyItem == GetSelectedItem()) UpdateSelectionUI(hModifyItem);

	return hRes;
}

// WM_MFCMAPI_REFRESHTABLE
// If node was expanded, collapse it to remove all children
// Then, if the node does have children, reexpand it.
_Check_return_ LRESULT CHierarchyTableTreeCtrl::msgOnRefreshTable(WPARAM wParam, LPARAM /*lParam*/)
{
	HRESULT hRes = S_OK;
	HTREEITEM hRefreshItem = (HTREEITEM)wParam;
	DebugPrintEx(DBGHierarchy, CLASS, L"msgOnRefreshTable", L"Received message refresh table: %p =\"%ws\"\n", hRefreshItem, LPCTSTRToWstring(GetItemText(hRefreshItem)).c_str());

	int iState = GetItemState(hRefreshItem, NULL);
	if (iState & TVIS_EXPANDED)
	{
		HTREEITEM hChild = GetChildItem(hRefreshItem);
		while (hChild)
		{
			hRes = S_OK;
			EC_B(DeleteItem(hChild));
			hChild = GetChildItem(hRefreshItem);
		}
		// Reset our expanded bits
		EC_B(SetItemState(hRefreshItem, NULL, TVIS_EXPANDED | TVIS_EXPANDEDONCE));
		hRes = S_OK;
		{
			SortListData* lpData = NULL;
			lpData = (SortListData*)GetItemData(hRefreshItem);

			if (lpData)
			{
				if (lpData->data.Node.lpHierarchyTable)
				{
					ULONG ulRowCount = NULL;
					WC_MAPI(lpData->data.Node.lpHierarchyTable->GetRowCount(
						NULL,
						&ulRowCount));
					if (S_OK != hRes || ulRowCount)
					{
						EC_B(Expand(hRefreshItem, TVE_EXPAND));
					}
				}
			}
		}

	}

	return hRes;
}

// This function steps through the list control to find the entry with this instance key
// return NULL if item not found
_Check_return_ HTREEITEM CHierarchyTableTreeCtrl::FindNode(_In_ LPSBinary lpInstance, HTREEITEM hParent)
{
	if (!lpInstance || !hParent) return NULL;

	DebugPrintEx(DBGGeneric, CLASS, L"FindNode", L"Looking for child of: %p =\"%ws\"\n", hParent, LPCTSTRToWstring(GetItemText(hParent)).c_str());

	LPSBinary lpCurInstance = NULL;
	SortListData* lpListData = NULL;
	HTREEITEM hCurrent = NULL;

	hCurrent = GetNextItem(hParent, TVGN_CHILD);

	while (hCurrent)
	{
		lpListData = (SortListData*)GetItemData(hCurrent);
		if (lpListData)
		{
			lpCurInstance = lpListData->data.Node.lpInstanceKey;
			if (lpCurInstance)
			{
				if (!memcmp(lpCurInstance->lpb, lpInstance->lpb, lpInstance->cb))
				{
					DebugPrintEx(DBGGeneric, CLASS, L"FindNode", L"Matched at %p =\"%ws\"\n", hCurrent, LPCTSTRToWstring(GetItemText(hCurrent)).c_str());
					return hCurrent;
				}
			}
		}
		hCurrent = GetNextItem(hCurrent, TVGN_NEXT);
	}

	DebugPrintEx(DBGGeneric, CLASS, L"FindNode", L"No match found\n");
	return NULL;
}

void CHierarchyTableTreeCtrl::OnCustomDraw(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
{
	CustomDrawTree(pNMHDR, pResult, m_HoverButton, m_hItemCurHover);
}