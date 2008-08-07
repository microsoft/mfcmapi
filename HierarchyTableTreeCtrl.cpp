// HierarchyTableTreeCtrl.cpp : implementation file
//

#include "stdafx.h"
#include "Error.h"

#include "HierarchyTableTreeCtrl.h"

#include "BaseDialog.h"
#include "HierarchyTableDlg.h"
#include "MapiObjects.h"
#include "MAPIFunctions.h"
#include "MFCUtilityFunctions.h"
#include "AdviseSink.h"
#include "registry.h"
#include "SingleMAPIPropListCtrl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

static TCHAR* CLASS = _T("CHierarchyTableTreeCtrl");

/////////////////////////////////////////////////////////////////////////////
// CHierarchyTableTreeCtrl

CHierarchyTableTreeCtrl::CHierarchyTableTreeCtrl(
												 CWnd* pCreateParent,
												 CMapiObjects *lpMapiObjects,
												 CHierarchyTableDlg *lpHostDlg,
												 ULONG ulDisplayFlags)
{
	TRACE_CONSTRUCTOR(CLASS);
	CRect pRect;

	m_cRef = 1;

	//We borrow our parent's Mapi objects
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

	m_bItemSelected = FALSE;

	Create(
		TVS_HASBUTTONS
		| TVS_HASLINES
		| TVS_LINESATROOT
		| TVS_EDITLABELS
		| TVS_DISABLEDRAGDROP
		| TVS_SHOWSELALWAYS
		| WS_CHILD
//		| WS_BORDER
		| WS_TABSTOP
//		| WS_CLIPCHILDREN
		| WS_CLIPSIBLINGS
		| WS_VISIBLE,
		pRect,
		pCreateParent,
		IDC_FOLDER_TREE);
}

CHierarchyTableTreeCtrl::~CHierarchyTableTreeCtrl()
{
	TRACE_DESTRUCTOR(CLASS);
	DestroyWindow();

	if (m_lpHostDlg) m_lpHostDlg->Release();
	if (m_lpMapiObjects) m_lpMapiObjects->Release();
}

STDMETHODIMP_(ULONG) CHierarchyTableTreeCtrl::AddRef()
{
	LONG lCount = InterlockedIncrement(&m_cRef);
	TRACE_ADDREF(CLASS,lCount);
	return lCount;
}

STDMETHODIMP_(ULONG) CHierarchyTableTreeCtrl::Release()
{
	LONG lCount = InterlockedDecrement(&m_cRef);
	TRACE_RELEASE(CLASS,lCount);
	if (!lCount)  delete this;
	return lCount;
}

BEGIN_MESSAGE_MAP(CHierarchyTableTreeCtrl, CTreeCtrl)
//{{AFX_MSG_MAP(CHierarchyTableTreeCtrl)
	ON_NOTIFY_REFLECT(TVN_GETDISPINFO, OnGetDispInfo)
	ON_NOTIFY_REFLECT(TVN_SELCHANGED, OnSelChanged)
	ON_NOTIFY_REFLECT(TVN_ENDLABELEDIT, OnEndLabelEdit)
	ON_NOTIFY_REFLECT(TVN_ITEMEXPANDING, OnItemExpanding)
	ON_NOTIFY_REFLECT(TVN_DELETEITEM, OnDeleteItem)
	ON_NOTIFY_REFLECT(NM_DBLCLK, OnDblclk)
	ON_NOTIFY_REFLECT(NM_RCLICK,OnRightClick)
	ON_WM_KEYDOWN()
	ON_WM_GETDLGCODE()
	ON_WM_CONTEXTMENU()
    ON_MESSAGE(WM_MFCMAPI_ADDITEM, msgOnAddItem)
	ON_MESSAGE(WM_MFCMAPI_DELETEITEM, msgOnDeleteItem)
	ON_MESSAGE(WM_MFCMAPI_MODIFYITEM, msgOnModifyItem)
	ON_MESSAGE(WM_MFCMAPI_REFRESHTABLE, msgOnRefreshTable)
//}}AFX_MSG_MAP
END_MESSAGE_MAP()


LRESULT CHierarchyTableTreeCtrl::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
//	LRESULT lResult = NULL;
//	DebugPrint(DBGWindowProc,_T("CHierarchyTableTreeCtrl::WindowProc message = 0x%x, wParam = 0x%X, lParam = 0x%X\n"),message,wParam,lParam);

	switch (message)
	{
	// I cannot handle notify's heading to my parent - have to depend on reflection for that
	case WM_NOTIFY:
		{
//			LPNMHDR pHdr = (LPNMHDR) lParam;
//			DebugPrint(DBGWindowProc,_T("CHierarchyTableTreeCtrl::OnNotify code = 0x%X, hwndFrom = 0x%X, idFrom = 0x%X\n"),pHdr->code, pHdr->hwndFrom,pHdr->idFrom);

//			switch(pHdr->code)
//			{
//			}
			break;
		}
	case WM_ERASEBKGND:
		{
			CTreeCtrl::OnEraseBkgnd((CDC*) wParam);
			return TRUE;
			break;
		}
	case WM_DESTROY:
		{
//			CTreeCtrl::OnDestroy();//just let the WindowProc handle it
		}
	}//end switch
	return CTreeCtrl::WindowProc(message,wParam,lParam);
}

/////////////////////////////////////////////////////////////////////////////
// CHierarchyTableTreeCtrl message handlers

HRESULT CHierarchyTableTreeCtrl::RefreshHierarchyTable()
{
	HRESULT			hRes = S_OK;

	//Turn off redraw while we work on the window
	SetRedraw(FALSE);

	m_bItemSelected = FALSE;//clear this just in case

	EC_B(DeleteItem(GetRootItem()));

	if (m_lpContainer)
		EC_H(AddRootNode(m_lpContainer));

	if (m_lpHostDlg) m_lpHostDlg->OnUpdateSingleMAPIPropListCtrl(NULL, NULL);

	//Turn redraw back on to update our view
	SetRedraw(TRUE);
	return hRes;
}

HRESULT CHierarchyTableTreeCtrl::LoadHierarchyTable(LPMAPICONTAINER lpMAPIContainer)
{
	HRESULT			hRes = S_OK;
	CWaitCursor		Wait;//Change the mouse to an hourglass while we work.

	if (m_lpContainer) m_lpContainer->Release();
	m_lpContainer = lpMAPIContainer;
	m_ulContainerType = NULL;

	//If we weren't passed a container to load, give up
	if (!m_lpContainer) return hRes;
	m_lpContainer->AddRef();

	m_ulContainerType = GetMAPIObjectType(lpMAPIContainer);

	WC_H(RefreshHierarchyTable());
	if (MAPI_E_NOT_FOUND == hRes)
	{
		WARNHRESMSG(hRes,IDS_HIERARCHNOTFOUND);
	}
	else CHECKHRESMSG(hRes,IDS_REFRESHHIERARCHYFAILED);

	return hRes;
}

HRESULT CHierarchyTableTreeCtrl::AddRootNode(LPMAPICONTAINER lpMAPIContainer)
{
	HRESULT			hRes		= S_OK;
	LPSPropValue	lpProps		= NULL;
	LPSPropValue	lpRootName	= NULL;//don't free
	LPSBinary		lpEIDBin	= NULL;//don't free

	CWaitCursor	Wait;//Change the mouse to an hourglass while we work.

	if (!m_hWnd) return S_OK;

	if (!lpMAPIContainer) return MAPI_E_INVALID_PARAMETER;

	enum{
		htPR_ENTRYID,
		htPR_DISPLAY_NAME,
		htPR_SUBFOLDERS,
		htPR_CONTAINER_FLAGS,
		htNUMCOLS
	};
	SizedSPropTagArray(htNUMCOLS,sptHTCols) = {htNUMCOLS,
		PR_ENTRYID,
		PR_DISPLAY_NAME,
		PR_SUBFOLDERS,
		PR_CONTAINER_FLAGS
	};
	ULONG cVals = 0;

	WC_H_GETPROPS(lpMAPIContainer->GetProps(
		(LPSPropTagArray) &sptHTCols,
		fMapiUnicode,
		&cVals,
		&lpProps));
	hRes = S_OK;

	//Get the entry ID for the Root Container
	if (!lpProps || PT_ERROR == PROP_TYPE(lpProps[htPR_ENTRYID].ulPropTag))
	{
		DebugPrint(DBGGeneric,_T("Could not find EntryID for Root Container. This is benign. Assuming NULL.\n"));
		lpEIDBin = NULL;
	}
	else lpEIDBin = &lpProps[htPR_ENTRYID].Value.bin;

	//Get the Display Name for the Root Container
	if (!lpProps || PT_ERROR == PROP_TYPE(lpProps[htPR_DISPLAY_NAME].ulPropTag))
	{
		DebugPrint(DBGGeneric,_T("Could not find Display Name for Root Container. This is benign. Assuming NULL.\n"));
		lpRootName = NULL;
	}
	else lpRootName = &lpProps[htPR_DISPLAY_NAME];

	CString szName;

	// Shouldn't have to check lpRootName for non-NULL since CheckString does it, but prefast is complaining
	if (lpRootName && CheckStringProp(lpRootName,PT_TSTRING))
	{
		szName = lpRootName->Value.LPSZ;
	}
	else
	{
		szName.LoadString(IDS_ROOTCONTAINER);
	}


	AddNode(
		cVals,
		lpProps,//pass our lpProps to be archived
		szName,
		lpEIDBin,
		NULL,
		lpProps?(ULONG)lpProps[htPR_SUBFOLDERS].Value.b:MAPI_E_NOT_FOUND,
		lpProps?lpProps[htPR_CONTAINER_FLAGS].Value.ul:MAPI_E_NOT_FOUND,
		TVI_ROOT,
		true);

	//Node owns the lpProps memory now
//	MAPIFreeBuffer(lpProps);
	return hRes;
}

void CHierarchyTableTreeCtrl::AddNode(ULONG			cProps,
									  LPSPropValue	lpProps,
									  LPCTSTR		szName,
									  LPSBinary		lpEntryID,
									  LPSBinary		lpInstanceKey,
									  ULONG			bSubfolders,
									  ULONG			ulContainerFlags,
									  HTREEITEM		hParent,
									  BOOL			bGetTable)
{
	HRESULT		hRes = S_OK;
	SortListData*	lpData = NULL;
	HTREEITEM	Item = NULL;

	DebugPrintEx(DBGGeneric,CLASS,_T("AddNode"),_T("Adding Node \"%s\" under node 0x%08X, bGetTable = 0x%X\n"),szName,hParent,bGetTable);
	//We're gonna set up a data item to pass off to the tree
	//Allocate some space
	EC_H(MAPIAllocateBuffer(
		(ULONG)sizeof(SortListData),
		(LPVOID *) &lpData));

	if (lpData)
	{
		memset(lpData, 0, sizeof(SortListData));
		lpData->ulSortDataType = SORTLIST_TREENODE;

		if (lpEntryID)
		{
			EC_H(MAPIAllocateMore(
				(ULONG)sizeof(SBinary),
				lpData,
				(LPVOID*) &lpData->data.Node.lpEntryID));

			//Copy the data over
			EC_H(CopySBinary(
				lpData->data.Node.lpEntryID,
				lpEntryID,
				lpData));
		}

		if (lpInstanceKey)
		{
			EC_H(MAPIAllocateMore(
				(ULONG)sizeof(SBinary),
				lpData,
				(LPVOID*) &lpData->data.Node.lpInstanceKey));
			EC_H(CopySBinary(
				lpData->data.Node.lpInstanceKey,
				lpInstanceKey,
				lpData));
		}
		lpData->data.Node.bSubfolders = bSubfolders;
		lpData->data.Node.ulContainerFlags = ulContainerFlags;
		lpData->cSourceProps = cProps;
		lpData->lpSourceProps = lpProps;

		TVINSERTSTRUCT tvInsert = {0};

		tvInsert.hParent = hParent;
		tvInsert.hInsertAfter = TVI_SORT;
		tvInsert.item.mask = TVIF_CHILDREN | TVIF_PARAM | TVIF_TEXT;
		tvInsert.item.cChildren = I_CHILDRENCALLBACK;
		tvInsert.item.lParam = (LPARAM) lpData;
		tvInsert.item.pszText = (LPTSTR) szName;

		Item = InsertItem(&tvInsert);

		if (bGetTable &&
			RegKeys[regkeyHIER_NOTIFS].ulCurDWORD &&
			(RegKeys[regkeyHIER_ROOT_NOTIFS].ulCurDWORD || hParent != TVI_ROOT))
		{
			GetHierarchyTable(Item,NULL,true);
		}
	}
	//NB: We don't free lpData because we have passed it off to the tree
}

void CHierarchyTableTreeCtrl::AddNode(LPSRow lpsRow, HTREEITEM hParent, BOOL bGetTable)
{
	if (!lpsRow) return;

	CString szName;
	LPSPropValue lpName = NULL;//don't free
	LPSPropValue lpEID = NULL;//don't free
	LPSPropValue lpInstance = NULL;//don't free
	LPSBinary lpEIDBin = NULL;//don't free
	LPSBinary lpInstanceBin = NULL;//don't free
	LPSPropValue lpSubfolders = NULL;//don't free
	LPSPropValue lpContainerFlags = NULL;//don't free

	lpName = PpropFindProp(
		lpsRow->lpProps,
		lpsRow->cValues,
		PR_DISPLAY_NAME);
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
	if (CheckStringProp(lpName,PT_TSTRING))
	{
		szName = lpName->Value.LPSZ;
	}
	else
	{
		szName.LoadString(IDS_UNKNOWNNAME);
	}
	lpSubfolders = PpropFindProp(
		lpsRow->lpProps,
		lpsRow->cValues,
		PR_SUBFOLDERS);
	lpContainerFlags = PpropFindProp(
		lpsRow->lpProps,
		lpsRow->cValues,
		PR_CONTAINER_FLAGS);
	AddNode(
		lpsRow->cValues,
		lpsRow->lpProps,//pass on props to be archived in node
		szName,
		lpEIDBin,
		lpInstanceBin,
		lpSubfolders?(ULONG)lpSubfolders->Value.b:MAPI_E_NOT_FOUND,
		lpContainerFlags?lpContainerFlags->Value.ul:MAPI_E_NOT_FOUND,
		hParent,
		bGetTable);
}

LPMAPITABLE CHierarchyTableTreeCtrl::GetHierarchyTable(HTREEITEM hItem,LPMAPICONTAINER lpMAPIContainer,BOOL bRegNotifs)
{
	HRESULT		hRes = S_OK;
	SortListData*	lpData = NULL;

	lpData = (SortListData*) GetItemData(hItem);

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
			//Get the hierarchy table for the node and shove it into the data
			LPMAPITABLE	lpHierarchyTable = NULL;

			//on the AB, something about this call triggers table reloads on the parent hierarchy table
			//no idea why they're triggered - doesn't happen for all AB providers
			WC_H(lpMAPIContainer->GetHierarchyTable(
				(m_ulDisplayFlags & dfDeleted?SHOW_SOFT_DELETES:NULL) | fMapiUnicode,
				&lpHierarchyTable));

			if (lpHierarchyTable)
			{
				enum {NAME,EID,INSTANCE,SUBFOLDERS,FLAGS,NUMCOLS};
				SizedSPropTagArray(NUMCOLS,sptHierarchyCols) = {NUMCOLS,
					PR_DISPLAY_NAME,
					PR_ENTRYID,
					PR_INSTANCE_KEY,
					PR_SUBFOLDERS,
					PR_CONTAINER_FLAGS};

				EC_H(lpHierarchyTable->SetColumns(
					(LPSPropTagArray) &sptHierarchyCols,
					TBL_BATCH));
			}

			lpData->data.Node.lpHierarchyTable = lpHierarchyTable;
			lpMAPIContainer->Release();
		}
	}

	if (lpData->data.Node.lpHierarchyTable && !lpData->data.Node.lpAdviseSink)
	{
		//set up our advise sink
		if (bRegNotifs &&
			(RegKeys[regkeyHIER_ROOT_NOTIFS].ulCurDWORD || GetRootItem() != hItem))
		{
			DebugPrintEx(DBGGeneric,CLASS,_T("GetHierarchyTable"),_T("Advise sink for \"%s\" = 0x%08X\n"),GetItemText(hItem),hItem);
			lpData->data.Node.lpAdviseSink = new CAdviseSink(m_hWnd,hItem);

			if (lpData->data.Node.lpAdviseSink)
			{
				WC_H(lpData->data.Node.lpHierarchyTable->Advise(
					fnevTableModified,
					(IMAPIAdviseSink *)lpData->data.Node.lpAdviseSink,
					&lpData->data.Node.ulAdviseConnection));
				if (MAPI_E_NO_SUPPORT == hRes)//Some tables don't support this!
				{
					if (lpData->data.Node.lpAdviseSink) lpData->data.Node.lpAdviseSink->Release();
					lpData->data.Node.lpAdviseSink = NULL;
					DebugPrint(DBGGeneric, _T("This table doesn't support notifications\n"));
					hRes = S_OK;//mask the error
				}
				else if (S_OK == hRes)
				{
					LPMDB lpMDB = m_lpMapiObjects->GetMDB();//do not release
					if (lpMDB)
					{
						LPSPropValue	lpProp = NULL;
						//Try to trigger some RPC to get the notifications going
						WC_H(HrGetOneProp(
							lpMDB,
							PR_TEST_LINE_SPEED,
							&lpProp));
						if (MAPI_E_NOT_FOUND == hRes)
						{
							//We're not on an Exchange server. We don't need to generate RPC after all.
							hRes = S_OK;
						}
						MAPIFreeBuffer(lpProp);
					}
				}
				DebugPrintEx(DBGGeneric,CLASS,_T("GetHierarchyTable"),_T("Advise sink 0x%08X, ulAdviseConnection = 0x%08X\n"),lpData->data.Node.lpAdviseSink,lpData->data.Node.ulAdviseConnection);
			}
		}
	}
	if (lpData->data.Node.lpAdviseSink)
	{
		SetItemState(hItem,TVIS_BOLD,TVIS_BOLD);
	}
	return lpData->data.Node.lpHierarchyTable;
}

//Add the first level contents of lpMAPIContainer under the Parent node
HRESULT CHierarchyTableTreeCtrl::ExpandNode(HTREEITEM hParent)
{
	LPMAPITABLE		lpHierarchyTable= NULL;
	LPSRowSet		pRows			= NULL;
	HRESULT			hRes			= S_OK;
	CWaitCursor		Wait;//Change the mouse to an hourglass while we work.

	if (!m_hWnd) return S_OK;
	if (!hParent) return MAPI_E_INVALID_PARAMETER;

	lpHierarchyTable = GetHierarchyTable(hParent,NULL,RegKeys[regkeyHIER_EXPAND_NOTIFS].ulCurDWORD);

	if (lpHierarchyTable)
	{
		//go to the first row
		EC_H(lpHierarchyTable->SeekRow(
			BOOKMARK_BEGINNING,
			0,
			NULL));
		hRes = S_OK;//don't let failure here fail the whole load

		ULONG i = 0;
		//get each row in turn and add it to the list
		if (!FAILED(hRes)) for (;;)
		{
			hRes = S_OK;
			//Note - we're saving the rows off in AddNode, so we don't FreeProws this...we just MAPIFreeBuffer the array
			if (pRows) MAPIFreeBuffer(pRows);
			pRows = NULL;
			EC_H(lpHierarchyTable->QueryRows(
				1,
				NULL,
				&pRows));
			if (FAILED(hRes) || !pRows || !pRows->cRows) break;
			//Now we can process the row!

			AddNode(
				pRows->aRow,
				hParent,
				(0 == RegKeys[regkeyHIER_NODE_LOAD_COUNT].ulCurDWORD) ||
				i < RegKeys[regkeyHIER_NODE_LOAD_COUNT].ulCurDWORD);
			i++;
		}
	}

	//Note - we're saving the props off in AddNode, so we don't FreeProws this...we just MAPIFreeBuffer the array
	if (pRows) MAPIFreeBuffer(pRows);
	return hRes;
}//CHierarchyTableTreeCtrl::ExpandNode

///////////////////////////////////////////////////////////////////////////////
//		 Message Handlers


void CHierarchyTableTreeCtrl::OnGetDispInfo(NMHDR* pNMHDR, LRESULT* pResult)
{
	NMTVDISPINFO* lpDispInfo = (LPNMTVDISPINFO) pNMHDR;

	if (lpDispInfo &&
		lpDispInfo->item.mask & TVIF_CHILDREN)
	{
		SortListData* lpData = NULL;
		lpData = (SortListData*) lpDispInfo->item.lParam;

		if (lpData)
		{
			//won't force the hierarchy table - just get it if we've already got it
			LPMAPITABLE lpHierarchyTable = lpData->data.Node.lpHierarchyTable;
			if (lpHierarchyTable)
			{
				lpDispInfo->item.cChildren = 1;
				HRESULT hRes = S_OK;
				ULONG ulRowCount = NULL;
				WC_H(lpHierarchyTable->GetRowCount(
					NULL,
					&ulRowCount));
				if (S_OK == hRes && !ulRowCount)
				{
					lpDispInfo->item.cChildren = 0;
				}
			}
			else
			{
				lpDispInfo->item.cChildren = 0;
				if ((lpData->data.Node.ulContainerFlags == MAPI_E_NOT_FOUND && lpData->data.Node.bSubfolders == MAPI_E_NOT_FOUND) ||
					(lpData->data.Node.ulContainerFlags != MAPI_E_NOT_FOUND && lpData->data.Node.ulContainerFlags & AB_SUBCONTAINERS) ||
					(lpData->data.Node.bSubfolders != MAPI_E_NOT_FOUND && lpData->data.Node.bSubfolders) ||
					m_ulDisplayFlags & dfDeleted)
				{
					lpDispInfo->item.cChildren = 1;
				}
			}
		}
	}
	*pResult = 0;
}

void CHierarchyTableTreeCtrl::UpdateSelectionUI(HTREEITEM hItem)
{
	HRESULT			hRes = S_OK;
	LPMAPICONTAINER	lpMAPIContainer = NULL;
	LPSPropValue	lpProps = NULL;
	ULONG			cVals = 0;
	ULONG			ulContCount = 0;
	ULONG			ulAssocContCount = 0;
	ULONG			ulDelMsgCount = 0;
	ULONG			ulDelAssocMsgCount = 0;
	ULONG			ulDelFolderCount = 0;

	enum{
		htPR_CONTENT_COUNT,
		htPR_ASSOC_CONTENT_COUNT,
		htPR_DELETED_FOLDER_COUNT,
		htPR_DELETED_MSG_COUNT,
		htPR_DELETED_ASSOC_MSG_COUNT,
		htNUMCOLS
	};
	SizedSPropTagArray(htNUMCOLS,sptHTCols) = {htNUMCOLS,
		PR_CONTENT_COUNT,
		PR_ASSOC_CONTENT_COUNT,
		PR_DELETED_FOLDER_COUNT,
		PR_DELETED_MSG_COUNT,
		PR_DELETED_ASSOC_MSG_COUNT
	};

	DebugPrintEx(DBGGeneric,CLASS,_T("UpdateSelectionUI"),_T("\n"));

	//Have to request modify or this object is read only in the single prop control.
	GetContainer(hItem,mfcmapiREQUEST_MODIFY, &lpMAPIContainer);

	//make sure we've gotten the hierarchy table for this node
	GetHierarchyTable(hItem,lpMAPIContainer,RegKeys[regkeyHIER_EXPAND_NOTIFS].ulCurDWORD);

	if (m_lpHostDlg && lpMAPIContainer)
	{
		// Get some props for status bar
		WC_H_GETPROPS(lpMAPIContainer->GetProps(
			(LPSPropTagArray) &sptHTCols,
			fMapiUnicode,
			&cVals,
			&lpProps));
		if (lpProps)
		{
			if (!(m_ulDisplayFlags & dfDeleted))
			{
				if (PT_ERROR == PROP_TYPE(lpProps[htPR_CONTENT_COUNT].ulPropTag))
				{
					WARNHRESMSG(lpProps[htPR_CONTENT_COUNT].Value.err,IDS_NODELACKSCONTENTCOUNT);
					ulContCount = 0;
				}
				else ulContCount = lpProps[htPR_CONTENT_COUNT].Value.ul;

				if (PT_ERROR == PROP_TYPE(lpProps[htPR_ASSOC_CONTENT_COUNT].ulPropTag))
				{
					WARNHRESMSG(lpProps[htPR_ASSOC_CONTENT_COUNT].Value.err,IDS_NODELACKSASSOCCONTENTCOUNT);
					ulAssocContCount = 0;
				}
				else ulAssocContCount = lpProps[htPR_ASSOC_CONTENT_COUNT].Value.ul;

				m_lpHostDlg->UpdateStatusBarText(
					STATUSLEFTPANE,
					IDS_STATUSTEXTCONTENTCOUNTS,
					ulContCount,
					ulAssocContCount);
			}
			else
			{
				if (PT_ERROR == PROP_TYPE(lpProps[htPR_DELETED_MSG_COUNT].ulPropTag))
				{
					WARNHRESMSG(lpProps[htPR_DELETED_MSG_COUNT].Value.err,IDS_NODELACKSDELETEDMESSAGECOUNT);
					ulDelMsgCount = 0;
				}
				else ulDelMsgCount = lpProps[htPR_DELETED_MSG_COUNT].Value.ul;

				if (PT_ERROR == PROP_TYPE(lpProps[htPR_DELETED_ASSOC_MSG_COUNT].ulPropTag))
				{
					WARNHRESMSG(lpProps[htPR_DELETED_ASSOC_MSG_COUNT].Value.err,IDS_NODELACKSDELETEDASSOCMESSAGECOUNT);
					ulDelAssocMsgCount = 0;
				}
				else ulDelAssocMsgCount = lpProps[htPR_DELETED_ASSOC_MSG_COUNT].Value.ul;

				if (PT_ERROR == PROP_TYPE(lpProps[htPR_DELETED_FOLDER_COUNT].ulPropTag))
				{
					WARNHRESMSG(lpProps[htPR_DELETED_FOLDER_COUNT].Value.err,IDS_NODELACKSDELETEDSUBFOLDERCOUNT);
					ulDelFolderCount = 0;
				}
				else ulDelFolderCount = lpProps[htPR_DELETED_FOLDER_COUNT].Value.ul;

				m_lpHostDlg->UpdateStatusBarText(
					STATUSLEFTPANE,
					IDS_STATUSTEXTDELETEDCOUNTS,
					ulDelMsgCount,
					ulDelAssocMsgCount,
					ulDelFolderCount);
			}
			MAPIFreeBuffer(lpProps);
		}

	}

	if (m_lpHostDlg)
	{
		m_lpHostDlg->OnUpdateSingleMAPIPropListCtrl(lpMAPIContainer, (SortListData*) GetItemData(hItem));
		m_lpHostDlg->UpdateTitleBarText(GetItemText(GetSelectedItem()));
	}

	if (lpMAPIContainer) lpMAPIContainer->Release();
}

void CHierarchyTableTreeCtrl::OnSelChanged(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMTREEVIEW pNMTV = (LPNMTREEVIEW) pNMHDR;

	if (pNMTV && pNMTV->itemNew.hItem)
	{
		m_bItemSelected = TRUE;

		UpdateSelectionUI(pNMTV->itemNew.hItem);
	}
	else
	{
		m_bItemSelected = FALSE;
	}

	*pResult = 0;
}

//This function will be called when we edit a node so we can attempt to commit the changes
void CHierarchyTableTreeCtrl::OnEndLabelEdit(NMHDR* pNMHDR, LRESULT* pResult)
{
	HRESULT			hRes = S_OK;
	LPMAPICONTAINER	lpMAPIContainer = NULL;
	SPropValue		sDisplayName;
	TV_DISPINFO*	pTVDispInfo = (TV_DISPINFO*)pNMHDR;

	*pResult = 0;

	if (!pTVDispInfo || !pTVDispInfo->item.pszText) return;

	GetContainer(pTVDispInfo->item.hItem,mfcmapiREQUEST_MODIFY, &lpMAPIContainer);
	if (!lpMAPIContainer) return;

	sDisplayName.ulPropTag = PR_DISPLAY_NAME;
	sDisplayName.Value.LPSZ = pTVDispInfo->item.pszText;

	EC_H(HrSetOneProp(lpMAPIContainer,&sDisplayName));

	lpMAPIContainer->Release();
}

void CHierarchyTableTreeCtrl::OnDblclk(NMHDR* /*pNMHDR*/, LRESULT* pResult)
{
	//Due to problems with focus...we have to post this instead of calling directly.
	//OnDisplayItem();
	//Post the message to display the item
	if (m_lpHostDlg)
		m_lpHostDlg->PostMessage(WM_COMMAND,ID_DISPLAYSELECTEDITEM,NULL);

	//Don't do default behavior for double-click (We only want '+' sign expansion.
	//Double click should display the item, not expand the tree.)
	*pResult = 1;
}

void CHierarchyTableTreeCtrl::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
{
	DebugPrintEx(DBGMenu,CLASS,_T("OnKeyDown"),_T("0x%X\n"),nChar);

	ULONG bCtrlPressed = GetKeyState(VK_CONTROL) <0;
	ULONG bShiftPressed = GetKeyState(VK_SHIFT) <0;
	ULONG bMenuPressed = GetKeyState(VK_MENU) <0;

	if (!bMenuPressed)
	{
		if (VK_RETURN == nChar && bCtrlPressed)
		{
			DebugPrintEx(DBGGeneric,CLASS,_T("OnKeyDown"),_T("calling Display Associated Contents\n"));
			if (m_lpHostDlg)
				m_lpHostDlg->PostMessage(WM_COMMAND,ID_DISPLAYASSOCIATEDCONTENTS,NULL);
		}
		else if (!m_lpHostDlg || !m_lpHostDlg->HandleKeyDown(nChar,bShiftPressed,bCtrlPressed,bMenuPressed))
		{
			CTreeCtrl::OnKeyDown(nChar,nRepCnt,nFlags);
		}
	}
}

//Assert that we want all keyboard input (including ENTER!)
UINT CHierarchyTableTreeCtrl::OnGetDlgCode()
{
	UINT iDlgCode = CTreeCtrl::OnGetDlgCode();
//	DebugPrintEx(DBGGeneric,CLASS,_T("OnGetDlgCode"),_T("default:0x%X\n"),iDlgCode);

	iDlgCode |= DLGC_WANTMESSAGE;

	// to make sure that the control key is not pressed
	if ((GetKeyState(VK_CONTROL) >= 0) && (m_hWnd == ::GetFocus()))
	{
		// to make sure that the Tab key is pressed
		if (GetKeyState(VK_TAB) < 0)
			iDlgCode &= ~(DLGC_WANTALLKEYS | DLGC_WANTMESSAGE |	DLGC_WANTTAB);
	}

//	DebugPrintEx(DBGGeneric,CLASS,_T("OnGetDlgCode"),_T("ret:0x%X\n"),iDlgCode);
	return iDlgCode;
}

void CHierarchyTableTreeCtrl::OnRightClick(NMHDR* /*pNMHDR*/, LRESULT* pResult)
{
	// Send WM_CONTEXTMENU to self
	(void)SendMessage(WM_CONTEXTMENU, (WPARAM) m_hWnd, GetMessagePos());

	// Mark message as handled and suppress default handling
	*pResult = 1;
}

void CHierarchyTableTreeCtrl::OnContextMenu(CWnd* /*pWnd*/, CPoint pos)
{
	HRESULT hRes = S_OK;

//	HTREEITEM htCurrentSelection = GetSelectedItem();

	// Select the item that is at the point myPoint.
	UINT uFlags;
    CPoint ptTree = pos;
    ScreenToClient(&ptTree);
	HTREEITEM hClickedItem = HitTest(ptTree, &uFlags);

	if ((hClickedItem != NULL) && (TVHT_ONITEM & uFlags))
	{
		Select(hClickedItem, TVGN_CARET);
	}

	if (m_lpHostDlg->m_nIDContextMenu)
	{
		EC_B(DisplayContextMenu(m_lpHostDlg->m_nIDContextMenu,IDR_MENU_HIERARCHY_TABLE,m_lpHostDlg,pos.x, pos.y));
	}
	else
	{
		EC_B(DisplayContextMenu(IDR_MENU_DEFAULT_POPUP,IDR_MENU_HIERARCHY_TABLE,m_lpHostDlg,pos.x, pos.y));
	}

	//Not doing this yet because when I do, the old selection becomes the target for my commands
	//Restore our selection
	//Select(htCurrentSelection, TVGN_CARET);
}

SortListData* CHierarchyTableTreeCtrl::GetSelectedItemData()
{
	HTREEITEM		Item = NULL;

	//Find the highlighted item
	Item = GetSelectedItem();

	if (Item)
	{
		return (SortListData*) GetItemData(Item);
	}
	return NULL;
} // CHierarchyTableTreeCtrl::GetSelectedItemData

LPSBinary CHierarchyTableTreeCtrl::GetSelectedItemEID()
{
	HTREEITEM		Item = NULL;

	//Find the highlighted item
	Item = GetSelectedItem();

	//get the EID associated with it
	if (Item)
	{
		SortListData* lpData = NULL;
		lpData = (SortListData*) GetItemData(Item);
		if (lpData)
			return lpData->data.Node.lpEntryID;
	}
	return NULL;
}

LPMAPICONTAINER CHierarchyTableTreeCtrl::GetSelectedContainer(__mfcmapiModifyEnum bModify)
{
	LPMAPICONTAINER lpSelectedContainer = NULL;

	GetContainer(GetSelectedItem(), bModify, &lpSelectedContainer);

	return lpSelectedContainer;
}

void CHierarchyTableTreeCtrl::GetContainer(
										   HTREEITEM Item,
										   __mfcmapiModifyEnum bModify,
										   LPMAPICONTAINER* lppContainer)
{
	HRESULT			hRes = S_OK;
	ULONG			ulObjType = NULL;
	ULONG			ulFlags = NULL;
	SortListData*	lpData = NULL;
	LPSBinary		lpCurBin = NULL;
	SBinary			NullBin = {0};
	LPMAPICONTAINER lpContainer = NULL;

	*lppContainer = NULL;

	if (!Item) return;

	DebugPrintEx(DBGGeneric,CLASS,_T("GetContainer"),_T("HTREEITEM = 0x%X, bModify = %d, m_ulContainerType = 0x%X\n"),Item, bModify, m_ulContainerType);

	lpData = (SortListData*) GetItemData(Item);

	if (!lpData)
	{
		//We didn't get an entryID, so log it and get out of here
		DebugPrintEx(DBGGeneric,CLASS,_T("GetContainer"),_T("GetItemData returned NULL or lpEntryID is NULL\n"));
		return;
	}

	ulFlags = (mfcmapiREQUEST_MODIFY == bModify?MAPI_MODIFY:NULL);

	lpCurBin = lpData->data.Node.lpEntryID;
	if (!lpCurBin) lpCurBin = &NullBin;

	//Check the type of the root container to know whether the MDB or AddrBook object is valid
	//This also allows NULL EID's to return the root container itself.
	//Use the Root container if we can't decide and log an error
	if (m_lpMapiObjects)
	{
		if (MAPI_ABCONT == m_ulContainerType)
		{
			LPADRBOOK lpAddrBook = m_lpMapiObjects->GetAddrBook(false);//do not release
			if (lpAddrBook)
			{
				DebugPrint(DBGGeneric,_T("\tCalling OpenEntry on address book with ulFlags = 0x%X\n"),ulFlags);

				WC_H(CallOpenEntry(
					NULL,
					lpAddrBook,
					NULL,
					NULL,
					lpCurBin->cb,
					(LPENTRYID) lpCurBin->lpb,
					NULL,
					ulFlags,
					&ulObjType,
					(LPUNKNOWN*)&lpContainer));
			}
		}
		else if (MAPI_FOLDER == m_ulContainerType)
		{
			LPMDB lpMDB = m_lpMapiObjects->GetMDB();//do not release
			if (lpMDB)
			{
				ulFlags = (mfcmapiREQUEST_MODIFY == bModify?MAPI_MODIFY:NULL) |
					(m_ulDisplayFlags & dfDeleted?(SHOW_SOFT_DELETES|MAPI_NO_CACHE):NULL);

				WC_H(CallOpenEntry(
					lpMDB,
					NULL,
					NULL,
					NULL,
					lpCurBin->cb,
					(LPENTRYID) lpCurBin->lpb,
					NULL,
					ulFlags,
					&ulObjType,
					(LPUNKNOWN*)&lpContainer));
			}
		}
	}

	//If we didn't get a container above, try to open from m_lpContainer
	if (!lpContainer && m_lpContainer)
	{
		WARNHRESMSG(MAPI_E_CALL_FAILED,IDS_UNKNOWNCONTAINERTYPE);
		hRes = S_OK;
		WC_H(CallOpenEntry(
			NULL,
			NULL,
			m_lpContainer,
			NULL,
			lpCurBin->cb,
			(LPENTRYID) lpCurBin->lpb,
			NULL,
			ulFlags,
			&ulObjType,
			(LPUNKNOWN*)&lpContainer));
	}

	//if we failed because write access was denied, try again if acceptable
	if (!lpContainer && FAILED(hRes) && mfcmapiREQUEST_MODIFY == bModify)
	{
		DebugPrint(DBGGeneric,_T("\tOpenEntry failed: 0x%X. Will try again without MAPI_MODIFY\n"),hRes);
		//We failed to open the item with MAPI_MODIFY.
		//Let's try to open it with NULL
		GetContainer(
			Item,
			mfcmapiDO_NOT_REQUEST_MODIFY,
			&lpContainer);
	}

	//Ok - we're just out of luck
	if (!lpContainer)
	{
		WARNHRESMSG(hRes,IDS_NOCONTAINER);
		hRes = MAPI_E_NOT_FOUND;
	}

	if (lpContainer) *lppContainer = lpContainer;
	DebugPrintEx(DBGGeneric,CLASS,_T("GetContainer"),_T("returning lpContainer = 0x%X, ulObjType = 0x%X and hRes = 0x%X\n"),lpContainer, ulObjType, hRes);
	return;
}


///////////////////////////////////////////////////////////////////////////////
//		 End - Message Handlers


//When + is clicked, add all entries in the table as children
void CHierarchyTableTreeCtrl::OnItemExpanding(NMHDR* pNMHDR, LRESULT* pResult)
{
	HRESULT			hRes = S_OK;

	*pResult = 0;

	NM_TREEVIEW* pNMTreeView = (NM_TREEVIEW*)pNMHDR;
	if (pNMTreeView)
	{
		if (pNMTreeView->action & TVE_EXPAND)
		{
			if (!(pNMTreeView->itemNew.state & TVIS_EXPANDEDONCE))
			{
				EC_H(ExpandNode(pNMTreeView->itemNew.hItem));
			}
		}
	}
}

//Tree control will call this for every node it deletes.
void CHierarchyTableTreeCtrl::OnDeleteItem(NMHDR* pNMHDR, LRESULT* pResult)
{
	LPNMTREEVIEW pNMTreeView = (LPNMTREEVIEW)pNMHDR;
	if (pNMTreeView)
	{
		SortListData* lpData = (SortListData*) pNMTreeView->itemOld.lParam;

		if (lpData && lpData->data.Node.lpAdviseSink)
		{
			DebugPrintEx(DBGGeneric,CLASS,_T("OnDeleteItem"),_T("Unadvising on \"%s\", 0x%08X, ulAdviseConnection = 0x%08X\n"),GetItemText(pNMTreeView->itemOld.hItem),lpData->data.Node.lpAdviseSink,lpData->data.Node.ulAdviseConnection);
		}
		FreeSortListData(lpData);
	}

	*pResult = 0;
}

//WM_MFCMAPI_ADDITEM
//If the parent has been expanded once, add the new row
//Otherwise, ditch the notification
LRESULT	CHierarchyTableTreeCtrl::msgOnAddItem(WPARAM wParam, LPARAM lParam)
{
	TABLE_NOTIFICATION* tab = (TABLE_NOTIFICATION*) wParam;
	HTREEITEM			hParent = (HTREEITEM) lParam;

	DebugPrintEx(DBGGeneric,CLASS,_T("msgOnAddItem"),_T("Received message add item under: 0x%08X =\"%s\"\n"),hParent,GetItemText(hParent));

	//only need to add the node if we're expanded
	int iState = GetItemState(hParent,NULL);
	if (iState & TVIS_EXPANDEDONCE)
	{
		HRESULT hRes = S_OK;
		//We make this copy here and pass it in to AddNode, where it is grabbed by BuildDataItem to be part of the item data
		//The mem will be freed when the item data is cleaned up - do not free here
		SRow NewRow = {0};
		NewRow.cValues = tab->row.cValues;
		NewRow.ulAdrEntryPad = tab->row.ulAdrEntryPad;
		EC_H(ScDupPropset(
			tab->row.cValues,
			tab->row.lpProps,
			MAPIAllocateBuffer,
			&NewRow.lpProps));
		AddNode(&NewRow,hParent,true);
	}

	return S_OK;
}//CHierarchyTableTreeCtrl::msgOnAddItem

//WM_MFCMAPI_DELETEITEM
//Remove the child node.
//TODO: Check if the parent now has no children?
LRESULT	CHierarchyTableTreeCtrl::msgOnDeleteItem(WPARAM wParam, LPARAM lParam)
{
	HRESULT				hRes = S_OK;
	TABLE_NOTIFICATION*	tab = (TABLE_NOTIFICATION*) wParam;
	HTREEITEM			hParent = (HTREEITEM) lParam;

	HTREEITEM			hItemToDelete = FindNode(
		&tab->propIndex.Value.bin,
		hParent);

	if (hItemToDelete)
	{
		DebugPrintEx(DBGGeneric,CLASS,_T("msgOnDeleteItem"),_T("Received message delete item: 0x%08X =\"%s\"\n"),hItemToDelete,GetItemText(hItemToDelete));
		EC_B(DeleteItem(hItemToDelete));
	}

	return hRes;
}

//WM_MFCMAPI_MODIFYITEM
//Update any UI for the node and resort if needed.
LRESULT	CHierarchyTableTreeCtrl::msgOnModifyItem(WPARAM wParam, LPARAM lParam)
{
	HRESULT				hRes = S_OK;
	TABLE_NOTIFICATION*	tab = (TABLE_NOTIFICATION*) wParam;
	HTREEITEM			hParent = (HTREEITEM) lParam;

	HTREEITEM			hModifyItem = FindNode(
		&tab->propIndex.Value.bin,
		hParent);

	if (hModifyItem)
	{
		DebugPrintEx(DBGGeneric,CLASS,_T("msgOnModifyItem"),_T("Received message modify item: 0x%08X =\"%s\"\n"),hModifyItem,GetItemText(hModifyItem));

		LPSPropValue lpName = NULL;//don't free
		lpName = PpropFindProp(
			tab->row.lpProps,
			tab->row.cValues,
			PR_DISPLAY_NAME);

		if (CheckStringProp(lpName,PT_TSTRING))
		{
			EC_B(SetItemText(hModifyItem,lpName->Value.LPSZ));
		}
		else
		{
			CString szText;
			szText.LoadString(IDS_UNKNOWNNAME);
			EC_B(SetItemText(hModifyItem,szText));
		}

		if (hParent) EC_B(SortChildren(hParent));
	}

	if (hModifyItem == GetSelectedItem()) UpdateSelectionUI(hModifyItem);

	return hRes;
}

//WM_MFCMAPI_REFRESHTABLE
//If node was expanded, collapse it to remove all children
//Then, if the node does have children, reexpand it.
LRESULT	CHierarchyTableTreeCtrl::msgOnRefreshTable(WPARAM wParam, LPARAM /*lParam*/)
{
	HRESULT		hRes = S_OK;
	HTREEITEM	hRefreshItem = (HTREEITEM) wParam;
	DebugPrintEx(DBGGeneric,CLASS,_T("msgOnRefreshTable"),_T("Received message refresh table: 0x%08X =\"%s\"\n"),hRefreshItem,GetItemText(hRefreshItem));

	int iState = GetItemState(hRefreshItem,NULL);
	if (iState & TVIS_EXPANDED)
	{
		//can't collapse when using I_CHILDRENCALLBACK
//		EC_B(Expand(hRefreshItem,TVE_COLLAPSE | TVE_COLLAPSERESET));
		HTREEITEM hChild = GetChildItem(hRefreshItem);
		while (hChild)
		{
			hRes = S_OK;
			EC_B(DeleteItem(hChild));
			hChild = GetChildItem(hRefreshItem);
		}
		//Reset our expanded bits
		EC_B(SetItemState(hRefreshItem,NULL,TVIS_EXPANDED | TVIS_EXPANDEDONCE));
		hRes = S_OK;
		{
			SortListData* lpData = NULL;
			lpData = (SortListData*) GetItemData(hRefreshItem);

			if (lpData)
			{
				if (lpData->data.Node.lpHierarchyTable)
				{
					ULONG ulRowCount = NULL;
					WC_H(lpData->data.Node.lpHierarchyTable->GetRowCount(
						NULL,
						&ulRowCount));
					if (S_OK != hRes || ulRowCount)
					{
						EC_B(Expand(hRefreshItem,TVE_EXPAND));
					}
				}
			}
		}

	}

	return hRes;
}

//This function steps through the list control to find the entry with this instance key
//return NULL if item not found
HTREEITEM CHierarchyTableTreeCtrl::FindNode(LPSBinary lpInstance, HTREEITEM hParent)
{
	if (!lpInstance || !hParent) return NULL;

	DebugPrintEx(DBGGeneric,CLASS,_T("FindNode"),_T("Looking for child of: 0x%08X =\"%s\"\n"),hParent,GetItemText(hParent));

	LPSBinary	lpCurInstance = NULL;
	SortListData*	lpListData = NULL;
	HTREEITEM	hCurrent = NULL;

	hCurrent = GetNextItem(hParent, TVGN_CHILD);

	while (hCurrent)
	{
		lpListData = (SortListData*) GetItemData(hCurrent);
		if (lpListData)
		{
			lpCurInstance = lpListData->data.Node.lpInstanceKey;
			if (lpCurInstance)
			{
				if (!memcmp(lpCurInstance->lpb, lpInstance->lpb, lpInstance->cb))
				{
					DebugPrintEx(DBGGeneric,CLASS,_T("FindNode"),_T("Matched at 0x%08X =\"%s\"\n"),hCurrent,GetItemText(hCurrent));
					return hCurrent;
				}
			}
		}
		hCurrent = GetNextItem(hCurrent,TVGN_NEXT);
	}

	DebugPrintEx(DBGGeneric,CLASS,_T("FindNode"),_T("No match found\n"));
	return NULL;
}
