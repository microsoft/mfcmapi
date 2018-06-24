#include <StdAfx.h>
#include <UI/Controls/HierarchyTableTreeCtrl.h>
#include <UI/Dialogs/BaseDialog.h>
#include <UI/Dialogs/HierarchyTable/HierarchyTableDlg.h>
#include <MAPI/Cache/MapiObjects.h>
#include <MAPI/MAPIFunctions.h>
#include <UI/UIFunctions.h>
#include <MAPI/AdviseSink.h>
#include <UI/Controls/SortList/NodeData.h>

namespace controls
{
	enum
	{
		htPR_DISPLAY_NAME_W,
		htPR_ENTRYID,
		htPR_INSTANCE_KEY,
		htPR_SUBFOLDERS,
		htPR_CONTAINER_FLAGS,
		htNUMCOLS
	};

	static const SizedSPropTagArray(htNUMCOLS, sptHTCols) = {
		htNUMCOLS,
		{PR_DISPLAY_NAME_W, PR_ENTRYID, PR_INSTANCE_KEY, PR_SUBFOLDERS, PR_CONTAINER_FLAGS}};

	enum
	{
		htcPR_CONTENT_COUNT,
		htcPR_ASSOC_CONTENT_COUNT,
		htcPR_DELETED_FOLDER_COUNT,
		htcPR_DELETED_MSG_COUNT,
		htcPR_DELETED_ASSOC_MSG_COUNT,
		htcNUMCOLS
	};
	static const SizedSPropTagArray(htNUMCOLS, sptHTCountCols) = {htcNUMCOLS,
																  {PR_CONTENT_COUNT,
																   PR_ASSOC_CONTENT_COUNT,
																   PR_DELETED_FOLDER_COUNT,
																   PR_DELETED_MSG_COUNT,
																   PR_DELETED_ASSOC_MSG_COUNT}};

	static std::wstring CLASS = L"CHierarchyTableTreeCtrl";

	CHierarchyTableTreeCtrl::CHierarchyTableTreeCtrl(
		_In_ CWnd* pCreateParent,
		_In_ cache::CMapiObjects* lpMapiObjects,
		_In_ dialog::CHierarchyTableDlg* lpHostDlg,
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

		m_lpContainer = nullptr;
		m_ulContainerType = NULL;

		m_ulDisplayFlags = ulDisplayFlags;

		m_bItemSelected = false;

		m_nIDContextMenu = nIDContextMenu;

		m_hItemCurHover = nullptr;
		m_HoverButton = false;

		Create(
			TVS_HASBUTTONS | TVS_LINESATROOT | TVS_EDITLABELS | TVS_DISABLEDRAGDROP | TVS_SHOWSELALWAYS |
				TVS_FULLROWSELECT | WS_CHILD |
				WS_TABSTOP
				//| WS_CLIPCHILDREN
				| WS_CLIPSIBLINGS | WS_VISIBLE,
			pRect,
			pCreateParent,
			IDC_FOLDER_TREE);
		TreeView_SetBkColor(m_hWnd, ui::MyGetSysColor(ui::cBackground));
		TreeView_SetTextColor(m_hWnd, ui::MyGetSysColor(ui::cText));
		::SendMessageA(m_hWnd, WM_SETFONT, reinterpret_cast<WPARAM>(ui::GetSegoeFont()), false);
	}

	CHierarchyTableTreeCtrl::~CHierarchyTableTreeCtrl()
	{
		TRACE_DESTRUCTOR(CLASS);
		m_bShuttingDown = true;
		CWnd::DestroyWindow();

		if (m_lpContainer) m_lpContainer->Release();
		if (m_lpHostDlg) m_lpHostDlg->Release();
		if (m_lpMapiObjects) m_lpMapiObjects->Release();
	}

	STDMETHODIMP_(ULONG) CHierarchyTableTreeCtrl::AddRef()
	{
		const auto lCount = InterlockedIncrement(&m_cRef);
		TRACE_ADDREF(CLASS, lCount);
		return lCount;
	}

	STDMETHODIMP_(ULONG) CHierarchyTableTreeCtrl::Release()
	{
		const auto lCount = InterlockedDecrement(&m_cRef);
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
		const auto hItemCurHover = m_hItemCurHover;
		switch (message)
		{
		case WM_MOUSEMOVE:
		{
			auto hRes = S_OK;

			TVHITTESTINFO tvHitTestInfo = {0};
			tvHitTestInfo.pt.x = GET_X_LPARAM(lParam);
			tvHitTestInfo.pt.y = GET_Y_LPARAM(lParam);

			WC_B(::SendMessage(m_hWnd, TVM_HITTEST, 0, reinterpret_cast<LPARAM>(&tvHitTestInfo)));
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

					WC_B(TrackMouseEvent(&tmEvent));
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

	_Check_return_ HRESULT CHierarchyTableTreeCtrl::RefreshHierarchyTable()
	{
		auto hRes = S_OK;

		// Turn off redraw while we work on the window
		SetRedraw(false);

		m_bItemSelected = false; // clear this just in case

		EC_B(DeleteItem(GetRootItem()));

		if (m_lpContainer) EC_H(AddRootNode(m_lpContainer));

		if (m_lpHostDlg) m_lpHostDlg->OnUpdateSingleMAPIPropListCtrl(nullptr, nullptr);

		// Turn redraw back on to update our view
		SetRedraw(true);
		return hRes;
	}

	_Check_return_ HRESULT CHierarchyTableTreeCtrl::LoadHierarchyTable(_In_ LPMAPICONTAINER lpMAPIContainer)
	{
		auto hRes = S_OK;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (m_lpContainer) m_lpContainer->Release();
		m_lpContainer = lpMAPIContainer;
		m_ulContainerType = NULL;

		// If we weren't passed a container to load, give up
		if (!m_lpContainer) return hRes;
		m_lpContainer->AddRef();

		m_ulContainerType = mapi::GetMAPIObjectType(lpMAPIContainer);

		WC_H(RefreshHierarchyTable());
		if (MAPI_E_NOT_FOUND == hRes)
		{
			WARNHRESMSG(hRes, IDS_HIERARCHNOTFOUND);
		}
		else
			CHECKHRESMSG(hRes, IDS_REFRESHHIERARCHYFAILED);

		return hRes;
	}

	// Removes any existing node data and replaces it with lpData
	void SetNodeData(HWND hWnd, HTREEITEM hItem, sortlistdata::SortListData* lpData)
	{
		if (lpData)
		{
			TVITEM tvItem = {0};
			tvItem.hItem = hItem;
			tvItem.mask = TVIF_PARAM;
			if (TreeView_GetItem(hWnd, &tvItem) && tvItem.lParam)
			{
				output::DebugPrintEx(DBGHierarchy, CLASS, L"SetNodeData", L"Node %p, replacing data\n", hItem);
				delete reinterpret_cast<sortlistdata::SortListData*>(tvItem.lParam);
			}
			else
			{
				output::DebugPrintEx(DBGHierarchy, CLASS, L"SetNodeData", L"Node %p, first data\n", hItem);
			}

			tvItem.lParam = reinterpret_cast<LPARAM>(lpData);
			TreeView_SetItem(hWnd, &tvItem);
			// The tree now owns our lpData
		}
	}

	_Check_return_ HRESULT CHierarchyTableTreeCtrl::AddRootNode(_In_ LPMAPICONTAINER lpMAPIContainer) const
	{
		auto hRes = S_OK;
		LPSPropValue lpProps = nullptr;
		LPSPropValue lpRootName = nullptr; // don't free
		LPSBinary lpEIDBin = nullptr; // don't free
		if (!m_hWnd) return S_OK;
		if (!lpMAPIContainer) return MAPI_E_INVALID_PARAMETER;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		ULONG cVals = 0;

		WC_H_GETPROPS(lpMAPIContainer->GetProps(LPSPropTagArray(&sptHTCols), fMapiUnicode, &cVals, &lpProps));
		hRes = S_OK;

		// Get the entry ID for the Root Container
		if (!lpProps || PT_ERROR == PROP_TYPE(lpProps[htPR_ENTRYID].ulPropTag))
		{
			output::DebugPrint(
				DBGHierarchy, L"Could not find EntryID for Root Container. This is benign. Assuming NULL.\n");
		}
		else
			lpEIDBin = &lpProps[htPR_ENTRYID].Value.bin;

		// Get the Display Name for the Root Container
		if (!lpProps || PT_ERROR == PROP_TYPE(lpProps[htPR_DISPLAY_NAME_W].ulPropTag))
		{
			output::DebugPrint(
				DBGHierarchy, L"Could not find Display Name for Root Container. This is benign. Assuming NULL.\n");
		}
		else
			lpRootName = &lpProps[htPR_DISPLAY_NAME_W];

		std::wstring szName;

		// Shouldn't have to check lpRootName for non-NULL since CheckString does it, but prefast is complaining
		if (lpRootName && mapi::CheckStringProp(lpRootName, PT_UNICODE))
		{
			szName = lpRootName->Value.lpszW;
		}
		else
		{
			szName = strings::loadstring(IDS_ROOTCONTAINER);
		}

		auto lpData = new sortlistdata::SortListData();
		if (lpData)
		{
			lpData->InitializeNode(
				cVals,
				lpProps, // pass our lpProps to be archived
				lpEIDBin,
				nullptr,
				true, // Always assume root nodes have children so we always paint an expanding icon
				lpProps ? lpProps[htPR_CONTAINER_FLAGS].Value.ul : MAPI_E_NOT_FOUND);

			AddNode(szName, TVI_ROOT, lpData, true);
		}

		// Node owns the lpProps memory now, so we don't free it
		return hRes;
	}

	void CHierarchyTableTreeCtrl::AddNode(
		_In_ const std::wstring& szName,
		HTREEITEM hParent,
		sortlistdata::SortListData* lpData,
		bool bGetTable) const
	{
		output::DebugPrintEx(
			DBGHierarchy,
			CLASS,
			L"AddNode",
			L"Adding Node \"%ws\" under node %p, bGetTable = 0x%X\n",
			szName.c_str(),
			hParent,
			bGetTable);
		TVINSERTSTRUCTW tvInsert = {nullptr};

		tvInsert.hParent = hParent;
		tvInsert.hInsertAfter = TVI_SORT;
		tvInsert.item.mask = TVIF_CHILDREN | TVIF_TEXT;
		tvInsert.item.cChildren = I_CHILDRENCALLBACK;
		tvInsert.item.pszText = const_cast<LPWSTR>(szName.c_str());
		const auto hItem =
			reinterpret_cast<HTREEITEM>(::SendMessage(m_hWnd, TVM_INSERTITEMW, 0, reinterpret_cast<LPARAM>(&tvInsert)));

		SetNodeData(m_hWnd, hItem, lpData);

		if (bGetTable && (registry::RegKeys[registry::regkeyHIER_ROOT_NOTIFS].ulCurDWORD || hParent != TVI_ROOT))
		{
			(void) GetHierarchyTable(hItem, nullptr, true);
		}
	}

	void CHierarchyTableTreeCtrl::AddNode(_In_ LPSRow lpsRow, HTREEITEM hParent, bool bGetTable) const
	{
		if (!lpsRow) return;

		std::wstring szName;
		const auto lpName = PpropFindProp(lpsRow->lpProps, lpsRow->cValues, PR_DISPLAY_NAME_W);
		if (mapi::CheckStringProp(lpName, PT_UNICODE))
		{
			szName = lpName->Value.lpszW;
		}
		else
		{
			szName = strings::loadstring(IDS_UNKNOWNNAME);
		}
		output::DebugPrintEx(DBGHierarchy, CLASS, L"AddNode", L"Adding to %p: %ws\n", hParent, szName.c_str());

		auto lpData = new sortlistdata::SortListData();
		if (lpData)
		{
			lpData->InitializeNode(lpsRow);

			AddNode(szName, hParent, lpData, bGetTable);
		}
	}

	_Check_return_ LPMAPITABLE CHierarchyTableTreeCtrl::GetHierarchyTable(
		HTREEITEM hItem,
		_In_opt_ LPMAPICONTAINER lpMAPIContainer,
		bool bRegNotifs) const
	{
		auto hRes = S_OK;
		const auto lpData = GetSortListData(hItem);

		if (!lpData || !lpData->Node()) return nullptr;

		if (!lpData->Node()->m_lpHierarchyTable)
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
				LPMAPITABLE lpHierarchyTable = nullptr;

				// on the AB, something about this call triggers table reloads on the parent hierarchy table
				// no idea why they're triggered - doesn't happen for all AB providers
				WC_MAPI(lpMAPIContainer->GetHierarchyTable(
					(m_ulDisplayFlags & dfDeleted ? SHOW_SOFT_DELETES : NULL) | fMapiUnicode, &lpHierarchyTable));

				if (lpHierarchyTable)
				{
					EC_MAPI(lpHierarchyTable->SetColumns(LPSPropTagArray(&sptHTCols), TBL_BATCH));
				}

				lpData->Node()->m_lpHierarchyTable = lpHierarchyTable;
				lpMAPIContainer->Release();
			}
		}

		if (lpData->Node()->m_lpHierarchyTable && !lpData->Node()->m_lpAdviseSink)
		{
			// set up our advise sink
			if (bRegNotifs &&
				(registry::RegKeys[registry::regkeyHIER_ROOT_NOTIFS].ulCurDWORD || GetRootItem() != hItem))
			{
				output::DebugPrintEx(
					DBGNotify,
					CLASS,
					L"GetHierarchyTable",
					L"Advise sink for \"%ws\" = %p\n",
					strings::LPCTSTRToWstring(GetItemText(hItem)).c_str(),
					hItem);
				lpData->Node()->m_lpAdviseSink = new mapi::mapiui::CAdviseSink(m_hWnd, hItem);

				if (lpData->Node()->m_lpAdviseSink)
				{
					WC_MAPI(lpData->Node()->m_lpHierarchyTable->Advise(
						fnevTableModified,
						static_cast<IMAPIAdviseSink*>(lpData->Node()->m_lpAdviseSink),
						&lpData->Node()->m_ulAdviseConnection));
					if (MAPI_E_NO_SUPPORT == hRes) // Some tables don't support this!
					{
						if (lpData->Node()->m_lpAdviseSink) lpData->Node()->m_lpAdviseSink->Release();
						lpData->Node()->m_lpAdviseSink = nullptr;
						output::DebugPrint(DBGNotify, L"This table doesn't support notifications\n");
					}
					else if (hRes == S_OK)
					{
						const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
						if (lpMDB)
						{
							lpData->Node()->m_lpAdviseSink->SetAdviseTarget(lpMDB);
							mapi::ForceRop(lpMDB);
						}
					}

					output::DebugPrintEx(
						DBGNotify,
						CLASS,
						L"GetHierarchyTable",
						L"Advise sink %p, ulAdviseConnection = 0x%08X\n",
						lpData->Node()->m_lpAdviseSink,
						static_cast<int>(lpData->Node()->m_ulAdviseConnection));
				}
			}
		}

		return lpData->Node()->m_lpHierarchyTable;
	}

	// Add the first level contents of lpMAPIContainer under the Parent node
	_Check_return_ HRESULT CHierarchyTableTreeCtrl::ExpandNode(HTREEITEM hParent) const
	{
		LPSRowSet pRows = nullptr;
		auto hRes = S_OK;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_hWnd) return S_OK;
		if (!hParent) return MAPI_E_INVALID_PARAMETER;
		output::DebugPrintEx(DBGHierarchy, CLASS, L"ExpandNode", L"Expanding %p\n", hParent);

		auto lpHierarchyTable =
			GetHierarchyTable(hParent, nullptr, 0 != registry::RegKeys[registry::regkeyHIER_EXPAND_NOTIFS].ulCurDWORD);
		if (lpHierarchyTable)
		{
			// go to the first row
			EC_MAPI(lpHierarchyTable->SeekRow(BOOKMARK_BEGINNING, 0, nullptr));
			hRes = S_OK; // don't let failure here fail the whole load

			ULONG i = 0;
			// get each row in turn and add it to the list
			// TODO: Query several rows at once
			if (!FAILED(hRes))
				for (;;)
				{
					hRes = S_OK;
					// Note - we're saving the rows off in AddNode, so we don't FreeProws this...we just MAPIFreeBuffer the array
					if (pRows) MAPIFreeBuffer(pRows);
					pRows = nullptr;
					EC_MAPI(lpHierarchyTable->QueryRows(1, NULL, &pRows));
					if (FAILED(hRes) || !pRows || !pRows->cRows) break;
					// Now we can process the row!

					AddNode(pRows->aRow, hParent, false);
					i++;
				}
		}

		// Note - we're saving the props off in AddNode, so we don't FreeProws this...we just MAPIFreeBuffer the array
		if (pRows) MAPIFreeBuffer(pRows);
		return hRes;
	}

	void CHierarchyTableTreeCtrl::OnGetDispInfo(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
	{
		const auto lpDispInfo = reinterpret_cast<LPNMTVDISPINFO>(pNMHDR);

		if (lpDispInfo && lpDispInfo->item.mask & TVIF_CHILDREN)
		{
			const auto lpData = reinterpret_cast<sortlistdata::SortListData*>(lpDispInfo->item.lParam);

			if (lpData && lpData->Node())
			{
				if (m_ulDisplayFlags & dfDeleted)
				{
					lpDispInfo->item.cChildren = 1;
				}
				else if (lpData->Node()->m_cSubfolders >= 0)
				{
					lpDispInfo->item.cChildren = lpData->Node()->m_cSubfolders ? 1 : 0;
				}
				else
				{
					LPCTSTR szName = nullptr;
					if (PROP_TYPE(lpData->lpSourceProps[0].ulPropTag) == PT_TSTRING)
						szName = lpData->lpSourceProps[0].Value.LPSZ;
					output::DebugPrintEx(
						DBGHierarchy,
						CLASS,
						L"OnGetDispInfo",
						L"Using Hierarchy table %d %p %ws\n",
						lpData->Node()->m_cSubfolders,
						lpData->Node()->m_lpHierarchyTable,
						strings::LPCTSTRToWstring(szName).c_str());
					// won't force the hierarchy table - just get it if we've already got it
					auto lpHierarchyTable = lpData->Node()->m_lpHierarchyTable;
					if (lpHierarchyTable)
					{
						lpDispInfo->item.cChildren = 1;
						auto hRes = S_OK;
						ULONG ulRowCount = NULL;
						WC_MAPI(lpHierarchyTable->GetRowCount(NULL, &ulRowCount));
						if (hRes == S_OK && !ulRowCount)
						{
							lpDispInfo->item.cChildren = 0;
						}
					}
				}
			}
		}
		*pResult = 0;
	}

	void CHierarchyTableTreeCtrl::UpdateSelectionUI(HTREEITEM hItem) const
	{
		auto hRes = S_OK;
		LPMAPICONTAINER lpMAPIContainer = nullptr;
		LPSPropValue lpProps = nullptr;
		ULONG cVals = 0;
		UINT uiMsg = IDS_STATUSTEXTNOFOLDER;
		std::wstring szParam1;
		std::wstring szParam2;
		std::wstring szParam3;

		output::DebugPrintEx(DBGHierarchy, CLASS, L"UpdateSelectionUI", L"%p\n", hItem);

		if (!m_lpHostDlg) return;

		// Have to request modify or this object is read only in the single prop control.
		GetContainer(hItem, mfcmapiREQUEST_MODIFY, &lpMAPIContainer);

		// make sure we've gotten the hierarchy table for this node
		(void) GetHierarchyTable(
			hItem, lpMAPIContainer, 0 != registry::RegKeys[registry::regkeyHIER_EXPAND_NOTIFS].ulCurDWORD);

		if (SUCCEEDED(hRes) && lpMAPIContainer)
		{
			// Get some props for status bar
			WC_H_GETPROPS(lpMAPIContainer->GetProps(LPSPropTagArray(&sptHTCountCols), fMapiUnicode, &cVals, &lpProps));
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
					else
						szParam1 = std::to_wstring(lpProps[htcPR_CONTENT_COUNT].Value.ul);

					if (PT_ERROR == PROP_TYPE(lpProps[htcPR_ASSOC_CONTENT_COUNT].ulPropTag))
					{
						WARNHRESMSG(lpProps[htcPR_ASSOC_CONTENT_COUNT].Value.err, IDS_NODELACKSASSOCCONTENTCOUNT);
						szParam2 = L"?"; // STRING_OK
					}
					else
						szParam2 = std::to_wstring(lpProps[htcPR_ASSOC_CONTENT_COUNT].Value.ul);
				}
				else
				{
					uiMsg = IDS_STATUSTEXTDELETEDCOUNTS;
					if (PT_ERROR == PROP_TYPE(lpProps[htcPR_DELETED_MSG_COUNT].ulPropTag))
					{
						WARNHRESMSG(lpProps[htcPR_DELETED_MSG_COUNT].Value.err, IDS_NODELACKSDELETEDMESSAGECOUNT);
						szParam1 = L"?"; // STRING_OK
					}
					else
						szParam1 = std::to_wstring(lpProps[htcPR_DELETED_MSG_COUNT].Value.ul);

					if (PT_ERROR == PROP_TYPE(lpProps[htcPR_DELETED_ASSOC_MSG_COUNT].ulPropTag))
					{
						WARNHRESMSG(
							lpProps[htcPR_DELETED_ASSOC_MSG_COUNT].Value.err, IDS_NODELACKSDELETEDASSOCMESSAGECOUNT);
						szParam2 = L"?"; // STRING_OK
					}
					else
						szParam2 = std::to_wstring(lpProps[htcPR_DELETED_ASSOC_MSG_COUNT].Value.ul);

					if (PT_ERROR == PROP_TYPE(lpProps[htcPR_DELETED_FOLDER_COUNT].ulPropTag))
					{
						WARNHRESMSG(lpProps[htcPR_DELETED_FOLDER_COUNT].Value.err, IDS_NODELACKSDELETEDSUBFOLDERCOUNT);
						szParam3 = L"?"; // STRING_OK
					}
					else
						szParam3 = std::to_wstring(lpProps[htcPR_DELETED_FOLDER_COUNT].Value.ul);
				}
				MAPIFreeBuffer(lpProps);
			}
		}

		m_lpHostDlg->UpdateStatusBarText(STATUSDATA1, uiMsg, szParam1, szParam2, szParam3);

		m_lpHostDlg->OnUpdateSingleMAPIPropListCtrl(lpMAPIContainer, GetSortListData(hItem));

		WCHAR szText[255] = {0};
		TVITEMEXW item = {0};
		item.mask = TVIF_TEXT;
		item.pszText = szText;
		item.cchTextMax = _countof(szText);
		item.hItem = GetSelectedItem();
		WC_B(::SendMessage(m_hWnd, TVM_GETITEMW, 0, reinterpret_cast<LPARAM>(&item)));
		m_lpHostDlg->UpdateTitleBarText(szText);

		if (lpMAPIContainer) lpMAPIContainer->Release();
	}

	_Check_return_ bool CHierarchyTableTreeCtrl::IsItemSelected() const { return m_bItemSelected; }

	void CHierarchyTableTreeCtrl::OnSelChanged(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
	{
		const auto pNMTV = reinterpret_cast<LPNMTREEVIEW>(pNMHDR);

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
	// TODO: In non-unicode builds, this gives us ANSI strings - need to figure out how to change that
	void CHierarchyTableTreeCtrl::OnEndLabelEdit(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
	{
		auto hRes = S_OK;
		const auto pTVDispInfo = reinterpret_cast<TV_DISPINFO*>(pNMHDR);
		*pResult = 0;

		if (!pTVDispInfo || !pTVDispInfo->item.pszText) return;

		LPMAPICONTAINER lpMAPIContainer = nullptr;
		GetContainer(pTVDispInfo->item.hItem, mfcmapiREQUEST_MODIFY, &lpMAPIContainer);
		if (!lpMAPIContainer) return;

		SPropValue sDisplayName;
		sDisplayName.ulPropTag = PR_DISPLAY_NAME;
		sDisplayName.Value.LPSZ = pTVDispInfo->item.pszText;

		EC_MAPI(HrSetOneProp(lpMAPIContainer, &sDisplayName));

		lpMAPIContainer->Release();
	}

	void CHierarchyTableTreeCtrl::OnDblclk(_In_ NMHDR* /*pNMHDR*/, _In_ LRESULT* pResult)
	{
		// Due to problems with focus...we have to post instead of calling OnDisplayItem directly.
		// Post the message to display the item
		if (m_lpHostDlg) m_lpHostDlg->PostMessage(WM_COMMAND, ID_DISPLAYSELECTEDITEM, NULL);

		// Don't do default behavior for double-click (We only want '+' sign expansion.
		// Double click should display the item, not expand the tree.)
		*pResult = 1;
	}

	void CHierarchyTableTreeCtrl::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
	{
		output::DebugPrintEx(DBGMenu, CLASS, L"OnKeyDown", L"0x%X\n", nChar);

		const auto bCtrlPressed = GetKeyState(VK_CONTROL) < 0;
		const auto bShiftPressed = GetKeyState(VK_SHIFT) < 0;
		const auto bMenuPressed = GetKeyState(VK_MENU) < 0;

		if (!bMenuPressed)
		{
			if (VK_RETURN == nChar && bCtrlPressed)
			{
				output::DebugPrintEx(DBGGeneric, CLASS, L"OnKeyDown", L"calling Display Associated Contents\n");
				if (m_lpHostDlg) m_lpHostDlg->PostMessage(WM_COMMAND, ID_DISPLAYASSOCIATEDCONTENTS, NULL);
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
		auto iDlgCode = CTreeCtrl::OnGetDlgCode() | DLGC_WANTMESSAGE;

		// to make sure that the control key is not pressed
		if (GetKeyState(VK_CONTROL) >= 0 && m_hWnd == ::GetFocus())
		{
			// to make sure that the Tab key is pressed
			if (GetKeyState(VK_TAB) < 0) iDlgCode &= ~(DLGC_WANTALLKEYS | DLGC_WANTMESSAGE | DLGC_WANTTAB);
		}

		return iDlgCode;
	}

	void CHierarchyTableTreeCtrl::OnRightClick(_In_ NMHDR* /*pNMHDR*/, _In_ LRESULT* pResult)
	{
		// Send WM_CONTEXTMENU to self
		(void) SendMessage(WM_CONTEXTMENU, reinterpret_cast<WPARAM>(m_hWnd), GetMessagePos());

		// Mark message as handled and suppress default handling
		*pResult = 1;
	}

	void CHierarchyTableTreeCtrl::OnContextMenu(_In_ CWnd* pWnd, CPoint pos)
	{
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

		ui::DisplayContextMenu(m_nIDContextMenu, IDR_MENU_HIERARCHY_TABLE, m_lpHostDlg->m_hWnd, pos.x, pos.y);
	}

	_Check_return_ sortlistdata::SortListData* CHierarchyTableTreeCtrl::GetSelectedItemData() const
	{
		// Find the highlighted item
		const auto Item = GetSelectedItem();
		if (Item)
		{
			return GetSortListData(Item);
		}

		return nullptr;
	}

	_Check_return_ sortlistdata::SortListData* CHierarchyTableTreeCtrl::GetSortListData(HTREEITEM iItem) const
	{
		return reinterpret_cast<sortlistdata::SortListData*>(GetItemData(iItem));
	}

	_Check_return_ LPSBinary CHierarchyTableTreeCtrl::GetSelectedItemEID() const
	{
		// Find the highlighted item
		const auto Item = GetSelectedItem();

		// get the EID associated with it
		if (Item)
		{
			const auto lpData = GetSortListData(Item);
			if (lpData && lpData->Node()) return lpData->Node()->m_lpEntryID;
		}
		return nullptr;
	}

	_Check_return_ LPMAPICONTAINER CHierarchyTableTreeCtrl::GetSelectedContainer(__mfcmapiModifyEnum bModify) const
	{
		LPMAPICONTAINER lpSelectedContainer = nullptr;

		GetContainer(GetSelectedItem(), bModify, &lpSelectedContainer);

		return lpSelectedContainer;
	}

	void CHierarchyTableTreeCtrl::GetContainer(
		HTREEITEM Item,
		__mfcmapiModifyEnum bModify,
		_In_ LPMAPICONTAINER* lppContainer) const
	{
		auto hRes = S_OK;
		ULONG ulObjType = 0;
		SBinary NullBin = {0};
		LPMAPICONTAINER lpContainer = nullptr;

		*lppContainer = nullptr;

		if (!Item) return;

		output::DebugPrintEx(
			DBGGeneric,
			CLASS,
			L"GetContainer",
			L"HTREEITEM = %p, bModify = %d, m_ulContainerType = 0x%X\n",
			Item,
			bModify,
			m_ulContainerType);

		const auto lpData = GetSortListData(Item);

		if (!lpData || !lpData->Node())
		{
			// We didn't get an entryID, so log it and get out of here
			output::DebugPrintEx(
				DBGGeneric, CLASS, L"GetContainer", L"GetSortListData returned NULL or lpEntryID is NULL\n");
			return;
		}

		auto ulFlags = mfcmapiREQUEST_MODIFY == bModify ? MAPI_MODIFY : NULL;

		auto lpCurBin = lpData->Node()->m_lpEntryID;
		if (!lpCurBin) lpCurBin = &NullBin;

		// Check the type of the root container to know whether the MDB or AddrBook object is valid
		// This also allows NULL EIDs to return the root container itself.
		// Use the Root container if we can't decide and log an error
		if (m_lpMapiObjects)
		{
			if (MAPI_ABCONT == m_ulContainerType)
			{
				const auto lpAddrBook = m_lpMapiObjects->GetAddrBook(false); // do not release
				if (lpAddrBook)
				{
					output::DebugPrint(
						DBGGeneric, L"\tCalling OpenEntry on address book with ulFlags = 0x%X\n", ulFlags);

					WC_H(mapi::CallOpenEntry(
						nullptr,
						lpAddrBook,
						nullptr,
						nullptr,
						lpCurBin->cb,
						reinterpret_cast<LPENTRYID>(lpCurBin->lpb),
						nullptr,
						ulFlags,
						&ulObjType,
						reinterpret_cast<LPUNKNOWN*>(&lpContainer)));
				}
			}
			else if (MAPI_FOLDER == m_ulContainerType)
			{
				const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
				if (lpMDB)
				{
					ulFlags = (mfcmapiREQUEST_MODIFY == bModify ? MAPI_MODIFY : NULL) |
							  (m_ulDisplayFlags & dfDeleted ? SHOW_SOFT_DELETES | MAPI_NO_CACHE : NULL);

					WC_H(mapi::CallOpenEntry(
						lpMDB,
						nullptr,
						nullptr,
						nullptr,
						lpCurBin->cb,
						reinterpret_cast<LPENTRYID>(lpCurBin->lpb),
						nullptr,
						ulFlags,
						&ulObjType,
						reinterpret_cast<LPUNKNOWN*>(&lpContainer)));
				}
			}
		}

		// If we didn't get a container above, try to open from m_lpContainer
		if (!lpContainer && m_lpContainer)
		{
			WARNHRESMSG(MAPI_E_CALL_FAILED, IDS_UNKNOWNCONTAINERTYPE);
			hRes = S_OK;
			WC_H(mapi::CallOpenEntry(
				nullptr,
				nullptr,
				m_lpContainer,
				nullptr,
				lpCurBin->cb,
				reinterpret_cast<LPENTRYID>(lpCurBin->lpb),
				nullptr,
				ulFlags,
				&ulObjType,
				reinterpret_cast<LPUNKNOWN*>(&lpContainer)));
		}

		// if we failed because write access was denied, try again if acceptable
		if (!lpContainer && FAILED(hRes) && mfcmapiREQUEST_MODIFY == bModify)
		{
			output::DebugPrint(DBGGeneric, L"\tOpenEntry failed: 0x%X. Will try again without MAPI_MODIFY\n", hRes);
			// We failed to open the item with MAPI_MODIFY.
			// Let's try to open it with NULL
			GetContainer(Item, mfcmapiDO_NOT_REQUEST_MODIFY, &lpContainer);
		}

		// Ok - we're just out of luck
		if (!lpContainer)
		{
			WARNHRESMSG(hRes, IDS_NOCONTAINER);
			hRes = MAPI_E_NOT_FOUND;
		}

		if (lpContainer) *lppContainer = lpContainer;
		output::DebugPrintEx(
			DBGGeneric,
			CLASS,
			L"GetContainer",
			L"returning lpContainer = %p, ulObjType = 0x%X and hRes = 0x%X\n",
			lpContainer,
			ulObjType,
			hRes);
	}

	// When + is clicked, add all entries in the table as children
	void CHierarchyTableTreeCtrl::OnItemExpanding(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
	{
		auto hRes = S_OK;
		*pResult = 0;

		const auto pNMTreeView = reinterpret_cast<NM_TREEVIEW*>(pNMHDR);
		if (pNMTreeView)
		{
			output::DebugPrintEx(
				DBGHierarchy,
				CLASS,
				L"OnItemExpanding",
				L"Expanding item %p \"%ws\" action = 0x%08X state = 0x%08X\n",
				pNMTreeView->itemNew.hItem,
				strings::LPCTSTRToWstring(GetItemText(pNMTreeView->itemOld.hItem)).c_str(),
				pNMTreeView->action,
				pNMTreeView->itemNew.state);
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
		const auto pNMTreeView = reinterpret_cast<LPNMTREEVIEW>(pNMHDR);
		if (pNMTreeView)
		{
			auto* lpData = reinterpret_cast<sortlistdata::SortListData*>(pNMTreeView->itemOld.lParam);
			output::DebugPrintEx(
				DBGHierarchy,
				CLASS,
				L"OnDeleteItem",
				L"Deleting item %p \"%ws\"\n",
				pNMTreeView->itemOld.hItem,
				strings::LPCTSTRToWstring(GetItemText(pNMTreeView->itemOld.hItem)).c_str());

			if (lpData && lpData->Node() && lpData->Node()->m_lpAdviseSink)
			{
				output::DebugPrintEx(
					DBGHierarchy,
					CLASS,
					L"OnDeleteItem",
					L"Unadvising %p, ulAdviseConnection = 0x%08X\n",
					lpData->Node()->m_lpAdviseSink,
					static_cast<int>(lpData->Node()->m_ulAdviseConnection));
			}

			delete lpData;

			if (!m_bShuttingDown)
			{
				// Collapse the parent if this is the last child
				const auto hPrev = TreeView_GetPrevSibling(m_hWnd, pNMTreeView->itemOld.hItem);
				const auto hNext = TreeView_GetNextSibling(m_hWnd, pNMTreeView->itemOld.hItem);

				if (!(hPrev || hNext))
				{
					output::DebugPrintEx(
						DBGHierarchy, CLASS, L"OnDeleteItem", L"%p has no siblings\n", pNMTreeView->itemOld.hItem);
					const auto hParent = TreeView_GetParent(m_hWnd, pNMTreeView->itemOld.hItem);
					TreeView_SetItemState(m_hWnd, hParent, 0, TVIS_EXPANDED | TVIS_EXPANDEDONCE);
					TVITEM tvItem = {0};
					tvItem.hItem = hParent;
					tvItem.mask = TVIF_PARAM;
					if (TreeView_GetItem(m_hWnd, &tvItem) && tvItem.lParam)
					{
						lpData = reinterpret_cast<sortlistdata::SortListData*>(tvItem.lParam);
						if (lpData && lpData->Node())
						{
							lpData->Node()->m_cSubfolders = 0;
						}
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
		const auto tab = reinterpret_cast<TABLE_NOTIFICATION*>(wParam);
		const auto hParent = reinterpret_cast<HTREEITEM>(lParam);

		output::DebugPrintEx(
			DBGHierarchy,
			CLASS,
			L"msgOnAddItem",
			L"Received message add item under: %p =\"%ws\"\n",
			hParent,
			strings::LPCTSTRToWstring(GetItemText(hParent)).c_str());

		// only need to add the node if we're expanded
		const int iState = GetItemState(hParent, NULL);
		if (iState & TVIS_EXPANDEDONCE)
		{
			auto hRes = S_OK;
			// We make this copy here and pass it in to AddNode, where it is grabbed by SortListData::InitializeContents to be part of the item data
			// The mem will be freed when the item data is cleaned up - do not free here
			SRow NewRow = {0};
			NewRow.cValues = tab->row.cValues;
			NewRow.ulAdrEntryPad = tab->row.ulAdrEntryPad;
			WC_MAPI(ScDupPropset(tab->row.cValues, tab->row.lpProps, MAPIAllocateBuffer, &NewRow.lpProps));
			AddNode(&NewRow, hParent, true);
		}
		else
		{
			// in case the item doesn't know it has children, let it know
			TVITEM tvItem = {0};
			tvItem.hItem = hParent;
			tvItem.mask = TVIF_PARAM;
			if (TreeView_GetItem(m_hWnd, &tvItem) && tvItem.lParam)
			{
				const auto lpData = reinterpret_cast<sortlistdata::SortListData*>(tvItem.lParam);
				if (lpData && lpData->Node())
				{
					lpData->Node()->m_cSubfolders = 1;
				}
			}
		}

		return S_OK;
	}

	// WM_MFCMAPI_DELETEITEM
	// Remove the child node.
	_Check_return_ LRESULT CHierarchyTableTreeCtrl::msgOnDeleteItem(WPARAM wParam, LPARAM lParam)
	{
		auto hRes = S_OK;
		auto tab = reinterpret_cast<TABLE_NOTIFICATION*>(wParam);
		const auto hParent = reinterpret_cast<HTREEITEM>(lParam);

		const auto hItemToDelete = FindNode(&tab->propIndex.Value.bin, hParent);

		if (hItemToDelete)
		{
			output::DebugPrintEx(
				DBGHierarchy,
				CLASS,
				L"msgOnDeleteItem",
				L"Received message delete item: %p =\"%ws\"\n",
				hItemToDelete,
				strings::LPCTSTRToWstring(GetItemText(hItemToDelete)).c_str());
			EC_B(DeleteItem(hItemToDelete));
		}

		return hRes;
	}

	// WM_MFCMAPI_MODIFYITEM
	// Update any UI for the node and resort if needed.
	_Check_return_ LRESULT CHierarchyTableTreeCtrl::msgOnModifyItem(WPARAM wParam, LPARAM lParam)
	{
		auto hRes = S_OK;
		auto tab = reinterpret_cast<TABLE_NOTIFICATION*>(wParam);
		const auto hParent = reinterpret_cast<HTREEITEM>(lParam);

		const auto hModifyItem = FindNode(&tab->propIndex.Value.bin, hParent);

		if (hModifyItem)
		{
			output::DebugPrintEx(
				DBGHierarchy,
				CLASS,
				L"msgOnModifyItem",
				L"Received message modify item: %p =\"%ws\"\n",
				hModifyItem,
				strings::LPCTSTRToWstring(GetItemText(hModifyItem)).c_str());

			const auto lpName = PpropFindProp(tab->row.lpProps, tab->row.cValues, PR_DISPLAY_NAME_W);

			std::wstring szText;
			if (mapi::CheckStringProp(lpName, PT_UNICODE))
			{
				szText = lpName->Value.lpszW;
			}
			else
			{
				szText = strings::loadstring(IDS_UNKNOWNNAME);
			}

			TVITEMEXW item = {0};
			item.mask = TVIF_TEXT;
			item.pszText = const_cast<LPWSTR>(szText.c_str());
			item.hItem = hModifyItem;
			EC_B(::SendMessage(m_hWnd, TVM_SETITEMW, 0, reinterpret_cast<LPARAM>(&item)));

			// We make this copy here and pass it in to the node
			// The mem will be freed when the item data is cleaned up - do not free here
			SRow NewRow = {0};
			NewRow.cValues = tab->row.cValues;
			NewRow.ulAdrEntryPad = tab->row.ulAdrEntryPad;
			WC_MAPI(ScDupPropset(tab->row.cValues, tab->row.lpProps, MAPIAllocateBuffer, &NewRow.lpProps));
			auto lpData = new sortlistdata::SortListData();
			if (lpData)
			{
				lpData->InitializeNode(&NewRow);
				SetNodeData(m_hWnd, hModifyItem, lpData);
			}

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
		auto hRes = S_OK;
		const auto hRefreshItem = reinterpret_cast<HTREEITEM>(wParam);
		output::DebugPrintEx(
			DBGHierarchy,
			CLASS,
			L"msgOnRefreshTable",
			L"Received message refresh table: %p =\"%ws\"\n",
			hRefreshItem,
			strings::LPCTSTRToWstring(GetItemText(hRefreshItem)).c_str());

		const int iState = GetItemState(hRefreshItem, NULL);
		if (iState & TVIS_EXPANDED)
		{
			auto hChild = GetChildItem(hRefreshItem);
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
				const auto lpData = GetSortListData(hRefreshItem);

				if (lpData && lpData->Node())
				{
					if (lpData->Node()->m_lpHierarchyTable)
					{
						ULONG ulRowCount = NULL;
						WC_MAPI(lpData->Node()->m_lpHierarchyTable->GetRowCount(NULL, &ulRowCount));
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
	_Check_return_ HTREEITEM CHierarchyTableTreeCtrl::FindNode(_In_ LPSBinary lpInstance, HTREEITEM hParent) const
	{
		if (!lpInstance || !hParent) return nullptr;

		output::DebugPrintEx(
			DBGGeneric,
			CLASS,
			L"FindNode",
			L"Looking for child of: %p =\"%ws\"\n",
			hParent,
			strings::LPCTSTRToWstring(GetItemText(hParent)).c_str());

		auto hCurrent = GetNextItem(hParent, TVGN_CHILD);

		while (hCurrent)
		{
			const auto lpListData = GetSortListData(hCurrent);
			if (lpListData && lpListData->Node())
			{
				const auto lpCurInstance = lpListData->Node()->m_lpInstanceKey;
				if (lpCurInstance)
				{
					if (!memcmp(lpCurInstance->lpb, lpInstance->lpb, lpInstance->cb))
					{
						output::DebugPrintEx(
							DBGGeneric,
							CLASS,
							L"FindNode",
							L"Matched at %p =\"%ws\"\n",
							hCurrent,
							strings::LPCTSTRToWstring(GetItemText(hCurrent)).c_str());
						return hCurrent;
					}
				}
			}

			hCurrent = GetNextItem(hCurrent, TVGN_NEXT);
		}

		output::DebugPrintEx(DBGGeneric, CLASS, L"FindNode", L"No match found\n");
		return nullptr;
	}

	void CHierarchyTableTreeCtrl::OnCustomDraw(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
	{
		ui::CustomDrawTree(pNMHDR, pResult, m_HoverButton, m_hItemCurHover);
	}
}