#include <StdAfx.h>
#include <UI/Controls/StyleTree/HierarchyTableTreeCtrl.h>
#include <UI/Dialogs/BaseDialog.h>
#include <UI/Dialogs/HierarchyTable/HierarchyTableDlg.h>
#include <core/mapi/cache/mapiObjects.h>
#include <UI/UIFunctions.h>
#include <core/mapi/adviseSink.h>
#include <core/sortlistdata/nodeData.h>
#include <core/utility/registry.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFunctions.h>

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

	CHierarchyTableTreeCtrl::~CHierarchyTableTreeCtrl()
	{
		TRACE_DESTRUCTOR(CLASS);
		m_bShuttingDown = true;
		CWnd::DestroyWindow();

		if (m_lpContainer) m_lpContainer->Release();
	}

	void CHierarchyTableTreeCtrl::Create(
		_In_ CWnd* pCreateParent,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		_In_ dialog::CHierarchyTableDlg* lpHostDlg,
		const tableDisplayFlags displayFlags,
		const UINT nIDContextMenu)
	{
		m_lpMapiObjects = lpMapiObjects;
		m_lpHostDlg = lpHostDlg;
		m_displayFlags = displayFlags;
		m_nIDContextMenu = nIDContextMenu;
		m_bSortNodes = true;

		// Setup callbacks
		HasChildrenCallback = [&](auto _1) { return HasChildren(_1); };
		ItemSelectedCallback = [&](auto _1) { return OnItemSelected(_1); };
		KeyDownCallback = [&](auto _1, auto _2, auto _3, auto _4) { return HandleKeyDown(_1, _2, _3, _4); };
		FreeNodeDataCallback = [&](auto _1) { return FreeNodeData(_1); };
		ExpandNodeCallback = [&](auto _1) { return ExpandNode(_1); };
		OnRefreshCallback = [&] { return OnRefresh(); };
		OnLabelEditCallback = [&](auto _1, auto _2) { return OnLabelEdit(_1, _2); };
		OnDisplaySelectedItemCallback = [&]() { return OnDisplaySelectedItem(); };
		OnLastChildDeletedCallback = [&](auto _1) { return OnLastChildDeleted(_1); };
		HandleContextMenuCallback = [&](auto _1, auto _2) { return HandleContextMenu(_1, _2); };
		OnCustomDrawCallback = [&](auto _1, auto _2, auto _3) { return OnCustomDraw(_1, _2, _3); };

		StyleTreeCtrl::Create(pCreateParent, false);
	}

	LRESULT CHierarchyTableTreeCtrl::WindowProc(const UINT message, const WPARAM wParam, const LPARAM lParam)
	{
		switch (message)
		{
		case WM_MFCMAPI_ADDITEM:
			return msgOnAddItem(wParam, lParam);
		case WM_MFCMAPI_DELETEITEM:
			return msgOnDeleteItem(wParam, lParam);
		case WM_MFCMAPI_MODIFYITEM:
			return msgOnModifyItem(wParam, lParam);
		case WM_MFCMAPI_REFRESHTABLE:
			return msgOnRefreshTable(wParam, lParam);
		}

		return StyleTreeCtrl::WindowProc(message, wParam, lParam);
	}

	void CHierarchyTableTreeCtrl::OnRefresh() const
	{
		AddRootNode();

		if (m_lpHostDlg) m_lpHostDlg->OnUpdateSingleMAPIPropListCtrl(nullptr, nullptr);
	}

	void CHierarchyTableTreeCtrl::LoadHierarchyTable(_In_ const LPMAPICONTAINER lpMAPIContainer)
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (m_lpContainer) m_lpContainer->Release();
		m_lpContainer = lpMAPIContainer;
		m_ulContainerType = NULL;

		// If we weren't passed a container to load, give up
		if (!m_lpContainer) return;
		m_lpContainer->AddRef();

		m_ulContainerType = mapi::GetMAPIObjectType(lpMAPIContainer);

		Refresh();
	}

	void CHierarchyTableTreeCtrl::FreeNodeData(const LPARAM lpData) const noexcept
	{
		auto* lpNodeData = reinterpret_cast<sortlistdata::sortListData*>(lpData);
		delete lpNodeData;
	}

	void CHierarchyTableTreeCtrl::OnItemAdded(HTREEITEM hItem) const
	{
		if (registry::hierRootNotifs || hItem != TVI_ROOT)
		{
			(void) GetHierarchyTable(hItem, nullptr, true);
		}
	}

	void CHierarchyTableTreeCtrl::AddRootNode() const
	{
		if (!m_lpContainer || !m_hWnd) return;
		LPSPropValue lpProps = nullptr;
		LPSPropValue lpRootName = nullptr; // don't free
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		ULONG cVals = 0;

		WC_H_GETPROPS_S(m_lpContainer->GetProps(LPSPropTagArray(&sptHTCols), fMapiUnicode, &cVals, &lpProps));

		// Get the Display Name for the Root Container
		if (!lpProps || PT_ERROR == PROP_TYPE(lpProps[htPR_DISPLAY_NAME_W].ulPropTag))
		{
			output::DebugPrint(
				output::dbgLevel::Hierarchy,
				L"Could not find Display Name for Root Container. This is benign. Assuming NULL.\n");
		}
		else
			lpRootName = &lpProps[htPR_DISPLAY_NAME_W];

		std::wstring szName;

		// Shouldn't have to check lpRootName for non-NULL since CheckString does it, but prefast is complaining
		if (lpRootName && strings::CheckStringProp(lpRootName, PT_UNICODE))
		{
			szName = lpRootName->Value.lpszW;
		}
		else
		{
			szName = strings::loadstring(IDS_ROOTCONTAINER);
		}

		auto lpData = new (std::nothrow) sortlistdata::sortListData();
		if (lpData)
		{
			sortlistdata::nodeData::init(
				lpData,
				cVals,
				lpProps // Pass our lpProps to be archived
			);

			(void) AddChildNode(
				szName, TVI_ROOT, reinterpret_cast<LPARAM>(lpData), [&](auto _1) { return OnItemAdded(_1); });
		}

		// Node owns the lpProps memory now, so we don't free it
	}

	void CHierarchyTableTreeCtrl::AddNode(
		_In_ const LPSRow lpsRow,
		HTREEITEM hParent,
		const std::function<void(HTREEITEM hItem)>& callback) const
	{
		if (!lpsRow) return;

		std::wstring szName;
		const auto lpName = PpropFindProp(lpsRow->lpProps, lpsRow->cValues, PR_DISPLAY_NAME_W);
		if (strings::CheckStringProp(lpName, PT_UNICODE))
		{
			szName = lpName->Value.lpszW;
		}
		else
		{
			szName = strings::loadstring(IDS_UNKNOWNNAME);
		}

		output::DebugPrintEx(
			output::dbgLevel::Hierarchy, CLASS, L"AddNode", L"Adding to %p: %ws\n", hParent, szName.c_str());

		auto lpData = new (std::nothrow) sortlistdata::sortListData();
		if (lpData)
		{
			sortlistdata::nodeData::init(lpData, lpsRow->cValues, lpsRow->lpProps);

			(void) AddChildNode(szName, hParent, reinterpret_cast<LPARAM>(lpData), callback);
		}
	}

	void CHierarchyTableTreeCtrl::Advise(HTREEITEM hItem, sortlistdata::sortListData* lpData) const
	{
		// Set up our advise sink
		if (lpData)
		{
			const auto node = lpData->cast<sortlistdata::nodeData>();
			if (node && !node->hasSink() && (registry::hierRootNotifs || hItem != GetRootItem()))
			{
				if (node->advise(m_hWnd, hItem, m_lpMapiObjects->GetMDB()))
				{
					output::DebugPrintEx(
						output::dbgLevel::Notify,
						CLASS,
						L"GetHierarchyTable",
						L"Advise sink for \"%ws\" = %p\n",
						GetItemTextW(hItem).c_str(),
						hItem);
				}
			}
		}
	}

	_Check_return_ LPMAPITABLE CHierarchyTableTreeCtrl::GetHierarchyTable(
		HTREEITEM hItem,
		_In_opt_ LPMAPICONTAINER lpMAPIContainer,
		const bool bRegNotifs) const
	{
		const auto lpData = GetSortListData(hItem);

		if (!lpData) return nullptr;
		const auto node = lpData->cast<sortlistdata::nodeData>();
		if (!node) return nullptr;

		if (!node->hasTable())
		{
			if (lpMAPIContainer)
			{
				lpMAPIContainer->AddRef();
			}
			else
			{
				lpMAPIContainer = GetContainer(hItem, modifyType::DO_NOT_REQUEST_MODIFY);
			}

			if (lpMAPIContainer)
			{
				// Get the hierarchy table for the node and shove it into the data
				LPMAPITABLE lpHierarchyTable = nullptr;

				// On the AB, something about this call triggers table reloads on the parent hierarchy table
				// No idea why they're triggered - doesn't happen for all AB providers
				WC_MAPI_S(lpMAPIContainer->GetHierarchyTable(
					(m_displayFlags && tableDisplayFlags::dfDeleted ? SHOW_SOFT_DELETES : NULL) | fMapiUnicode,
					&lpHierarchyTable));

				if (lpHierarchyTable)
				{
					EC_MAPI_S(lpHierarchyTable->SetColumns(LPSPropTagArray(&sptHTCols), TBL_BATCH));
				}

				node->setTable(lpHierarchyTable);
				lpMAPIContainer->Release();
			}
		}

		// Set up our advise sink
		if (bRegNotifs)
		{
			Advise(hItem, lpData);
		}

		return node->getTable();
	}

	// Add the first level contents of lpMAPIContainer under the Parent node
	void CHierarchyTableTreeCtrl::ExpandNode(HTREEITEM hParent) const
	{
		output::DebugPrintEx(output::dbgLevel::Hierarchy, CLASS, L"ExpandNode", L"Expanding %p\n", hParent);
		if (!m_hWnd || !hParent) return;

		CWaitCursor Wait; // Change the mouse to an hourglass while we work.
		auto pRows = LPSRowSet{};
		auto lpHierarchyTable = GetHierarchyTable(hParent, nullptr, registry::hierExpandNotifications);
		if (lpHierarchyTable)
		{
			// Go to the first row
			EC_MAPI_S(lpHierarchyTable->SeekRow(BOOKMARK_BEGINNING, 0, nullptr));

			auto i = ULONG{};
			// Get each row in turn and add it to the list
			// TODO: Query several rows at once
			for (;;)
			{
				// Note - we're saving the rows off in AddNode, so we don't FreeProws this...we just MAPIFreeBuffer the array
				if (pRows) MAPIFreeBuffer(pRows);
				pRows = nullptr;
				const auto hRes = EC_MAPI(lpHierarchyTable->QueryRows(1, NULL, &pRows));
				if (FAILED(hRes) || !pRows || !pRows->cRows) break;
				// Now we can process the row!

				AddNode(pRows->aRow, hParent, nullptr);
				i++;
			}
		}

		// Note - we're saving the props off in AddNode, so we don't FreeProws this...we just MAPIFreeBuffer the array
		if (pRows) MAPIFreeBuffer(pRows);
	}

	bool CHierarchyTableTreeCtrl::HasChildren(_In_ HTREEITEM hItem) const
	{
		if (m_displayFlags && tableDisplayFlags::dfDeleted)
		{
			return true;
		}

		const auto lpData = reinterpret_cast<sortlistdata::sortListData*>(hItem);

		if (lpData)
		{
			const auto node = lpData->cast<sortlistdata::nodeData>();
			if (node)
			{
				auto cSubfolders = node->getSubfolders();
				if (cSubfolders >= 0)
				{
					return cSubfolders > 0;
				}

				LPCTSTR szName = nullptr;
				if (PROP_TYPE(lpData->lpSourceProps[0].ulPropTag) == PT_TSTRING)
					szName = lpData->lpSourceProps[0].Value.LPSZ;
				output::DebugPrintEx(
					output::dbgLevel::Hierarchy,
					CLASS,
					L"HasChildren",
					L"Using Hierarchy table %d %ws\n",
					cSubfolders,
					strings::LPCTSTRToWstring(szName).c_str());
				// Won't force the hierarchy table - just get it if we've already got it
				return node->hasChildren();
			}
		}

		return false;
	}

	void CHierarchyTableTreeCtrl::OnItemSelected(HTREEITEM hItem) const
	{
		output::DebugPrintEx(output::dbgLevel::Hierarchy, CLASS, L"OnItemSelected", L"%p\n", hItem);

		if (!m_lpHostDlg) return;

		// Have to request modify or this object is read only in the single prop control.
		auto lpMAPIContainer = GetContainer(hItem, modifyType::REQUEST_MODIFY);

		// Make sure we've gotten the hierarchy table for this node
		(void) GetHierarchyTable(hItem, lpMAPIContainer, registry::hierExpandNotifications);

		UINT uiMsg = IDS_STATUSTEXTNOFOLDER;
		std::wstring szParam1;
		std::wstring szParam2;
		std::wstring szParam3;

		if (lpMAPIContainer)
		{
			// Get some props for status bar
			LPSPropValue lpProps = nullptr;
			ULONG cVals = 0;
			WC_H_GETPROPS_S(
				lpMAPIContainer->GetProps(LPSPropTagArray(&sptHTCountCols), fMapiUnicode, &cVals, &lpProps));
			if (lpProps)
			{
				if (!(m_displayFlags && tableDisplayFlags::dfDeleted))
				{
					uiMsg = IDS_STATUSTEXTCONTENTCOUNTS;
					if (PT_ERROR == PROP_TYPE(lpProps[htcPR_CONTENT_COUNT].ulPropTag))
					{
						WARNHRESMSG(lpProps[htcPR_CONTENT_COUNT].Value.err, IDS_NODELACKSCONTENTCOUNT);
						szParam1 = L"?"; // STRING_OK
					}
					else
					{
						szParam1 = std::to_wstring(lpProps[htcPR_CONTENT_COUNT].Value.ul);
					}

					if (PT_ERROR == PROP_TYPE(lpProps[htcPR_ASSOC_CONTENT_COUNT].ulPropTag))
					{
						WARNHRESMSG(lpProps[htcPR_ASSOC_CONTENT_COUNT].Value.err, IDS_NODELACKSASSOCCONTENTCOUNT);
						szParam2 = L"?"; // STRING_OK
					}
					else
					{
						szParam2 = std::to_wstring(lpProps[htcPR_ASSOC_CONTENT_COUNT].Value.ul);
					}
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
					{
						szParam1 = std::to_wstring(lpProps[htcPR_DELETED_MSG_COUNT].Value.ul);
					}

					if (PT_ERROR == PROP_TYPE(lpProps[htcPR_DELETED_ASSOC_MSG_COUNT].ulPropTag))
					{
						WARNHRESMSG(
							lpProps[htcPR_DELETED_ASSOC_MSG_COUNT].Value.err, IDS_NODELACKSDELETEDASSOCMESSAGECOUNT);
						szParam2 = L"?"; // STRING_OK
					}
					else
					{
						szParam2 = std::to_wstring(lpProps[htcPR_DELETED_ASSOC_MSG_COUNT].Value.ul);
					}

					if (PT_ERROR == PROP_TYPE(lpProps[htcPR_DELETED_FOLDER_COUNT].ulPropTag))
					{
						WARNHRESMSG(lpProps[htcPR_DELETED_FOLDER_COUNT].Value.err, IDS_NODELACKSDELETEDSUBFOLDERCOUNT);
						szParam3 = L"?"; // STRING_OK
					}
					else
					{
						szParam3 = std::to_wstring(lpProps[htcPR_DELETED_FOLDER_COUNT].Value.ul);
					}
				}

				MAPIFreeBuffer(lpProps);
			}
		}

		m_lpHostDlg->UpdateStatusBarText(statusPane::data1, uiMsg, szParam1, szParam2, szParam3);

		m_lpHostDlg->OnUpdateSingleMAPIPropListCtrl(lpMAPIContainer, GetSortListData(hItem));

		WCHAR szText[255] = {};
		auto item = TVITEMEXW{};
		item.mask = TVIF_TEXT;
		item.pszText = szText;
		item.cchTextMax = _countof(szText);
		item.hItem = GetSelectedItem();
		WC_B_S(::SendMessage(m_hWnd, TVM_GETITEMW, 0, reinterpret_cast<LPARAM>(&item)));
		m_lpHostDlg->UpdateTitleBarText(szText);

		if (lpMAPIContainer) lpMAPIContainer->Release();
	}

	void CHierarchyTableTreeCtrl::OnLabelEdit(HTREEITEM hItem, LPTSTR szText)
	{
		auto lpMAPIContainer = GetContainer(hItem, modifyType::REQUEST_MODIFY);
		if (!lpMAPIContainer) return;

		SPropValue sDisplayName;
		sDisplayName.ulPropTag = PR_DISPLAY_NAME;
		sDisplayName.Value.LPSZ = szText;

		EC_MAPI_S(HrSetOneProp(lpMAPIContainer, &sDisplayName));

		lpMAPIContainer->Release();
	}

	void CHierarchyTableTreeCtrl::OnDisplaySelectedItem()
	{
		// Due to problems with focus...we have to post instead of calling OnDisplayItem directly.
		// Post the message to display the item
		if (m_lpHostDlg) m_lpHostDlg->PostMessage(WM_COMMAND, ID_DISPLAYSELECTEDITEM, NULL);
	}

	bool CHierarchyTableTreeCtrl::HandleKeyDown(
		const UINT nChar,
		const bool bShiftPressed,
		const bool bCtrlPressed,
		const bool bMenuPressed)
	{
		if (VK_RETURN == nChar && bCtrlPressed)
		{
			output::DebugPrintEx(
				output::dbgLevel::Generic, CLASS, L"OnKeyDown", L"calling Display Associated Contents\n");
			if (m_lpHostDlg) m_lpHostDlg->PostMessage(WM_COMMAND, ID_DISPLAYASSOCIATEDCONTENTS, NULL);
			return true;
		}

		if (m_lpHostDlg) return m_lpHostDlg->HandleKeyDown(nChar, bShiftPressed, bCtrlPressed, bMenuPressed);

		return false;
	}

	void CHierarchyTableTreeCtrl::HandleContextMenu(const int x, const int y)
	{
		ui::DisplayContextMenu(m_nIDContextMenu, IDR_MENU_HIERARCHY_TABLE, m_lpHostDlg->m_hWnd, x, y);
	}

	_Check_return_ sortlistdata::sortListData* CHierarchyTableTreeCtrl::GetSelectedItemData() const
	{
		// Find the highlighted item
		const auto Item = GetSelectedItem();
		if (Item)
		{
			return GetSortListData(Item);
		}

		return nullptr;
	}

	_Check_return_ sortlistdata::sortListData* CHierarchyTableTreeCtrl::GetSortListData(HTREEITEM iItem) const
	{
		return reinterpret_cast<sortlistdata::sortListData*>(GetItemData(iItem));
	}

	_Check_return_ std::vector<BYTE> CHierarchyTableTreeCtrl::GetSelectedItemEID() const
	{
		// Find the highlighted item
		const auto Item = GetSelectedItem();

		// Get the EID associated with it
		if (Item)
		{
			const auto lpData = GetSortListData(Item);
			if (lpData)
			{
				const auto node = lpData->cast<sortlistdata::nodeData>();
				if (node)
				{
					return node->getEntryID();
				}
			}
		}

		return {};
	}

	_Check_return_ LPMAPICONTAINER CHierarchyTableTreeCtrl::GetSelectedContainer(const modifyType bModify) const
	{
		return GetContainer(GetSelectedItem(), bModify);
	}

	_Check_return_ LPMAPICONTAINER CHierarchyTableTreeCtrl::GetContainer(HTREEITEM Item, const modifyType bModify) const
	{
		if (!Item) return nullptr;

		auto hRes = S_OK;
		auto ulObjType = ULONG{};
		auto NullBin = SBinary{};
		LPMAPICONTAINER lpContainer = nullptr;

		output::DebugPrintEx(
			output::dbgLevel::Generic,
			CLASS,
			L"GetContainer",
			L"HTREEITEM = %p, bModify = %d, m_ulContainerType = 0x%X\n",
			Item,
			bModify,
			m_ulContainerType);

		const auto lpData = GetSortListData(Item);
		if (!lpData)
		{
			// We didn't get an entryID, so log it and get out of here
			output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"GetContainer", L"GetSortListData returned NULL\n");
			return nullptr;
		}

		const auto node = lpData->cast<sortlistdata::nodeData>();
		if (!node)
		{
			// We didn't get an entryID, so log it and get out of here
			output::DebugPrintEx(
				output::dbgLevel::Generic, CLASS, L"GetContainer", L"GetSortListData: lpEntryID is NULL\n");
			return nullptr;
		}

		auto ulFlags = modifyType::REQUEST_MODIFY == bModify ? MAPI_MODIFY : NULL;

		const auto curBin = node->getEntryID();

		// Check the type of the root container to know whether the MDB or AddrBook object is valid
		// This also allows NULL EIDs to return the root container itself.
		// Use the Root container if we can't decide and log an error
		if (m_lpMapiObjects)
		{
			if (m_ulContainerType == MAPI_ABCONT)
			{
				const auto lpAddrBook = m_lpMapiObjects->GetAddrBook(false); // do not release
				if (lpAddrBook)
				{
					output::DebugPrint(
						output::dbgLevel::Generic,
						L"\tCalling OpenEntry on address book with ulFlags = 0x%X\n",
						ulFlags);

					lpContainer = mapi::CallOpenEntry<LPMAPICONTAINER>(
						nullptr,
						lpAddrBook,
						nullptr,
						nullptr,
						static_cast<ULONG>(curBin.size()),
						mapi::toEntryID(curBin),
						nullptr,
						ulFlags,
						&ulObjType);
				}
			}
			else if (m_ulContainerType == MAPI_FOLDER)
			{
				const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
				if (lpMDB)
				{
					ulFlags =
						(modifyType::REQUEST_MODIFY == bModify ? MAPI_MODIFY : NULL) |
						(m_displayFlags && tableDisplayFlags::dfDeleted ? SHOW_SOFT_DELETES | MAPI_NO_CACHE : NULL);

					lpContainer = mapi::CallOpenEntry<LPMAPICONTAINER>(
						lpMDB,
						nullptr,
						nullptr,
						nullptr,
						static_cast<ULONG>(curBin.size()),
						mapi::toEntryID(curBin),
						nullptr,
						ulFlags,
						&ulObjType);
				}
			}
		}

		// If we didn't get a container above, try to open from m_lpContainer
		if (!lpContainer && m_lpContainer)
		{
			WARNHRESMSG(MAPI_E_CALL_FAILED, IDS_UNKNOWNCONTAINERTYPE);
			lpContainer = mapi::CallOpenEntry<LPMAPICONTAINER>(
				nullptr,
				nullptr,
				m_lpContainer,
				nullptr,
				static_cast<ULONG>(curBin.size()),
				mapi::toEntryID(curBin),
				nullptr,
				ulFlags,
				&ulObjType);
		}

		// If we failed because write access was denied, try again if acceptable
		if (!lpContainer && FAILED(hRes) && modifyType::REQUEST_MODIFY == bModify)
		{
			output::DebugPrint(
				output::dbgLevel::Generic, L"\tOpenEntry failed: 0x%X. Will try again without MAPI_MODIFY\n", hRes);
			// We failed to open the item with MAPI_MODIFY.
			// Let's try to open it with NULL
			lpContainer = GetContainer(Item, modifyType::DO_NOT_REQUEST_MODIFY);
		}

		// Ok - we're just out of luck
		if (!lpContainer)
		{
			WARNHRESMSG(hRes, IDS_NOCONTAINER);
			hRes = MAPI_E_NOT_FOUND;
		}

		output::DebugPrintEx(
			output::dbgLevel::Generic,
			CLASS,
			L"GetContainer",
			L"returning lpContainer = %p, ulObjType = 0x%X and hRes = 0x%X\n",
			lpContainer,
			ulObjType,
			hRes);

		return lpContainer;
	}

	void CHierarchyTableTreeCtrl::OnLastChildDeleted(const LPARAM lpData)
	{
		const auto lpNodeData = reinterpret_cast<sortlistdata::sortListData*>(lpData);
		if (lpNodeData)
		{
			const auto node = lpNodeData->cast<sortlistdata::nodeData>();
			if (node)
			{
				node->setSubfolders(false);
			}
		}
	}

	// WM_MFCMAPI_ADDITEM
	// If the parent has been expanded once, add the new row
	// Otherwise, ditch the notification
	_Check_return_ LRESULT CHierarchyTableTreeCtrl::msgOnAddItem(const WPARAM wParam, const LPARAM lParam)
	{
		const auto tab = reinterpret_cast<TABLE_NOTIFICATION*>(wParam);
		const auto hParent = reinterpret_cast<HTREEITEM>(lParam);

		output::DebugPrintEx(
			output::dbgLevel::Hierarchy,
			CLASS,
			L"msgOnAddItem",
			L"Received message add item under: %p =\"%ws\"\n",
			hParent,
			GetItemTextW(hParent).c_str());

		// Only need to add the node if we're expanded
		const int iState = GetItemState(hParent, NULL);
		if (iState & TVIS_EXPANDEDONCE)
		{
			// We make this copy here and pass it in to AddNode, where it is grabbed by sortListData::InitializeContents to be part of the item data
			// The mem will be freed when the item data is cleaned up - do not free here
			auto NewRow = SRow{};
			NewRow.cValues = tab->row.cValues;
			NewRow.ulAdrEntryPad = tab->row.ulAdrEntryPad;
			WC_MAPI_S(ScDupPropset(tab->row.cValues, tab->row.lpProps, MAPIAllocateBuffer, &NewRow.lpProps));
			AddNode(&NewRow, hParent, [&](auto _1) { return OnItemAdded(_1); });
		}
		else
		{
			// in case the item doesn't know it has children, let it know
			auto tvItem = TVITEM{};
			tvItem.hItem = hParent;
			tvItem.mask = TVIF_PARAM;
			if (TreeView_GetItem(m_hWnd, &tvItem) && tvItem.lParam)
			{
				const auto lpData = reinterpret_cast<sortlistdata::sortListData*>(tvItem.lParam);
				if (lpData)
				{
					const auto node = lpData->cast<sortlistdata::nodeData>();
					if (node)
					{
						node->setSubfolders(true);
					}
				}
			}
		}

		return S_OK;
	}

	// WM_MFCMAPI_DELETEITEM
	// Remove the child node.
	_Check_return_ LRESULT CHierarchyTableTreeCtrl::msgOnDeleteItem(const WPARAM wParam, const LPARAM lParam)
	{
		auto hRes = S_OK;
		auto tab = reinterpret_cast<TABLE_NOTIFICATION*>(wParam);
		const auto hParent = reinterpret_cast<HTREEITEM>(lParam);

		const auto hItemToDelete = FindNode(mapi::getBin(tab->propIndex), hParent);

		if (hItemToDelete)
		{
			output::DebugPrintEx(
				output::dbgLevel::Hierarchy,
				CLASS,
				L"msgOnDeleteItem",
				L"Received message delete item: %p =\"%ws\"\n",
				hItemToDelete,
				GetItemTextW(hItemToDelete).c_str());
			hRes = EC_B(DeleteItem(hItemToDelete));
		}

		return hRes;
	}

	// WM_MFCMAPI_MODIFYITEM
	// Update any UI for the node and resort if needed.
	_Check_return_ LRESULT CHierarchyTableTreeCtrl::msgOnModifyItem(const WPARAM wParam, const LPARAM lParam)
	{
		auto hRes = S_OK;
		auto tab = reinterpret_cast<TABLE_NOTIFICATION*>(wParam);
		const auto hParent = reinterpret_cast<HTREEITEM>(lParam);

		const auto hModifyItem = FindNode(mapi::getBin(tab->propIndex), hParent);

		if (hModifyItem)
		{
			output::DebugPrintEx(
				output::dbgLevel::Hierarchy,
				CLASS,
				L"msgOnModifyItem",
				L"Received message modify item: %p =\"%ws\"\n",
				hModifyItem,
				GetItemTextW(hModifyItem).c_str());

			const auto lpName = PpropFindProp(tab->row.lpProps, tab->row.cValues, PR_DISPLAY_NAME_W);

			std::wstring szText;
			if (strings::CheckStringProp(lpName, PT_UNICODE))
			{
				szText = lpName->Value.lpszW;
			}
			else
			{
				szText = strings::loadstring(IDS_UNKNOWNNAME);
			}

			auto item = TVITEMEXW{};
			item.mask = TVIF_TEXT;
			item.pszText = const_cast<LPWSTR>(szText.c_str());
			item.hItem = hModifyItem;
			EC_B_S(::SendMessage(m_hWnd, TVM_SETITEMW, 0, reinterpret_cast<LPARAM>(&item)));

			const auto lpData = GetSortListData(hModifyItem);
			if (lpData)
			{
				auto node = lpData->cast<sortlistdata::nodeData>();
				if (node)
				{
					// We make this copy here and pass it in to the existing node
					// The mem will be freed when the item data is cleaned up - do not free here
					LPSPropValue newProps = {};
					hRes = WC_MAPI(ScDupPropset(tab->row.cValues, tab->row.lpProps, MAPIAllocateBuffer, &newProps));
					node->init(tab->row.cValues, newProps);
				}
			}

			if (hParent)
			{
				hRes = EC_B(SortChildren(hParent));
			}
		}

		if (hModifyItem == GetSelectedItem()) OnItemSelected(hModifyItem);

		return hRes;
	}

	// WM_MFCMAPI_REFRESHTABLE
	// If node was expanded, collapse it to remove all children
	// Then, if the node does have children, reexpand it.
	_Check_return_ LRESULT CHierarchyTableTreeCtrl::msgOnRefreshTable(const WPARAM wParam, LPARAM /*lParam*/)
	{
		auto hRes = S_OK;
		const auto hRefreshItem = reinterpret_cast<HTREEITEM>(wParam);
		output::DebugPrintEx(
			output::dbgLevel::Hierarchy,
			CLASS,
			L"msgOnRefreshTable",
			L"Received message refresh table: %p =\"%ws\"\n",
			hRefreshItem,
			GetItemTextW(hRefreshItem).c_str());

		const int iState = GetItemState(hRefreshItem, NULL);
		if (iState & TVIS_EXPANDED)
		{
			auto hChild = GetChildItem(hRefreshItem);
			while (hChild)
			{
				EC_B_S(DeleteItem(hChild));
				hChild = GetChildItem(hRefreshItem);
			}

			// Reset our expanded bits
			hRes = EC_B(SetItemState(hRefreshItem, NULL, TVIS_EXPANDED | TVIS_EXPANDEDONCE));
			const auto lpData = GetSortListData(hRefreshItem);

			if (lpData)
			{
				const auto node = lpData->cast<sortlistdata::nodeData>();
				if (node && node->hasChildren())
				{
					hRes = EC_B(Expand(hRefreshItem, TVE_EXPAND));
				}
			}
		}

		return hRes;
	}

	// This function steps through the list control to find the entry with this instance key
	// Return NULL if item not found
	_Check_return_ HTREEITEM CHierarchyTableTreeCtrl::FindNode(_In_ const SBinary& instance, HTREEITEM hParent) const
	{
		if (!hParent) return nullptr;

		output::DebugPrintEx(
			output::dbgLevel::Generic,
			CLASS,
			L"FindNode",
			L"Looking for child of: %p =\"%ws\"\n",
			hParent,
			GetItemTextW(hParent).c_str());

		auto hCurrent = GetNextItem(hParent, TVGN_CHILD);

		while (hCurrent)
		{
			const auto lpListData = GetSortListData(hCurrent);

			if (lpListData)
			{
				const auto node = lpListData->cast<sortlistdata::nodeData>();
				if (node && node->matchInstanceKey(&instance))
				{
					output::DebugPrintEx(
						output::dbgLevel::Generic,
						CLASS,
						L"FindNode",
						L"Matched at %p =\"%ws\"\n",
						hCurrent,
						GetItemTextW(hCurrent).c_str());
					return hCurrent;
				}
			}

			hCurrent = GetNextItem(hCurrent, TVGN_NEXT);
		}

		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"FindNode", L"No match found\n");
		return nullptr;
	}

	void
	CHierarchyTableTreeCtrl::OnCustomDraw(_In_ NMHDR* pNMHDR, _In_ LRESULT* /*pResult*/, _In_ HTREEITEM hItemCurHover)
	{
		const auto lvcd = reinterpret_cast<LPNMTVCUSTOMDRAW>(pNMHDR);
		if (!lvcd) return;

		switch (lvcd->nmcd.dwDrawStage)
		{
		case CDDS_ITEMPOSTPAINT:
		{
			const auto hItem = reinterpret_cast<HTREEITEM>(lvcd->nmcd.dwItemSpec);
			if (hItem)
			{
				// If we've advised on this object, add a little icon to let the user know
				// Paint the advise icon, IDB_ADVISE
				auto tvi = TVITEM{};
				tvi.mask = TVIF_PARAM;
				tvi.hItem = hItem;
				TreeView_GetItem(lvcd->nmcd.hdr.hwndFrom, &tvi);
				const auto lpData = reinterpret_cast<sortlistdata::sortListData*>(tvi.lParam);
				if (lpData)
				{
					const auto node = lpData->cast<sortlistdata::nodeData>();
					if (node && node->hasSink())
					{
						auto rect = RECT{};
						TreeView_GetItemRect(lvcd->nmcd.hdr.hwndFrom, hItem, &rect, 1);
						rect.left = rect.right;
						rect.right += rect.bottom - rect.top;
						DrawBitmap(lvcd->nmcd.hdc, rect, ui::uiBitmap::Notify, hItem == hItemCurHover);
					}
				}
			}
			break;
		}
		}
	}
} // namespace controls