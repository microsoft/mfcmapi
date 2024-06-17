#include <StdAfx.h>
#include <UI/Controls/SortList/ContentsTableListCtrl.h>
#include <core/sortlistdata/sortListData.h>
#include <core/mapi/cache/mapiObjects.h>
#include <UI/UIFunctions.h>
#include <core/mapi/adviseSink.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Dialogs/Editors/TagArrayEditor.h>
#include <core/mapi/extraPropTags.h>
#include <core/smartview/SmartView.h>
#include <core/sortlistdata/contentsData.h>
#include <UI/Dialogs/BaseDialog.h>
#include <UI/Dialogs/ContentsTable/ContentsTableDlg.h>
#include <core/mapi/cache/namedProps.h>
#include <core/mapi/mapiMemory.h>
#include <core/utility/registry.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/interpret/flags.h>
#include <core/interpret/proptags.h>
#include <core/mapi/mapiOutput.h>
#include <core/mapi/mapiFunctions.h>
#include <core/property/parseProperty.h>
#include <core/utility/clipboard.h>

namespace controls::sortlistctrl
{
	static std::wstring CLASS = L"CContentsTableListCtrl";

	CContentsTableListCtrl::CContentsTableListCtrl(
		_In_ CWnd* pCreateParent,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		_In_ LPSPropTagArray sptDefaultDisplayColumnTags,
		_In_ const std::vector<columns::TagNames>& lpDefaultDisplayColumns,
		UINT nIDContextMenu,
		bool bIsAB,
		_In_ dialog::CContentsTableDlg* lpHostDlg)
		: m_lpMapiObjects(lpMapiObjects), m_lpHostDlg(lpHostDlg),
		  m_sptDefaultDisplayColumnTags(sptDefaultDisplayColumnTags),
		  m_lpDefaultDisplayColumns(lpDefaultDisplayColumns), m_nIDContextMenu(nIDContextMenu), m_bIsAB(bIsAB)

	{
		TRACE_CONSTRUCTOR(CLASS);
		if (m_lpHostDlg) m_lpHostDlg->AddRef();

		Create(pCreateParent, LVS_NOCOLUMNHEADER, IDC_LIST_CTRL, true);
	}

	CContentsTableListCtrl::~CContentsTableListCtrl()
	{
		TRACE_DESTRUCTOR(CLASS);

		NotificationOff();

		if (m_lpRes) MAPIFreeBuffer(const_cast<LPSRestriction>(m_lpRes));
		if (m_lpContentsTable) m_lpContentsTable->Release();
		if (m_lpHostDlg) m_lpHostDlg->Release();
	}

	BEGIN_MESSAGE_MAP(CContentsTableListCtrl, CSortListCtrl)
#pragma warning(push)
#pragma warning( \
	disable : 26454) // Warning C26454 Arithmetic overflow: 'operator' operation produces a negative unsigned result at compile time
	ON_NOTIFY_REFLECT(LVN_ITEMCHANGED, OnItemChanged)
#pragma warning(pop)
	ON_WM_KEYDOWN()
	ON_WM_CONTEXTMENU()
	ON_MESSAGE(WM_MFCMAPI_ADDITEM, msgOnAddItem)
	ON_MESSAGE(WM_MFCMAPI_THREADADDITEM, msgOnThreadAddItem)
	ON_MESSAGE(WM_MFCMAPI_DELETEITEM, msgOnDeleteItem)
	ON_MESSAGE(WM_MFCMAPI_MODIFYITEM, msgOnModifyItem)
	ON_MESSAGE(WM_MFCMAPI_REFRESHTABLE, msgOnRefreshTable)
	END_MESSAGE_MAP()

	LRESULT CContentsTableListCtrl::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
	{
		auto hRes = S_OK;

		switch (message)
		{
		case WM_ERASEBKGND:
			if (!m_lpContentsTable)
			{
				return true;
			}

			break;
		case WM_PAINT:
			if (LVS_NOCOLUMNHEADER & GetStyle())
			{
				ui::DrawHelpText(m_hWnd, IDS_HELPTEXTSTARTHERE);
				return true;
			}

			break;
		case WM_LBUTTONDBLCLK:
			hRes = WC_H(DoExpandCollapse());
			if (hRes == S_FALSE)
			{
				// Post the message to display the item
				if (m_lpHostDlg) m_lpHostDlg->PostMessage(WM_COMMAND, ID_DISPLAYSELECTEDITEM, NULL);
			}
			else
			{
				CHECKHRESMSG(hRes, IDS_EXPANDCOLLAPSEFAILED);
			}

			return NULL;
		}

		return CSortListCtrl::WindowProc(message, wParam, lParam);
	}

	void CContentsTableListCtrl::OnContextMenu(_In_ CWnd* pWnd, CPoint pos)
	{
		if (pWnd && -1 == pos.x && -1 == pos.y)
		{
			POINT point = {};
			const auto iItem = GetNextItem(-1, LVNI_SELECTED);
			if (GetItemPosition(iItem, &point))
			{
				if (::ClientToScreen(pWnd->m_hWnd, &point))
				{
					pos = point;
				}
			}
		}

		ui::DisplayContextMenu(m_nIDContextMenu, IDR_MENU_TABLE, m_lpHostDlg->m_hWnd, pos.x, pos.y);
	}

	_Check_return_ ULONG CContentsTableListCtrl::GetContainerType() const noexcept { return m_ulContainerType; }

	_Check_return_ bool CContentsTableListCtrl::IsContentsTableSet() const noexcept
	{
		return m_lpContentsTable != nullptr;
	}

	void CContentsTableListCtrl::SetContentsTable(
		_In_opt_ LPMAPITABLE lpContentsTable,
		tableDisplayFlags displayFlags,
		ULONG ulContainerType)
	{
		// If nothing to do, exit early
		if (m_bInLoadOp || lpContentsTable == m_lpContentsTable) return;

		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		m_displayFlags = displayFlags;
		m_ulContainerType = ulContainerType;

		output::DebugPrintEx(
			output::dbgLevel::Generic,
			CLASS,
			L"SetContentsTable",
			L"replacing %p with %p\n",
			m_lpContentsTable,
			lpContentsTable);
		output::DebugPrintEx(
			output::dbgLevel::Generic, CLASS, L"SetContentsTable", L"New container type: 0x%X\n", m_ulContainerType);
		// Clean up the old contents table and grab the new one
		if (m_lpContentsTable)
		{
			// If we don't Unadvise before releasing our reference, we'll leak an advise sink
			NotificationOff();

			m_lpContentsTable->Release();
			m_lpContentsTable = nullptr;
		}

		m_lpContentsTable = lpContentsTable;
		if (m_lpContentsTable) m_lpContentsTable->AddRef();

		// Set up the columns on the new contents table and refresh!
		DoSetColumns(true, registry::editColumnsOnLoad);
	}

	void CContentsTableListCtrl::GetStatus()
	{
		if (!IsContentsTableSet()) return;

		ULONG ulTableStatus = NULL;
		ULONG ulTableType = NULL;

		const auto hRes = EC_MAPI(m_lpContentsTable->GetStatus(&ulTableStatus, &ulTableType));

		if (SUCCEEDED(hRes))
		{
			dialog::editor::CEditor MyData(this, IDS_GETSTATUS, IDS_GETSTATUSPROMPT, CEDITOR_BUTTON_OK);
			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_ULTABLESTATUS, true));
			MyData.SetHex(0, ulTableStatus);
			auto szFlags = flags::InterpretFlags(flagTableStatus, ulTableStatus);
			MyData.AddPane(viewpane::TextPane::CreateMultiLinePane(1, IDS_ULTABLESTATUS, szFlags, true));
			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_ULTABLETYPE, true));
			MyData.SetHex(2, ulTableType);
			szFlags = flags::InterpretFlags(flagTableType, ulTableType);
			MyData.AddPane(viewpane::TextPane::CreateMultiLinePane(3, IDS_ULTABLETYPE, szFlags, true));

			static_cast<void>(MyData.DisplayDialog());
		}
	}

	// Takes a tag array and builds the UI out of it - does NOT touch the table
	void CContentsTableListCtrl::SetUIColumns(_In_ LPSPropTagArray lpTags)
	{
		if (!lpTags) return;

		// find a PR_DISPLAY_NAME column for later use
		m_ulDisplayNameColumn = NODISPLAYNAME;
		for (ULONG i = 0; i < lpTags->cValues; i++)
		{
			if (PROP_ID(mapi::getTag(lpTags, i)) == PROP_ID(PR_DISPLAY_NAME))
			{
				m_ulDisplayNameColumn = i;
				break;
			}
		}

		// Didn't find display name - fall back on some other columns
		if (NODISPLAYNAME == m_ulDisplayNameColumn)
		{
			for (ULONG i = 0; i < lpTags->cValues; i++)
			{
				if (PROP_ID(mapi::getTag(lpTags, i)) == PROP_ID(PR_SUBJECT) ||
					PROP_ID(mapi::getTag(lpTags, i)) == PROP_ID(PR_RULE_NAME) ||
					PROP_ID(mapi::getTag(lpTags, i)) == PROP_ID(PR_MEMBER_NAME) ||
					PROP_ID(mapi::getTag(lpTags, i)) == PROP_ID(PR_ATTACH_LONG_FILENAME) ||
					PROP_ID(mapi::getTag(lpTags, i)) == PROP_ID(PR_ATTACH_FILENAME))
				{
					m_ulDisplayNameColumn = i;
					break;
				}
			}
		}

		output::DebugPrintEx(
			output::dbgLevel::Generic, CLASS, L"SetColumns", L"calculating and inserting column headers\n");
		MySetRedraw(false);

		// Delete all of the old column headers
		DeleteAllColumns();

		AddColumns(lpTags);

		AutoSizeColumns(true);

		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"SetColumns", L"Done inserting column headers\n");

		MySetRedraw(true);
	}

	void CContentsTableListCtrl::DoSetColumns(bool bAddExtras, bool bDisplayEditor)
	{
		output::DebugPrintEx(
			output::dbgLevel::Generic, CLASS, L"DoSetColumns", L"bDisplayEditor = %d\n", bDisplayEditor);

		if (!IsContentsTableSet())
		{
			// Clear out the selected item view since we have no contents table
			if (m_lpHostDlg) m_lpHostDlg->OnUpdateSingleMAPIPropListCtrl(nullptr, nullptr);

			// Make sure we're clear
			DeleteAllColumns();
			ModifyStyle(0, LVS_NOCOLUMNHEADER);
			RefreshTable();
			return;
		}

		// these arrays get allocated during the func and need to be freed
		LPSPropTagArray lpConcatTagArray = nullptr;
		LPSPropTagArray lpModifiedTags = nullptr;
		LPSPropTagArray lpOriginalColSet = nullptr;

		auto bModified = false;

		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		EC_MAPI_S(m_lpContentsTable->QueryColumns(NULL, &lpOriginalColSet));
		// this is just a pointer - do not free
		auto lpFinalTagArray = lpOriginalColSet;

		if (bAddExtras)
		{
			// build an array with the source set and m_sptDefaultDisplayColumnTags combined
			lpConcatTagArray = mapi::ConcatSPropTagArrays(
				m_sptDefaultDisplayColumnTags,
				lpFinalTagArray); // build on the final array we've computed thus far
			lpFinalTagArray = lpConcatTagArray;
		}

		if (bDisplayEditor)
		{
			LPMDB lpMDB = nullptr;
			if (m_lpMapiObjects)
			{
				lpMDB = m_lpMapiObjects->GetMDB(); // do not release
			}

			dialog::editor::CTagArrayEditor MyEditor(
				this,
				IDS_COLUMNSET,
				IDS_COLUMNSETPROMPT,
				m_lpContentsTable,
				lpFinalTagArray, // build on the final array we've computed thus far
				m_bIsAB,
				lpMDB);

			if (MyEditor.DisplayDialog())
			{
				lpModifiedTags = MyEditor.DetachModifiedTagArray();
				if (lpModifiedTags)
				{
					lpFinalTagArray = lpModifiedTags;
					bModified = true;
				}
			}
		}
		else
		{
			// Apply lpFinalTagArray through SetColumns
			EC_MAPI_S(m_lpContentsTable->SetColumns(lpFinalTagArray, TBL_BATCH));
			bModified = true;
		}

		if (bModified)
		{
			// Cycle our notification, turning off the old one if necessary
			NotificationOff();
			NotificationOn();

			SetUIColumns(lpFinalTagArray);
			RefreshTable();
		}

		MAPIFreeBuffer(lpModifiedTags);
		MAPIFreeBuffer(lpConcatTagArray);
		MAPIFreeBuffer(lpOriginalColSet);
	}

	void
	CContentsTableListCtrl::AddColumn(UINT uidHeaderName, ULONG ulCurHeaderCol, ULONG ulCurTagArrayRow, ULONG ulPropTag)
	{
		HDITEM hdItem = {};
		auto lpMyHeader = GetHeaderCtrl();
		std::wstring szHeaderString;
		LPMDB lpMDB = nullptr;
		if (m_lpMapiObjects) lpMDB = m_lpMapiObjects->GetMDB(); // do not release

		if (uidHeaderName)
		{
			szHeaderString = strings::loadstring(uidHeaderName);
		}
		else
		{
			const auto propTagNames = proptags::PropTagToPropName(ulPropTag, m_bIsAB);
			szHeaderString = propTagNames.bestGuess;
			if (szHeaderString.empty())
			{
				const auto namePropNames = cache::NameIDToStrings(ulPropTag, lpMDB, nullptr, nullptr, m_bIsAB);

				szHeaderString = namePropNames.name;
			}

			if (szHeaderString.empty())
			{
				szHeaderString = strings::format(L"0x%08X", ulPropTag); // STRING_OK
			}
		}

		const auto iRetVal = InsertColumnW(ulCurHeaderCol, szHeaderString);

		if (-1 == iRetVal)
		{
			// We failed to insert a column header
			error::ErrDialog(__FILE__, __LINE__, IDS_EDCOLUMNHEADERFAILED);
		}

		if (lpMyHeader)
		{
			hdItem.mask = HDI_LPARAM;
			auto lpHeaderData = new (std::nothrow) HeaderData; // Will be deleted in CSortListCtrl::DeleteAllColumns
			if (lpHeaderData)
			{
				lpHeaderData->ulTagArrayRow = ulCurTagArrayRow;
				lpHeaderData->ulPropTag = ulPropTag;
				lpHeaderData->bIsAB = m_bIsAB;
				lpHeaderData->szHeaderString = szHeaderString;
				lpHeaderData->szTipString = proptags::TagToString(ulPropTag, lpMDB, m_bIsAB, false);

				hdItem.lParam = reinterpret_cast<LPARAM>(lpHeaderData);
				EC_B_S(lpMyHeader->SetItem(ulCurHeaderCol, &hdItem));
			}
		}
	}

	// Sets up column headers based on passed in named columns
	// Put all named columns first, followed by a column for each property in the contents table
	void CContentsTableListCtrl::AddColumns(_In_ LPSPropTagArray lpCurColTagArray)
	{
		if (!lpCurColTagArray || !m_lpHostDlg) return;

		m_ulHeaderColumns = lpCurColTagArray->cValues;

		ULONG ulCurHeaderCol = 0;
		if (registry::doColumnNames)
		{
			output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"AddColumns", L"Adding named columns\n");
			// If we have named columns, put them up front

			// Walk through the list of default display columns and add them to our header list
			for (const auto& displayCol : m_lpDefaultDisplayColumns)
			{
				ULONG ulCurTagArrayRow = 0;
				if (mapi::FindPropInPropTagArray(
						lpCurColTagArray,
						mapi::getTag(m_sptDefaultDisplayColumnTags, displayCol.ulMatchingTableColumn),
						&ulCurTagArrayRow))
				{
					AddColumn(
						displayCol.uidName,
						ulCurHeaderCol,
						ulCurTagArrayRow,
						mapi::getTag(lpCurColTagArray, ulCurTagArrayRow));
					// Strike out the value in the tag array so we can ignore it later!
					mapi::setTag(lpCurColTagArray, ulCurTagArrayRow) = NULL;

					ulCurHeaderCol++;
				}
			}
		}

		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"AddColumns", L"Adding unnamed columns\n");
		// Now, walk through the current tag table and add each unstruck column to our list
		for (ULONG ulCurTableCol = 0; ulCurTableCol < lpCurColTagArray->cValues; ulCurTableCol++)
		{
			if (mapi::getTag(lpCurColTagArray, ulCurTableCol) != NULL)
			{
				AddColumn(NULL, ulCurHeaderCol, ulCurTableCol, mapi::getTag(lpCurColTagArray, ulCurTableCol));
				ulCurHeaderCol++;
			}
		}

		if (ulCurHeaderCol)
		{
			ModifyStyle(LVS_NOCOLUMNHEADER, 0);
		}

		// this would be bad
		if (ulCurHeaderCol < m_ulHeaderColumns)
		{
			error::ErrDialog(__FILE__, __LINE__, IDS_EDTOOMANYCOLUMNS);
		}

		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"AddColumns", L"Done adding columns\n");
	}

	void CContentsTableListCtrl::SetRestriction(_In_opt_ const _SRestriction* lpRes) noexcept
	{
		MAPIFreeBuffer(const_cast<LPSRestriction>(m_lpRes));
		m_lpRes = lpRes;
	}

	_Check_return_ const _SRestriction* CContentsTableListCtrl::GetRestriction() const noexcept { return m_lpRes; }

	_Check_return_ restrictionType CContentsTableListCtrl::GetRestrictionType() const noexcept
	{
		return m_RestrictionType;
	}

	void CContentsTableListCtrl::SetRestrictionType(restrictionType RestrictionType) noexcept
	{
		m_RestrictionType = RestrictionType;
	}

	_Check_return_ HRESULT CContentsTableListCtrl::ApplyRestriction() const
	{
		if (!m_lpContentsTable) return MAPI_E_INVALID_PARAMETER;

		auto hRes = S_OK;
		output::DebugPrintEx(
			output::dbgLevel::Generic, CLASS, L"ApplyRestriction", L"m_RestrictionType = 0x%X\n", m_RestrictionType);
		// Apply our restrictions
		if (m_RestrictionType == restrictionType::normal)
		{
			output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"ApplyRestriction", L"applying restriction:\n");

			if (m_lpMapiObjects)
			{
				const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
				output::outputRestriction(output::dbgLevel::Generic, nullptr, m_lpRes, lpMDB);
			}

			hRes = EC_MAPI(m_lpContentsTable->Restrict(const_cast<LPSRestriction>(m_lpRes), TBL_BATCH));
		}
		else
		{
			hRes = WC_H_MSG(IDS_TABLENOSUPPORTRES, m_lpContentsTable->Restrict(nullptr, TBL_BATCH));
		}

		return hRes;
	}

	// Idea here is to do our MAPI work here on this thread, then send messages (SendMessage) back to the control to add the data to the view
	// This way, control functions only happen on the main thread
	// ::SendMessage will be handled on main thread, but block until the call returns.
	// This is the ideal behavior for this worker thread.
	void ThreadFuncLoadTable(const HWND hWndHost, CContentsTableListCtrl* lpListCtrl, LPMAPITABLE lpContentsTable)
	{
		if (!lpListCtrl || !lpContentsTable || FAILED(EC_MAPI(MAPIInitialize(nullptr))))
		{
			if (lpContentsTable) lpContentsTable->Release();
			if (lpListCtrl) lpListCtrl->Release();
			lpListCtrl->ClearLoading();
			return;
		}

		ULONG ulTotal = 0;
		const ULONG ulThrottleLevel = registry::throttleLevel;
		const auto rowCount = ulThrottleLevel ? ulThrottleLevel : CContentsTableListCtrl::NUMROWSPERLOOP;
		LPSRowSet pRows = nullptr;
		ULONG iCurListBoxRow = 0;

		static_cast<void>(::SendMessage(hWndHost, WM_MFCMAPI_CLEARSINGLEMAPIPROPLIST, NULL, NULL));
		auto szCount = std::to_wstring(lpListCtrl->GetItemCount());
		dialog::CBaseDialog::UpdateStatus(
			hWndHost, statusPane::data1, strings::formatmessage(IDS_STATUSTEXTNUMITEMS, szCount.c_str()));

		// potentially lengthy op - check abort before and after
		if (!lpListCtrl->bAbortLoad())
		{
			WC_H_S(lpListCtrl->ApplyRestriction());
		}

		if (!lpListCtrl->bAbortLoad()) // only check abort once for this group of ops
		{
			// go to the first row
			EC_MAPI_S(lpContentsTable->SeekRow(BOOKMARK_BEGINNING, 0, nullptr));

			EC_MAPI_S(lpContentsTable->GetRowCount(NULL, &ulTotal));

			output::DebugPrintEx(
				output::dbgLevel::Generic, CLASS, L"ThreadFuncLoadTable", L"ulTotal = 0x%X\n", ulTotal);

			if (ulTotal)
			{
				dialog::CBaseDialog::UpdateStatus(
					hWndHost, statusPane::data2, strings::formatmessage(IDS_LOADINGITEMS, 0, ulTotal));
			}
		}

		const auto lpRes = lpListCtrl->GetRestriction();
		const auto resType = lpListCtrl->GetRestrictionType();
		// get rows and add them to the list
		while (true)
		{
			if (lpListCtrl->bAbortLoad()) break;
			auto hRes = S_OK;
			dialog::CBaseDialog::UpdateStatus(hWndHost, statusPane::infoText, strings::loadstring(IDS_ESCSTOPLOADING));
			if (pRows) FreeProws(pRows);
			pRows = nullptr;
			if (resType == restrictionType::findrow && lpRes)
			{
				output::DebugPrintEx(
					output::dbgLevel::Generic, CLASS, L"DoFindRows", L"running FindRow with restriction:\n");
				output::outputRestriction(output::dbgLevel::Generic, nullptr, lpRes, nullptr);

				if (lpListCtrl->bAbortLoad()) break;
				hRes = WC_MAPI(lpContentsTable->FindRow(const_cast<LPSRestriction>(lpRes), BOOKMARK_CURRENT, NULL));
				if (FAILED(hRes)) break;

				if (lpListCtrl->bAbortLoad()) break;
				hRes = EC_MAPI(lpContentsTable->QueryRows(1, NULL, &pRows));
			}
			else
			{
				output::DebugPrintEx(
					output::dbgLevel::Generic,
					CLASS,
					L"ThreadFuncLoadTable",
					L"Calling QueryRows. Asking for 0x%X rows.\n",
					rowCount);
				// Pull back a sizable block of rows to add to the list box
				if (lpListCtrl->bAbortLoad()) break;
				hRes = EC_MAPI(lpContentsTable->QueryRows(rowCount, NULL, &pRows));
			}

			if (FAILED(hRes) || !pRows || !pRows->cRows) break;

			output::DebugPrintEx(
				output::dbgLevel::Generic, CLASS, L"ThreadFuncLoadTable", L"Got this many rows: 0x%X\n", pRows->cRows);

			for (ULONG iCurPropRow = 0; iCurPropRow < pRows->cRows; iCurPropRow++)
			{
				if (lpListCtrl->bAbortLoad()) break;
				if (ulTotal)
				{
					dialog::CBaseDialog::UpdateStatus(
						hWndHost,
						statusPane::data2,
						strings::formatmessage(IDS_LOADINGITEMS, iCurListBoxRow + 1, ulTotal));
				}

				output::DebugPrintEx(
					output::dbgLevel::Generic,
					CLASS,
					L"ThreadFuncLoadTable",
					L"Asking to add %p to %u\n",
					&pRows->aRow[iCurPropRow],
					iCurListBoxRow);
				hRes = WC_H(static_cast<HRESULT>(::SendMessage(
					lpListCtrl->m_hWnd,
					WM_MFCMAPI_THREADADDITEM,
					iCurListBoxRow,
					reinterpret_cast<LPARAM>(&pRows->aRow[iCurPropRow]))));
				if (FAILED(hRes)) continue;

				// We've handed this row off to the control, so wipe it from the array.
				pRows->aRow[iCurPropRow].cValues = 0;
				pRows->aRow[iCurPropRow].lpProps = nullptr;
				iCurListBoxRow++;
			}

			// We've removed some or all rows from the array. FreeProws will free the rest, as well as the parent array.
			FreeProws(pRows);
			pRows = nullptr;

			if (ulThrottleLevel && iCurListBoxRow >= ulThrottleLevel)
				break; // Only render ulThrottleLevel rows if throttle is on
		}

		if (lpListCtrl->bAbortLoad())
		{
			dialog::CBaseDialog::UpdateStatus(
				hWndHost, statusPane::infoText, strings::loadstring(IDS_TABLELOADCANCELLED));
		}
		else
		{
			dialog::CBaseDialog::UpdateStatus(hWndHost, statusPane::infoText, strings::loadstring(IDS_TABLELOADED));
		}

		dialog::CBaseDialog::UpdateStatus(hWndHost, statusPane::data2, strings::emptystring);
		output::DebugPrintEx(
			output::dbgLevel::Generic, CLASS, L"ThreadFuncLoadTable", L"added %u items\n", iCurListBoxRow);
		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"ThreadFuncLoadTable", L"Releasing pointers.\n");

		// Bunch of cleanup
		if (pRows) FreeProws(pRows);
		if (lpContentsTable) lpContentsTable->Release();
		if (lpListCtrl) lpListCtrl->Release();
		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"ThreadFuncLoadTable", L"Pointers released.\n");

		MAPIUninitialize();

		lpListCtrl->ClearLoading();
	}

	_Check_return_ bool CContentsTableListCtrl::IsLoading() const noexcept { return m_bInLoadOp; }

	void CContentsTableListCtrl::ClearLoading() noexcept { m_bInLoadOp = false; }

	void CContentsTableListCtrl::LoadContentsTableIntoView()
	{
		if (m_bInLoadOp || !m_lpHostDlg) return;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.
		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"LoadContentsTableIntoView", L"\n");

		// Ensure we're not currently loading
		OnCancelTableLoad();

		EC_B_S(DeleteAllItems());

		if (!m_lpContentsTable) return;

		m_bInLoadOp = true;
		// Do not call return after this point!

		// Addref the objects we're passing to the thread
		// ThreadFuncLoadTable should release them before exiting
		this->AddRef();
		m_lpContentsTable->AddRef();

		output::DebugPrintEx(
			output::dbgLevel::Generic, CLASS, L"LoadContentsTableIntoView", L"Creating load thread.\n");

		auto loadThread = std::thread(ThreadFuncLoadTable, m_lpHostDlg->m_hWnd, this, m_lpContentsTable);

		m_LoadThreadHandle.swap(loadThread);
	}

	void CContentsTableListCtrl::OnCancelTableLoad()
	{
		// Signal the abort
		// This is the only function which ever sets this flag.
		output::DebugPrintEx(
			output::dbgLevel::Generic,
			CLASS,
			L"OnCancelTableLoad",
			L"Setting abort flag and waiting for thread to discover it\n");
		InterlockedExchange(&m_bAbortLoad, true);

		CWaitCursor Wait; // Change the mouse to an hourglass while we work.
		auto bVKF5Hit = false;

		// Pump messages until the thread signals we're out of the load op (which is the last thing we do on the thread).
		// Don't pmp F5. Just remember those and we'll send them up later.
		while (m_bInLoadOp)
		{
			auto msg = MSG{};
			if (PeekMessage(&msg, m_hWnd, 0, 0, PM_REMOVE))
			{
				if (msg.message == WM_KEYDOWN && msg.wParam == VK_F5)
				{
					output::DebugPrintEx(
						output::dbgLevel::Generic, CLASS, L"OnCancelTableLoad", L"Ditching refresh (F5)\n");
					bVKF5Hit = true;
				}
				else
				{
					DispatchMessage(&msg);
				}
			}
		}

		// Now wait for the thread to actually shut down.
		if (m_LoadThreadHandle.joinable()) m_LoadThreadHandle.join();

		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"OnCancelTableLoad", L"Load thread has shut down.\n");

		// Finally, reset the abort so we're ready to load again (if needed)
		InterlockedExchange(&m_bAbortLoad, false);

		if (bVKF5Hit) // If we ditched a refresh message, repost it now
		{
			output::DebugPrintEx(
				output::dbgLevel::Generic, CLASS, L"OnCancelTableLoad", L"Posting skipped refresh message\n");
			PostMessage(WM_KEYDOWN, VK_F5, 0);
		}
	}

	void CContentsTableListCtrl::SetRowStrings(int iRow, _In_ LPSRow lpsRowData)
	{
		if (!lpsRowData) return;

		const auto lpMyHeader = GetHeaderCtrl();

		if (!lpMyHeader) return;

		for (ULONG iColumn = 0; iColumn < m_ulHeaderColumns; iColumn++)
		{
			HDITEM hdItem = {};
			hdItem.mask = HDI_LPARAM;
			EC_B_S(lpMyHeader->GetItem(iColumn, &hdItem));

			if (hdItem.lParam)
			{
				const auto ulCol = reinterpret_cast<LPHEADERDATA>(hdItem.lParam)->ulTagArrayRow;

				if (ulCol < lpsRowData->cValues)
				{
					std::wstring PropString;
					const auto pProp = &lpsRowData->lpProps[ulCol];

					// If we've got a MAPI_E_NOT_FOUND error, just don't display it.
					if (registry::suppressNotFound && pProp && PROP_TYPE(pProp->ulPropTag) == PT_ERROR &&
						pProp->Value.err == MAPI_E_NOT_FOUND)
					{
						if (0 == iColumn)
						{
							SetItemText(iRow, iColumn, L"");
						}

						continue;
					}

					property::parseProperty(pProp, &PropString, nullptr);

					auto szFlags = smartview::InterpretNumberAsString(
						pProp->Value, pProp->ulPropTag, NULL, nullptr, nullptr, false);
					if (!szFlags.empty())
					{
						PropString += L" ("; // STRING_OK
						PropString += szFlags;
						PropString += L")"; // STRING_OK
					}

					SetItemText(iRow, iColumn, PropString);
				}
				else
				{
					// This is an odd case which just shouldn't happen.
					// If SetColumns failed in DoSetColumns, we might have columns
					// mapped past the end of the table. Just log the error and give up.
					WARNHRESMSG(MAPI_E_NOT_FOUND, IDS_COLOUTOFRANGE);
					break;
				}
			}
		}
	}

	ULONG GetDepth(_In_ LPSRow lpsRowData) noexcept
	{
		if (!lpsRowData) return 0;

		auto ulDepth = 0;
		const auto lpDepth = PpropFindProp(lpsRowData->lpProps, lpsRowData->cValues, PR_DEPTH);
		if (lpDepth && PR_DEPTH == lpDepth->ulPropTag) ulDepth = lpDepth->Value.l;
		if (ulDepth > 5) ulDepth = 5; // Just in case

		return ulDepth;
	}

	static TypeIcon _ObjTypeIcons[] = {
		{MAPI_STORE, sortIcon::mapiStore},
		{MAPI_ADDRBOOK, sortIcon::mapiAddrBook},
		{MAPI_FOLDER, sortIcon::mapiFolder},
		{MAPI_ABCONT, sortIcon::mapiAbcont},
		{MAPI_MESSAGE, sortIcon::mapiMessage},
		{MAPI_MAILUSER, sortIcon::mapiMailUser},
		{MAPI_ATTACH, sortIcon::mapiAttach},
		{MAPI_DISTLIST, sortIcon::mapiDistList},
		{MAPI_PROFSECT, sortIcon::mapiProfSect},
		{MAPI_STATUS, sortIcon::mapiStatus},
		{MAPI_SESSION, sortIcon::mapiSession},
		{MAPI_FORMINFO, sortIcon::mapiFormInfo},
	};

	sortIcon GetImage(_In_ LPSRow lpsRowData) noexcept
	{
		if (!lpsRowData) return sortIcon::siDefault;

		sortIcon ulImage = sortIcon::siDefault;

		const auto lpRowType = PpropFindProp(lpsRowData->lpProps, lpsRowData->cValues, PR_ROW_TYPE);
		if (lpRowType && PR_ROW_TYPE == lpRowType->ulPropTag)
		{
			switch (lpRowType->Value.l)
			{
			case TBL_LEAF_ROW:
				break;
			case TBL_EMPTY_CATEGORY:
			case TBL_COLLAPSED_CATEGORY:
				ulImage = sortIcon::nodeCollapsed;
				break;
			case TBL_EXPANDED_CATEGORY:
				ulImage = sortIcon::nodeExpanded;
				break;
			}
		}

		if (sortIcon::siDefault == ulImage)
		{
			const auto lpObjType = PpropFindProp(lpsRowData->lpProps, lpsRowData->cValues, PR_OBJECT_TYPE);
			if (lpObjType && PR_OBJECT_TYPE == lpObjType->ulPropTag)
			{
				for (const auto& _ObjTypeIcon : _ObjTypeIcons)
				{
					if (_ObjTypeIcon.objType == lpObjType->Value.ul)
					{
						ulImage = _ObjTypeIcon.image;
						break;
					}
				}
			}
		}

		// We still don't have a good icon - make some heuristic guesses
		if (ulImage == sortIcon::siDefault)
		{
			auto lpProp = PpropFindProp(lpsRowData->lpProps, lpsRowData->cValues, PR_SERVICE_UID);
			if (!lpProp)
			{
				lpProp = PpropFindProp(lpsRowData->lpProps, lpsRowData->cValues, PR_PROVIDER_UID);
			}

			if (lpProp)
			{
				ulImage = sortIcon::mapiProfSect;
			}
		}

		return ulImage;
	}

	void CContentsTableListCtrl::RefreshItem(int iRow, _In_ LPSRow lpsRowData, bool bItemExists)
	{
		sortlistdata::sortListData* lpData = nullptr;

		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"RefreshItem", L"item %d\n", iRow);

		if (bItemExists)
		{
			lpData = GetSortListData(iRow);
		}
		else
		{
			const auto ulDepth = GetDepth(lpsRowData);
			const auto ulImage = GetImage(lpsRowData);

			lpData = InsertRow(iRow, L"TempRefreshItem", ulDepth, ulImage); // STRING_OK
		}

		if (lpData)
		{
			sortlistdata::contentsData::init(lpData, lpsRowData);

			SetRowStrings(iRow, lpsRowData);
			// Do this last so that our row can't get sorted before we're done!
			lpData->setFullyLoaded(true);
		}
	}

	// Crack open the given SPropValue and render it to the given row in the list.
	void CContentsTableListCtrl::AddItemToListBox(int iRow, _In_ LPSRow lpsRowToAdd)
	{
		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"AddItemToListBox", L"item %d\n", iRow);

		RefreshItem(iRow, lpsRowToAdd, false);

		if (m_lpHostDlg) m_lpHostDlg->UpdateStatusBarText(statusPane::data1, IDS_STATUSTEXTNUMITEMS, GetItemCount());
	}

	void CContentsTableListCtrl::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
	{
		output::DebugPrintEx(output::dbgLevel::Menu, CLASS, L"OnKeyDown", L"0x%X\n", nChar);

		if (!m_lpHostDlg) return;
		const auto bCtrlPressed = GetKeyState(VK_CONTROL) < 0;
		const auto bShiftPressed = GetKeyState(VK_SHIFT) < 0;
		const auto bMenuPressed = GetKeyState(VK_MENU) < 0;

		if (!bMenuPressed)
		{
			if ('C' == nChar && bCtrlPressed && !bShiftPressed)
			{
				// Handle copy to clipboard here
				CopyRows();
			}
		}

		if (!bMenuPressed)
		{
			if ('A' == nChar && bCtrlPressed)
			{
				SelectAll();
			}
			else if (VK_RETURN == nChar && S_FALSE != DoExpandCollapse())
			{
				// nothing to do - DoExpandCollapse did the work
				// we only want to go to the next case if it returned S_FALSE
			}
			else if (!m_lpHostDlg || !m_lpHostDlg->HandleKeyDown(nChar, bShiftPressed, bCtrlPressed, bMenuPressed))
			{
				CSortListCtrl::OnKeyDown(nChar, nRepCnt, nFlags);
			}
		}
	}

	void CContentsTableListCtrl::CopyRows() const
	{
		const auto iNumItems = GetSelectedCount();

		if (!iNumItems) return;
		auto iSelectedItem = -1;
		const auto lpMyHeader = GetHeaderCtrl();
		auto str = std::wstring{};

		for (UINT iArrayPos = 0; iArrayPos < iNumItems; iArrayPos++)
		{
			iSelectedItem = GetNextItem(iSelectedItem, LVNI_SELECTED);
			if (-1 != iSelectedItem)
			{
				str += strings::format(L"Row: %d\r\n", iSelectedItem);

				for (ULONG iColumn = 0; iColumn < m_ulHeaderColumns; iColumn++)
				{
					auto hdItem = HDITEM{};
					hdItem.mask = HDI_LPARAM;
					EC_B_S(lpMyHeader->GetItem(iColumn, &hdItem));

					if (hdItem.lParam)
					{
						const auto headerData = reinterpret_cast<LPHEADERDATA>(hdItem.lParam);
						const auto prop = GetItemText(iSelectedItem, iColumn);
						str += headerData->szHeaderString + L"::" + prop + L"\r\n";
					}
				}
			}
		}

		clipboard::CopyTo(str);
	}

	_Check_return_ LPENTRYLIST CContentsTableListCtrl::GetSelectedItemEIDs() const
	{
		const auto iNumItems = GetSelectedCount();

		if (!iNumItems) return nullptr;
		if (iNumItems > ULONG_MAX / sizeof(SBinary)) return nullptr;

		const auto lpTempList = mapi::allocate<LPENTRYLIST>(sizeof(ENTRYLIST));
		if (lpTempList)
		{
			lpTempList->cValues = iNumItems;
			lpTempList->lpbin = mapi::allocate<LPSBinary>(static_cast<ULONG>(sizeof(SBinary)) * iNumItems, lpTempList);
			if (lpTempList->lpbin)
			{
				auto iSelectedItem = -1;

				for (UINT iArrayPos = 0; iArrayPos < iNumItems; iArrayPos++)
				{
					lpTempList->lpbin[iArrayPos].cb = 0;
					lpTempList->lpbin[iArrayPos].lpb = nullptr;
					iSelectedItem = GetNextItem(iSelectedItem, LVNI_SELECTED);
					if (-1 != iSelectedItem)
					{
						const auto lpData = GetSortListData(iSelectedItem);
						if (lpData)
						{
							const auto contents = lpData->cast<sortlistdata::contentsData>();
							if (contents && contents->getEntryID())
							{
								lpTempList->lpbin[iArrayPos] = mapi::CopySBinary(*contents->getEntryID(), lpTempList);
							}
						}
					}
				}
			}
		}

		return lpTempList;
	}

	_Check_return_ std::vector<int> CContentsTableListCtrl::GetSelectedItemNums() const
	{
		auto iItem = -1;
		std::vector<int> iItems;
		do
		{
			iItem = GetNextItem(iItem, LVNI_SELECTED);
			if (iItem != -1)
			{
				iItems.push_back(iItem);
				output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"GetSelectedItemNums", L"iItem: 0x%X\n", iItem);
			}
		} while (iItem != -1);

		return iItems;
	}

	_Check_return_ std::vector<sortlistdata::sortListData*> CContentsTableListCtrl::GetSelectedItemData() const
	{
		auto iItem = -1;
		std::vector<sortlistdata::sortListData*> items;
		do
		{
			iItem = GetNextItem(iItem, LVNI_SELECTED);
			if (iItem != -1)
			{
				items.push_back(GetSortListData(iItem));
			}
		} while (iItem != -1);

		return items;
	}

	_Check_return_ sortlistdata::sortListData* CContentsTableListCtrl::GetFirstSelectedItemData() const
	{
		const auto iItem = GetNextItem(-1, LVNI_SELECTED);
		if (-1 == iItem) return nullptr;

		return GetSortListData(iItem);
	}

	// Pass iCurItem as -1 to get the primary selected item.
	// Call again with the previous iCurItem to get the next one.
	// Stop calling when iCurItem = -1 and/or lppProp is NULL
	// If iCurItem is NULL, just returns the focused item
	_Check_return_ int CContentsTableListCtrl::GetNextSelectedItemNum(_Inout_opt_ int* iCurItem) const
	{
		auto iItem = NULL;

		if (iCurItem) // intentionally not dereffing - checking to see if NULL was actually passed
		{
			iItem = *iCurItem;
		}
		else
		{
			iItem = -1;
		}

		output::DebugPrintEx(
			output::dbgLevel::Generic, CLASS, L"GetNextSelectedItemNum", L"iItem before = 0x%X\n", iItem);

		iItem = GetNextItem(iItem, LVNI_SELECTED);

		output::DebugPrintEx(
			output::dbgLevel::Generic, CLASS, L"GetNextSelectedItemNum", L"iItem after = 0x%X\n", iItem);

		if (iCurItem) *iCurItem = iItem;

		return iItem;
	}

	_Check_return_ sortlistdata::sortListData* CContentsTableListCtrl::GetSortListData(int iItem) const
	{
		return reinterpret_cast<sortlistdata::sortListData*>(GetItemData(iItem));
	}

	// Pass iCurItem as -1 to get the primary selected item.
	// Call again with the previous iCurItem to get the next one.
	// Stop calling when iCurItem = -1 and/or lppProp is NULL
	// If iCurItem is NULL, just returns the focused item
	_Check_return_ LPMAPIPROP
	CContentsTableListCtrl::OpenNextSelectedItemProp(_Inout_opt_ int* iCurItem, modifyType bModify) const
	{
		const auto iItem = GetNextSelectedItemNum(iCurItem);
		if (-1 != iItem)
		{
			return m_lpHostDlg->OpenItemProp(iItem, bModify);
		}

		return nullptr;
	}

	_Check_return_ LPMAPIPROP CContentsTableListCtrl::DefaultOpenItemProp(int iItem, modifyType bModify) const
	{
		if (!m_lpMapiObjects || -1 == iItem) return nullptr;

		output::DebugPrintEx(
			output::dbgLevel::Generic,
			CLASS,
			L"DefaultOpenItemProp",
			L"iItem = %d, bModify = %d, m_ulContainerType = 0x%X\n",
			iItem,
			bModify,
			m_ulContainerType);

		const auto lpListData = GetSortListData(iItem);
		if (!lpListData) return nullptr;

		const auto contents = lpListData->cast<sortlistdata::contentsData>();
		if (!contents) return nullptr;

		const auto lpEID = contents->getEntryID();
		if (!lpEID || lpEID->cb == 0) return nullptr;

		output::DebugPrint(output::dbgLevel::Generic, L"Item being opened:\n");
		output::outputBinary(output::dbgLevel::Generic, nullptr, *lpEID);

		// Find the highlighted item EID
		LPMAPIPROP lpMAPIProp = nullptr;
		switch (m_ulContainerType)
		{
		case MAPI_ABCONT:
		{
			const auto lpAB = m_lpMapiObjects->GetAddrBook(false); // do not release
			lpMAPIProp = mapi::CallOpenEntry<LPMAPIPROP>(
				nullptr,
				lpAB, // use AB
				nullptr,
				nullptr,
				lpEID,
				nullptr,
				bModify == modifyType::REQUEST_MODIFY ? MAPI_MODIFY : MAPI_BEST_ACCESS,
				nullptr);
		}

		break;
		case MAPI_FOLDER:
		{
			const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
			LPCIID lpInterface = nullptr;

			if (registry::useMessageRaw)
			{
				lpInterface = &IID_IMessageRaw;
			}

			lpMAPIProp = mapi::CallOpenEntry<LPMAPIPROP>(
				lpMDB, // use MDB
				nullptr,
				nullptr,
				nullptr,
				lpEID,
				lpInterface,
				bModify == modifyType::REQUEST_MODIFY ? MAPI_MODIFY : MAPI_BEST_ACCESS,
				nullptr);
		}

		break;
		default:
		{
			const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
			lpMAPIProp = mapi::CallOpenEntry<LPMAPIPROP>(
				nullptr,
				nullptr,
				nullptr,
				lpMAPISession, // use session
				lpEID,
				nullptr,
				bModify == modifyType::REQUEST_MODIFY ? MAPI_MODIFY : MAPI_BEST_ACCESS,
				nullptr);
		}

		break;
		}

		if (!lpMAPIProp && modifyType::REQUEST_MODIFY == bModify)
		{
			output::DebugPrint(output::dbgLevel::Generic, L"\tOpenEntry failed. Will try again without MAPI_MODIFY\n");
			// We got access denied when we passed MAPI_MODIFY
			// Let's try again without it.
			lpMAPIProp = DefaultOpenItemProp(iItem, modifyType::DO_NOT_REQUEST_MODIFY);
		}

		output::DebugPrintEx(
			output::dbgLevel::Generic, CLASS, L"DefaultOpenItemProp", L"returning lpMAPIProp = %p\n", lpMAPIProp);

		return lpMAPIProp;
	}

	void CContentsTableListCtrl::SelectAll()
	{
		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"SelectAll", L"\n");
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.
		MySetRedraw(false);
		for (auto iIndex = 0; iIndex < GetItemCount(); iIndex++)
		{
			EC_B_S(SetItemState(iIndex, LVIS_SELECTED, LVIS_SELECTED | LVIS_FOCUSED));
		}

		MySetRedraw(true);
		if (m_lpHostDlg) m_lpHostDlg->OnUpdateSingleMAPIPropListCtrl(nullptr, nullptr);
	}

	// This is a tough function. I wonder if I'm handling this event correctly
	void CContentsTableListCtrl::OnItemChanged(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
	{
		const auto pNMListView = reinterpret_cast<LPNMLISTVIEW>(pNMHDR);

		*pResult = 0;

		if (!pNMListView || !(pNMListView->uChanged & LVIF_STATE)) return;
		// We get spurious ItemChanged events while scrolling with the keyboard. Ignore them.
		if (GetKeyState(VK_RIGHT) < 0 || GetKeyState(VK_LEFT) < 0) return;

		// Keep all our logic in here
		if (pNMListView->uNewState & LVIS_FOCUSED && m_lpHostDlg)
		{
			LPMAPIPROP lpMAPIProp = nullptr;
			sortlistdata::sortListData* lpData = nullptr;
			std::wstring szTitle;
			if (1 == GetSelectedCount())
			{
				// go get the original row for display in the prop list control
				lpData = GetSortListData(pNMListView->iItem);
				auto row = SRow{};
				if (lpData)
				{
					row = lpData->getRow();
				}

				lpMAPIProp = m_lpHostDlg->OpenItemProp(pNMListView->iItem, modifyType::REQUEST_MODIFY);

				szTitle = strings::loadstring(IDS_DISPLAYNAMENOTFOUND);

				// try to use our rowset first
				if (NODISPLAYNAME != m_ulDisplayNameColumn && row.lpProps && m_ulDisplayNameColumn < row.cValues)
				{
					if (strings::CheckStringProp(&row.lpProps[m_ulDisplayNameColumn], PT_STRING8))
					{
						szTitle = strings::stringTowstring(row.lpProps[m_ulDisplayNameColumn].Value.lpszA);
					}
					else if (strings::CheckStringProp(&row.lpProps[m_ulDisplayNameColumn], PT_UNICODE))
					{
						szTitle = row.lpProps[m_ulDisplayNameColumn].Value.lpszW;
					}
					else
					{
						szTitle = mapi::GetTitle(lpMAPIProp);
					}
				}
				else if (lpMAPIProp)
				{
					szTitle = mapi::GetTitle(lpMAPIProp);
				}
			}

			// Update the main window with our changes
			m_lpHostDlg->OnUpdateSingleMAPIPropListCtrl(lpMAPIProp, lpData);
			m_lpHostDlg->UpdateTitleBarText(szTitle);

			if (lpMAPIProp) lpMAPIProp->Release();
		}
	}

	_Check_return_ bool CContentsTableListCtrl::IsAdviseSet() const noexcept { return m_lpAdviseSink != nullptr; }

	void CContentsTableListCtrl::NotificationOn()
	{
		if (m_lpAdviseSink || !m_lpContentsTable) return;

		output::DebugPrintEx(
			output::dbgLevel::Generic,
			CLASS,
			L"NotificationOn",
			L"registering table notification on %p\n",
			m_lpContentsTable);

		m_lpAdviseSink = new (std::nothrow) mapi::adviseSink(m_hWnd, nullptr);

		if (m_lpAdviseSink)
		{
			const auto hRes = WC_MAPI(m_lpContentsTable->Advise(
				fnevTableModified, static_cast<IMAPIAdviseSink*>(m_lpAdviseSink), &m_ulAdviseConnection));
			if (hRes == MAPI_E_NO_SUPPORT) // Some tables don't support this!
			{
				if (m_lpAdviseSink) m_lpAdviseSink->Release();
				m_lpAdviseSink = nullptr;
				output::DebugPrint(output::dbgLevel::Generic, L"This table doesn't support notifications\n");
			}
			else if (hRes == S_OK)
			{
				const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
				if (lpMDB)
				{
					m_lpAdviseSink->SetAdviseTarget(lpMDB);
					mapi::ForceRop(lpMDB);
				}
			}
		}

		output::DebugPrintEx(
			output::dbgLevel::Generic,
			CLASS,
			L"NotificationOn",
			L"Table notification results (Sink:%p, ulConnection:0x%X) on %p\n",
			m_lpAdviseSink,
			static_cast<int>(m_ulAdviseConnection),
			m_lpContentsTable);
	}

	// This function gets called a lot, make sure it's ok to call it too often...:)
	// If there exists a current advise sink, unadvise it. Otherwise, don't complain.
	void CContentsTableListCtrl::NotificationOff()
	{
		if (!m_lpAdviseSink) return;
		output::DebugPrintEx(
			output::dbgLevel::Generic,
			CLASS,
			L"NotificationOff",
			L"clearing table notification (Sink:%p, ulConnection:0x%X) on %p\n",
			m_lpAdviseSink,
			static_cast<int>(m_ulAdviseConnection),
			m_lpContentsTable);

		if (m_ulAdviseConnection && m_lpContentsTable) m_lpContentsTable->Unadvise(m_ulAdviseConnection);

		m_ulAdviseConnection = NULL;
		if (m_lpAdviseSink) m_lpAdviseSink->Release();
		m_lpAdviseSink = nullptr;
	}

	void CContentsTableListCtrl::RefreshTable()
	{
		if (!m_lpHostDlg) return;
		if (m_bInLoadOp)
		{
			output::DebugPrintEx(
				output::dbgLevel::Generic, CLASS, L"RefreshTable", L"called during table load - ditching call\n");
			return;
		}

		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"RefreshTable", L"\n");

		LoadContentsTableIntoView();

		// Reset the title while we're at it
		m_lpHostDlg->UpdateTitleBarText();
	}

	// Call ExpandRow or CollapseRow as needed
	// Returns
	// S_OK - made a call to ExpandRow/CollapseRow and it succeeded
	// S_FALSE - no call was needed
	// other errors as appropriate
	_Check_return_ HRESULT CContentsTableListCtrl::DoExpandCollapse()
	{
		auto hRes = S_FALSE;

		const auto iItem = GetNextItem(-1, LVNI_SELECTED);

		// nothing selected, no work done
		if (-1 == iItem) return S_FALSE;

		const auto lpData = GetSortListData(iItem);

		// No lpData or wrong type of row - no work done
		if (!lpData) return S_FALSE;
		const auto contents = lpData->cast<sortlistdata::contentsData>();
		if (!contents || contents->getRowType() == TBL_LEAF_ROW || contents->getRowType() == TBL_EMPTY_CATEGORY)
			return S_FALSE;

		auto bDidWork = false;
		LVITEM lvItem = {};
		lvItem.iItem = iItem;
		lvItem.iSubItem = 0;
		lvItem.mask = LVIF_IMAGE;
		switch (contents->getRowType())
		{
		default:
			break;
		case TBL_COLLAPSED_CATEGORY:
		{
			if (contents->getInstanceKey())
			{
				auto instanceKey = contents->getInstanceKey();
				LPSRowSet lpRowSet = nullptr;
				ULONG ulRowsAdded = 0;

				hRes = EC_MAPI(m_lpContentsTable->ExpandRow(
					instanceKey->cb, instanceKey->lpb, 256, NULL, &lpRowSet, &ulRowsAdded));
				if (hRes == S_OK && lpRowSet)
				{
					for (ULONG i = 0; i < lpRowSet->cRows; i++)
					{
						// add the item to the NEXT slot
						AddItemToListBox(iItem + i + 1, &lpRowSet->aRow[i]);
						// Since we handed the props off to the list box, null it out of the row set
						// so we don't free it later with FreeProws
						lpRowSet->aRow[i].lpProps = nullptr;
					}
				}

				FreeProws(lpRowSet);
				contents->setRowType(TBL_EXPANDED_CATEGORY);
				lvItem.iImage = static_cast<int>(sortIcon::nodeExpanded);
				bDidWork = true;
			}
		}

		break;
		case TBL_EXPANDED_CATEGORY:
			if (contents->getInstanceKey())
			{
				auto instanceKey = contents->getInstanceKey();
				ULONG ulRowsRemoved = 0;

				hRes = EC_MAPI(m_lpContentsTable->CollapseRow(instanceKey->cb, instanceKey->lpb, NULL, &ulRowsRemoved));
				if (hRes == S_OK && ulRowsRemoved)
				{
					for (int i = iItem + ulRowsRemoved; i > iItem; i--)
					{
						if (SUCCEEDED(hRes))
						{
							hRes = EC_B(DeleteItem(i));
						}
					}
				}

				contents->setRowType(TBL_COLLAPSED_CATEGORY);
				lvItem.iImage = static_cast<int>(sortIcon::nodeCollapsed);
				bDidWork = true;
			}

			break;
		}

		if (bDidWork)
		{
			hRes = EC_B(SetItem(&lvItem)); // Set new image for the row

			// Save the row type (header/leaf) into lpData
			const auto lpProp = lpData->GetOneProp(PR_ROW_TYPE);
			if (lpProp && PR_ROW_TYPE == lpProp->ulPropTag)
			{
				lpProp->Value.l = contents->getRowType();
			}

			auto sRowData = lpData->getRow();
			SetRowStrings(iItem, &sRowData);
		}

		return hRes;
	}

	void CContentsTableListCtrl::OnOutputTable(const std::wstring& szFileName) const
	{
		if (m_bInLoadOp) return;
		const auto fTable = output::MyOpenFile(szFileName, true);
		if (fTable)
		{
			output::outputTable(output::dbgLevel::NoDebug, fTable, m_lpContentsTable);
			output::CloseFile(fTable);
		}
	}

	void CContentsTableListCtrl::SetSortTable(_In_ LPSSortOrderSet lpSortOrderSet, ULONG ulFlags) const
	{
		if (!m_lpContentsTable) return;
		EC_MAPI_S(m_lpContentsTable->SortTable(lpSortOrderSet, ulFlags));
	}

	// WM_MFCMAPI_THREADADDITEM
	_Check_return_ LRESULT CContentsTableListCtrl::msgOnThreadAddItem(WPARAM wParam, LPARAM lParam)
	{
		const auto iNewRow = static_cast<int>(wParam);
		const auto lpsRow = reinterpret_cast<LPSRow>(lParam);

		if (!lpsRow) return MAPI_E_INVALID_PARAMETER;

		output::DebugPrintEx(
			output::dbgLevel::Generic,
			CLASS,
			L"msgOnThreadAddItem",
			L"Received message to add %p to row %d\n",
			lpsRow,
			iNewRow);
		AddItemToListBox(iNewRow, lpsRow);

		return S_OK;
	}

	// WM_MFCMAPI_ADDITEM
	_Check_return_ LRESULT CContentsTableListCtrl::msgOnAddItem(WPARAM wParam, LPARAM /*lParam*/)
	{
		auto tab = reinterpret_cast<TABLE_NOTIFICATION*>(wParam);

		if (!tab) return MAPI_E_INVALID_PARAMETER;

		// If a row is added, propPrior will contain information about the row preceding the
		// added row. If propPrior.ulPropTag is NULL, then the new item goes first on the list
		auto iNewRow = 0;
		if (PR_NULL == tab->propPrior.ulPropTag)
		{
			iNewRow = 0;
		}
		else
		{
			iNewRow = FindRow(mapi::getBin(tab->propPrior)) + 1;
		}

		// We make this copy here and pass it in to AddItemToListBox, where it is grabbed by sortListData::InitializeContents to be part of the item data
		// The mem will be freed when the item data is cleaned up - do not free here
		SRow NewRow = {};
		NewRow.cValues = tab->row.cValues;
		NewRow.ulAdrEntryPad = tab->row.ulAdrEntryPad;
		const auto hRes =
			EC_MAPI(ScDupPropset(tab->row.cValues, tab->row.lpProps, MAPIAllocateBuffer, &NewRow.lpProps));

		output::DebugPrintEx(
			output::dbgLevel::Generic, CLASS, L"msgOnAddItem", L"Received message to add row to row %d\n", iNewRow);
		AddItemToListBox(iNewRow, &NewRow);

		return hRes;
	}

	// WM_MFCMAPI_DELETEITEM
	_Check_return_ LRESULT CContentsTableListCtrl::msgOnDeleteItem(WPARAM wParam, LPARAM /*lParam*/)
	{
		auto tab = reinterpret_cast<TABLE_NOTIFICATION*>(wParam);

		if (!tab) return MAPI_E_INVALID_PARAMETER;

		const auto iItem = FindRow(mapi::getBin(tab->propIndex));

		output::DebugPrintEx(
			output::dbgLevel::Generic, CLASS, L"msgOnDeleteItem", L"Received message to delete item 0x%d\n", iItem);

		if (iItem == -1) return S_OK;

		const auto hRes = EC_B(DeleteItem(iItem));

		if (S_OK != hRes || !m_lpHostDlg) return hRes;

		const auto iCount = GetItemCount();
		if (iCount == 0)
		{
			m_lpHostDlg->OnUpdateSingleMAPIPropListCtrl(nullptr, nullptr);
		}

		m_lpHostDlg->UpdateStatusBarText(statusPane::data1, IDS_STATUSTEXTNUMITEMS, iCount);
		return hRes;
	}

	// WM_MFCMAPI_MODIFYITEM
	_Check_return_ LRESULT CContentsTableListCtrl::msgOnModifyItem(WPARAM wParam, LPARAM /*lParam*/)
	{
		auto hRes = S_OK;
		auto tab = reinterpret_cast<TABLE_NOTIFICATION*>(wParam);

		if (!tab) return MAPI_E_INVALID_PARAMETER;

		const auto iItem = FindRow(mapi::getBin(tab->propIndex));

		if (-1 != iItem)
		{
			output::DebugPrintEx(
				output::dbgLevel::Generic,
				CLASS,
				L"msgOnModifyItem",
				L"Received message to modify row %d with %p\n",
				iItem,
				&tab->row);

			// We make this copy here and pass it in to RefreshItem, where it is grabbed by sortListData::InitializeContents to be part of the item data
			// The mem will be freed when the item data is cleaned up - do not free here
			SRow NewRow = {};
			NewRow.cValues = tab->row.cValues;
			NewRow.ulAdrEntryPad = tab->row.ulAdrEntryPad;
			hRes = EC_MAPI(ScDupPropset(tab->row.cValues, tab->row.lpProps, MAPIAllocateBuffer, &NewRow.lpProps));

			RefreshItem(iItem, &NewRow, true);
		}

		return hRes;
	}

	// WM_MFCMAPI_REFRESHTABLE
	_Check_return_ LRESULT CContentsTableListCtrl::msgOnRefreshTable(WPARAM /*wParam*/, LPARAM /*lParam*/)
	{
		output::DebugPrintEx(
			output::dbgLevel::Generic, CLASS, L"msgOnRefreshTable", L"Received message refresh table\n");
		RefreshTable();

		return S_OK;
	}

	// This function steps through the list control to find the entry with this instance key
	// return -1 if item not found
	_Check_return_ int CContentsTableListCtrl::FindRow(_In_ const SBinary& instance) const
	{
		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"msgOnGetIndex", L"Getting index for %p\n", &instance);

		auto iItem = 0;
		for (iItem = 0; iItem < GetItemCount(); iItem++)
		{
			const auto lpListData = GetSortListData(iItem);
			if (lpListData)
			{
				const auto contents = lpListData->cast<sortlistdata::contentsData>();
				if (contents)
				{
					const auto lpCurInstance = contents->getInstanceKey();
					if (lpCurInstance)
					{
						if (!memcmp(lpCurInstance->lpb, instance.lpb, instance.cb))
						{
							output::DebugPrintEx(
								output::dbgLevel::Generic, CLASS, L"msgOnGetIndex", L"Matched at 0x%08X\n", iItem);
							return iItem;
						}
					}
				}
			}
		}

		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"msgOnGetIndex", L"No match found: 0x%08X\n", iItem);
		return -1;
	}
} // namespace controls::sortlistctrl