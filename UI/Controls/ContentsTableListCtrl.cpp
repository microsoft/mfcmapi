#include <StdAfx.h>
#include <UI/Controls/SortList/SortListData.h>
#include <UI/Controls/ContentsTableListCtrl.h>
#include <MAPI/MapiObjects.h>
#include <MAPI/MAPIFunctions.h>
#include <UI/UIFunctions.h>
#include <MAPI/AdviseSink.h>
#include <Interpret/InterpretProp.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Dialogs/Editors/TagArrayEditor.h>
#include <Interpret/ExtraPropTags.h>
#include <Interpret/SmartView/SmartView.h>
#include <process.h>
#include <UI/Controls/SortList/ContentsData.h>
#include <UI/Dialogs/BaseDialog.h>
#include <UI/Dialogs/ContentsTable/ContentsTableDlg.h>
#include <MAPI/NamedPropCache.h>

namespace controls
{
	namespace sortlistctrl
	{
		static std::wstring CLASS = L"CContentsTableListCtrl";

#define NODISPLAYNAME 0xffffffff

		CContentsTableListCtrl::CContentsTableListCtrl(
			_In_ CWnd* pCreateParent,
			_In_ CMapiObjects* lpMapiObjects,
			_In_ LPSPropTagArray sptExtraColumnTags,
			_In_ const std::vector<TagNames>& lpExtraDisplayColumns,
			UINT nIDContextMenu,
			bool bIsAB,
			_In_ dialog::CContentsTableDlg *lpHostDlg)
		{
			TRACE_CONSTRUCTOR(CLASS);

			auto hRes = S_OK;

			EC_H(Create(pCreateParent, LVS_NOCOLUMNHEADER, IDC_LIST_CTRL, true));

			m_bAbortLoad = false; // no need to synchronize this - the thread hasn't started yet
			m_bInLoadOp = false;
			m_LoadThreadHandle = nullptr;

			// We borrow our parent's Mapi objects
			m_lpMapiObjects = lpMapiObjects;
			if (m_lpMapiObjects) m_lpMapiObjects->AddRef();

			m_lpHostDlg = lpHostDlg;
			if (m_lpHostDlg) m_lpHostDlg->AddRef();

			m_sptExtraColumnTags = sptExtraColumnTags;
			m_lpExtraDisplayColumns = lpExtraDisplayColumns;
			m_ulDisplayFlags = dfNormal;
			m_ulDisplayNameColumn = NODISPLAYNAME;

			m_ulHeaderColumns = 0;
			m_RestrictionType = mfcmapiNO_RESTRICTION;
			m_lpRes = nullptr;
			m_lpContentsTable = nullptr;
			m_ulContainerType = NULL;
			m_ulAdviseConnection = 0;
			m_lpAdviseSink = nullptr;
			m_nIDContextMenu = nIDContextMenu;
			m_bIsAB = bIsAB;
		}

		CContentsTableListCtrl::~CContentsTableListCtrl()
		{
			TRACE_DESTRUCTOR(CLASS);

			if (m_LoadThreadHandle) CloseHandle(m_LoadThreadHandle);

			NotificationOff();

			if (m_lpRes) MAPIFreeBuffer(const_cast<LPSRestriction>(m_lpRes));
			if (m_lpContentsTable) m_lpContentsTable->Release();
			if (m_lpMapiObjects) m_lpMapiObjects->Release();
			if (m_lpHostDlg) m_lpHostDlg->Release();
		}

		BEGIN_MESSAGE_MAP(CContentsTableListCtrl, CSortListCtrl)
			ON_NOTIFY_REFLECT(LVN_ITEMCHANGED, OnItemChanged)
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
					DrawHelpText(m_hWnd, IDS_HELPTEXTSTARTHERE);
					return true;
				}

				break;
			case WM_LBUTTONDBLCLK:
				WC_H(DoExpandCollapse());
				if (S_FALSE == hRes)
				{
					// Post the message to display the item
					if (m_lpHostDlg)
						m_lpHostDlg->PostMessage(WM_COMMAND, ID_DISPLAYSELECTEDITEM, NULL);
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
				POINT point = { 0 };
				const auto iItem = GetNextItem(
					-1,
					LVNI_SELECTED);
				GetItemPosition(iItem, &point);
				::ClientToScreen(pWnd->m_hWnd, &point);
				pos = point;
			}

			DisplayContextMenu(m_nIDContextMenu, IDR_MENU_TABLE, m_lpHostDlg->m_hWnd, pos.x, pos.y);
		}

		_Check_return_ ULONG CContentsTableListCtrl::GetContainerType() const
		{
			return m_ulContainerType;
		}

		_Check_return_ bool CContentsTableListCtrl::IsContentsTableSet() const
		{
			return m_lpContentsTable != nullptr;
		}

		_Check_return_ HRESULT CContentsTableListCtrl::SetContentsTable(
			_In_opt_ LPMAPITABLE lpContentsTable,
			ULONG ulDisplayFlags,
			ULONG ulContainerType)
		{
			const auto hRes = S_OK;

			// If nothing to do, exit early
			if (lpContentsTable == m_lpContentsTable) return S_OK;
			if (m_bInLoadOp) return MAPI_E_INVALID_PARAMETER;

			CWaitCursor Wait; // Change the mouse to an hourglass while we work.

			m_ulDisplayFlags = ulDisplayFlags;
			m_ulContainerType = ulContainerType;

			DebugPrintEx(DBGGeneric, CLASS, L"SetContentsTable", L"replacing %p with %p\n", m_lpContentsTable, lpContentsTable);
			DebugPrintEx(DBGGeneric, CLASS, L"SetContentsTable", L"New container type: 0x%X\n", m_ulContainerType);
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
			DoSetColumns(
				true,
				0 != RegKeys[regkeyEDIT_COLUMNS_ON_LOAD].ulCurDWORD);

			return hRes;
		}

		void CContentsTableListCtrl::GetStatus()
		{
			auto hRes = S_OK;

			if (!IsContentsTableSet()) return;

			ULONG ulTableStatus = NULL;
			ULONG ulTableType = NULL;

			EC_MAPI(m_lpContentsTable->GetStatus(
				&ulTableStatus,
				&ulTableType));

			if (!FAILED(hRes))
			{
				dialog::editor::CEditor MyData(
					this,
					IDS_GETSTATUS,
					IDS_GETSTATUSPROMPT,
					CEDITOR_BUTTON_OK);
				MyData.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_ULTABLESTATUS, true));
				MyData.SetHex(0, ulTableStatus);
				auto szFlags = interpretprop::InterpretFlags(flagTableStatus, ulTableStatus);
				MyData.InitPane(1, viewpane::TextPane::CreateMultiLinePane(IDS_ULTABLESTATUS, szFlags, true));
				MyData.InitPane(2, viewpane::TextPane::CreateSingleLinePane(IDS_ULTABLETYPE, true));
				MyData.SetHex(2, ulTableType);
				szFlags = interpretprop::InterpretFlags(flagTableType, ulTableType);
				MyData.InitPane(3, viewpane::TextPane::CreateMultiLinePane(IDS_ULTABLETYPE, szFlags, true));

				WC_H(MyData.DisplayDialog());
			}
		}

		// Takes a tag array and builds the UI out of it - does NOT touch the table
		_Check_return_ HRESULT CContentsTableListCtrl::SetUIColumns(_In_ LPSPropTagArray lpTags)
		{
			auto hRes = S_OK;
			if (!lpTags) return MAPI_E_INVALID_PARAMETER;

			// find a PR_DISPLAY_NAME column for later use
			m_ulDisplayNameColumn = NODISPLAYNAME;
			for (ULONG i = 0; i < lpTags->cValues; i++)
			{
				if (PROP_ID(lpTags->aulPropTag[i]) == PROP_ID(PR_DISPLAY_NAME))
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
					if (PROP_ID(lpTags->aulPropTag[i]) == PROP_ID(PR_SUBJECT) ||
						PROP_ID(lpTags->aulPropTag[i]) == PROP_ID(PR_RULE_NAME) ||
						PROP_ID(lpTags->aulPropTag[i]) == PROP_ID(PR_MEMBER_NAME) ||
						PROP_ID(lpTags->aulPropTag[i]) == PROP_ID(PR_ATTACH_LONG_FILENAME) ||
						PROP_ID(lpTags->aulPropTag[i]) == PROP_ID(PR_ATTACH_FILENAME))
					{
						m_ulDisplayNameColumn = i;
						break;
					}
				}
			}

			DebugPrintEx(DBGGeneric, CLASS, L"SetColumns", L"calculating and inserting column headers\n");
			MySetRedraw(false);

			// Delete all of the old column headers
			DeleteAllColumns();

			EC_H(AddColumns(lpTags));

			AutoSizeColumns(true);

			DebugPrintEx(DBGGeneric, CLASS, L"SetColumns", L"Done inserting column headers\n");

			MySetRedraw(true);
			return hRes;
		}

		void CContentsTableListCtrl::DoSetColumns(bool bAddExtras, bool bDisplayEditor)
		{
			auto hRes = S_OK;
			DebugPrintEx(DBGGeneric, CLASS, L"DoSetColumns", L"bDisplayEditor = %d\n", bDisplayEditor);

			if (!IsContentsTableSet())
			{
				// Clear out the selected item view since we have no contents table
				if (m_lpHostDlg)
					m_lpHostDlg->OnUpdateSingleMAPIPropListCtrl(nullptr, nullptr);

				// Make sure we're clear
				DeleteAllColumns();
				ModifyStyle(0, LVS_NOCOLUMNHEADER);
				EC_H(RefreshTable());
				return;
			}

			// these arrays get allocated during the func and need to be freed
			LPSPropTagArray lpConcatTagArray = nullptr;
			LPSPropTagArray lpModifiedTags = nullptr;
			LPSPropTagArray lpOriginalColSet = nullptr;

			auto bModified = false;

			CWaitCursor Wait; // Change the mouse to an hourglass while we work.

			EC_MAPI(m_lpContentsTable->QueryColumns(
				NULL,
				&lpOriginalColSet));
			// this is just a pointer - do not free
			auto lpFinalTagArray = lpOriginalColSet;
			hRes = S_OK;

			if (bAddExtras)
			{
				// build an array with the source set and m_sptExtraColumnTags combined
				EC_H(ConcatSPropTagArrays(
					m_sptExtraColumnTags,
					lpFinalTagArray, // build on the final array we've computed thus far
					&lpConcatTagArray));
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

				WC_H(MyEditor.DisplayDialog());

				if (S_OK == hRes)
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
				EC_MAPI(m_lpContentsTable->SetColumns(
					lpFinalTagArray,
					TBL_BATCH));
				bModified = true;
			}

			if (bModified)
			{
				// Cycle our notification, turning off the old one if necessary
				NotificationOff();
				WC_H(NotificationOn());
				hRes = S_OK;

				EC_H(SetUIColumns(lpFinalTagArray));
				EC_H(RefreshTable());
			}

			MAPIFreeBuffer(lpModifiedTags);
			MAPIFreeBuffer(lpConcatTagArray);
			MAPIFreeBuffer(lpOriginalColSet);
		}

		_Check_return_ HRESULT CContentsTableListCtrl::AddColumn(UINT uidHeaderName, ULONG ulCurHeaderCol, ULONG ulCurTagArrayRow, ULONG ulPropTag)
		{
			auto hRes = S_OK;
			HDITEM hdItem = { 0 };
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
				const auto propTagNames = interpretprop::PropTagToPropName(ulPropTag, m_bIsAB);
				szHeaderString = propTagNames.bestGuess;
				if (szHeaderString.empty())
				{
					const auto namePropNames = NameIDToStrings(
						ulPropTag,
						lpMDB,
						nullptr,
						nullptr,
						m_bIsAB);

					szHeaderString = namePropNames.name;
				}

				if (szHeaderString.empty())
				{
					szHeaderString = strings::format(L"0x%08X", ulPropTag); // STRING_OK
				}
			}

			const auto iRetVal = InsertColumn(ulCurHeaderCol, strings::wstringTotstring(szHeaderString).c_str());

			if (-1 == iRetVal)
			{
				// We failed to insert a column header
				error::ErrDialog(__FILE__, __LINE__, IDS_EDCOLUMNHEADERFAILED);
			}

			if (lpMyHeader)
			{
				hdItem.mask = HDI_LPARAM;
				auto lpHeaderData = new HeaderData; // Will be deleted in CSortListCtrl::DeleteAllColumns
				if (lpHeaderData)
				{
					lpHeaderData->ulTagArrayRow = ulCurTagArrayRow;
					lpHeaderData->ulPropTag = ulPropTag;
					lpHeaderData->bIsAB = m_bIsAB;
					lpHeaderData->szTipString = interpretprop::TagToString(ulPropTag, lpMDB, m_bIsAB, false);

					hdItem.lParam = reinterpret_cast<LPARAM>(lpHeaderData);
					EC_B(lpMyHeader->SetItem(ulCurHeaderCol, &hdItem));
				}
			}

			return hRes;
		}

		// Sets up column headers based on passed in named columns
		// Put all named columns first, followed by a column for each property in the contents table
		_Check_return_ HRESULT CContentsTableListCtrl::AddColumns(_In_ LPSPropTagArray lpCurColTagArray)
		{
			auto hRes = S_OK;

			if (!lpCurColTagArray || !m_lpHostDlg) return MAPI_E_INVALID_PARAMETER;

			m_ulHeaderColumns = lpCurColTagArray->cValues;

			ULONG ulCurHeaderCol = 0;
			if (RegKeys[regkeyDO_COLUMN_NAMES].ulCurDWORD)
			{
				DebugPrintEx(DBGGeneric, CLASS, L"AddColumns", L"Adding named columns\n");
				// If we have named columns, put them up front

				// Walk through the list of named/extra columns and add them to our header list
				for (const auto& extraCol : m_lpExtraDisplayColumns)
				{
					const auto ulExtraColRowNum = extraCol.ulMatchingTableColumn;
					const auto ulExtraColTag = m_sptExtraColumnTags->aulPropTag[ulExtraColRowNum];

					ULONG ulCurTagArrayRow = 0;
					if (FindPropInPropTagArray(lpCurColTagArray, ulExtraColTag, &ulCurTagArrayRow))
					{
						hRes = S_OK;
						EC_H(AddColumn(
							extraCol.uidName,
							ulCurHeaderCol,
							ulCurTagArrayRow,
							lpCurColTagArray->aulPropTag[ulCurTagArrayRow]));
						// Strike out the value in the tag array so we can ignore it later!
						lpCurColTagArray->aulPropTag[ulCurTagArrayRow] = NULL;

						ulCurHeaderCol++;
					}
				}
			}

			DebugPrintEx(DBGGeneric, CLASS, L"AddColumns", L"Adding unnamed columns\n");
			// Now, walk through the current tag table and add each unstruck column to our list
			for (ULONG ulCurTableCol = 0; ulCurTableCol < lpCurColTagArray->cValues; ulCurTableCol++)
			{
				if (lpCurColTagArray->aulPropTag[ulCurTableCol] != NULL)
				{
					hRes = S_OK;
					EC_H(AddColumn(
						NULL,
						ulCurHeaderCol,
						ulCurTableCol,
						lpCurColTagArray->aulPropTag[ulCurTableCol]));
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

			DebugPrintEx(DBGGeneric, CLASS, L"AddColumns", L"Done adding columns\n");
			return hRes;
		}

		void CContentsTableListCtrl::SetRestriction(_In_opt_ const _SRestriction* lpRes)
		{
			MAPIFreeBuffer(const_cast<LPSRestriction>(m_lpRes));
			m_lpRes = lpRes;
		}

		_Check_return_ const _SRestriction* CContentsTableListCtrl::GetRestriction() const
		{
			return m_lpRes;
		}

		_Check_return_ __mfcmapiRestrictionTypeEnum CContentsTableListCtrl::GetRestrictionType() const
		{
			return m_RestrictionType;
		}

		void CContentsTableListCtrl::SetRestrictionType(__mfcmapiRestrictionTypeEnum RestrictionType)
		{
			m_RestrictionType = RestrictionType;
		}

		_Check_return_ HRESULT CContentsTableListCtrl::ApplyRestriction() const
		{
			if (!m_lpContentsTable) return MAPI_E_INVALID_PARAMETER;

			auto hRes = S_OK;
			DebugPrintEx(DBGGeneric, CLASS, L"ApplyRestriction", L"m_RestrictionType = 0x%X\n", m_RestrictionType);
			// Apply our restrictions
			if (mfcmapiNORMAL_RESTRICTION == m_RestrictionType)
			{
				DebugPrintEx(DBGGeneric, CLASS, L"ApplyRestriction", L"applying restriction:\n");

				if (m_lpMapiObjects)
				{
					const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
					DebugPrintRestriction(DBGGeneric, m_lpRes, lpMDB);
				}

				EC_MAPI(m_lpContentsTable->Restrict(
					const_cast<LPSRestriction>(m_lpRes),
					TBL_BATCH));
			}
			else
			{
				WC_H_MSG(m_lpContentsTable->Restrict(
					nullptr,
					TBL_BATCH),
					IDS_TABLENOSUPPORTRES);
			}

			return hRes;
		}

		struct ThreadLoadTableInfo
		{
			HWND hWndHost;
			CContentsTableListCtrl* lpListCtrl;
			LPMAPITABLE lpContentsTable;
			LONG volatile* lpbAbort;
		};

#define bABORTSET (*lpThreadInfo->lpbAbort) // This is safe
#define BREAKONABORT if (bABORTSET) break;
#define CHECKABORT(__fn) if (!bABORTSET) {__fn;}
#define NUMROWSPERLOOP 255

		// Idea here is to do our MAPI work here on this thread, then send messages (SendMessage) back to the control to add the data to the view
		// This way, control functions only happen on the main thread
		// ::SendMessage will be handled on main thread, but block until the call returns.
		// This is the ideal behavior for this worker thread.
		unsigned STDAPICALLTYPE ThreadFuncLoadTable(_In_ void* lpParam)
		{
			auto hRes = S_OK;
			ULONG ulTotal = 0;
			ULONG ulThrottleLevel = 0;
			LPSRowSet pRows = nullptr;
			ULONG iCurListBoxRow = 0;
			const auto lpThreadInfo = static_cast<ThreadLoadTableInfo*>(lpParam);
			if (!lpThreadInfo || !lpThreadInfo->lpbAbort) return 0;

			auto lpListCtrl = lpThreadInfo->lpListCtrl;
			auto lpContentsTable = lpThreadInfo->lpContentsTable;
			if (!lpListCtrl || !lpContentsTable) return 0;

			const auto hWndHost = lpThreadInfo->hWndHost;

			// required on da new thread before we do any MAPI work
			EC_MAPI(MAPIInitialize(nullptr));

			(void) ::SendMessage(hWndHost, WM_MFCMAPI_CLEARSINGLEMAPIPROPLIST, NULL, NULL);
			auto szCount = std::to_wstring(lpListCtrl->GetItemCount());
			dialog::CBaseDialog::UpdateStatus(hWndHost, STATUSDATA1, strings::formatmessage(IDS_STATUSTEXTNUMITEMS, szCount.c_str()));

			// potentially lengthy op - check abort before and after
			CHECKABORT(WC_H(lpListCtrl->ApplyRestriction()));
			hRes = S_OK; // Don't care if the restrict failed - let's try to go on

			if (!bABORTSET) // only check abort once for this group of ops
			{
				// go to the first row
				EC_MAPI(lpContentsTable->SeekRow(
					BOOKMARK_BEGINNING,
					0,
					nullptr));
				hRes = S_OK; // don't let failure here fail the whole load

				EC_MAPI(lpContentsTable->GetRowCount(
					NULL,
					&ulTotal));
				hRes = S_OK; // don't let failure here fail the whole load

				DebugPrintEx(DBGGeneric, CLASS, L"ThreadFuncLoadTable", L"ulTotal = 0x%X\n", ulTotal);

				ulThrottleLevel = RegKeys[regkeyTHROTTLE_LEVEL].ulCurDWORD;

				if (ulTotal)
				{
					dialog::CBaseDialog::UpdateStatus(hWndHost, STATUSDATA2, strings::formatmessage(IDS_LOADINGITEMS, 0, ulTotal));
				}
			}

			const auto lpRes = lpListCtrl->GetRestriction();
			// get rows and add them to the list
			if (!FAILED(hRes)) for (;;)
			{
				BREAKONABORT;
				dialog::CBaseDialog::UpdateStatus(hWndHost, STATUSINFOTEXT, strings::loadstring(IDS_ESCSTOPLOADING));
				hRes = S_OK;
				if (pRows) FreeProws(pRows);
				pRows = nullptr;
				if (mfcmapiFINDROW_RESTRICTION == lpListCtrl->GetRestrictionType() && lpRes)
				{
					DebugPrintEx(DBGGeneric, CLASS, L"DoFindRows", L"running FindRow with restriction:\n");
					DebugPrintRestriction(DBGGeneric, lpRes, nullptr);

					CHECKABORT(WC_MAPI(lpContentsTable->FindRow(
						const_cast<LPSRestriction>(lpRes),
						BOOKMARK_CURRENT,
						NULL)));

					if (MAPI_E_NOT_FOUND != hRes) // MAPI_E_NOT_FOUND signals we didn't find any more rows.
					{
						CHECKABORT(EC_MAPI(lpContentsTable->QueryRows(
							1,
							NULL,
							&pRows)));
					}
					else
					{
						break;
					}
				}
				else
				{
					DebugPrintEx(DBGGeneric, CLASS, L"ThreadFuncLoadTable", L"Calling QueryRows. Asking for 0x%X rows.\n", ulThrottleLevel ? ulThrottleLevel : NUMROWSPERLOOP);
					// Pull back a sizable block of rows to add to the list box
					CHECKABORT(EC_MAPI(lpContentsTable->QueryRows(
						ulThrottleLevel ? ulThrottleLevel : NUMROWSPERLOOP,
						NULL,
						&pRows)));
					if (FAILED(hRes)) break;
				}

				if (FAILED(hRes) || !pRows || !pRows->cRows) break;

				DebugPrintEx(DBGGeneric, CLASS, L"ThreadFuncLoadTable", L"Got this many rows: 0x%X\n", pRows->cRows);

				for (ULONG iCurPropRow = 0; iCurPropRow < pRows->cRows; iCurPropRow++)
				{
					hRes = S_OK;
					BREAKONABORT; // This check is cheap enough not to be a perf concern anymore
					if (ulTotal)
					{
						dialog::CBaseDialog::UpdateStatus(hWndHost, STATUSDATA2, strings::formatmessage(IDS_LOADINGITEMS, iCurListBoxRow + 1, ulTotal));
					}

					DebugPrintEx(DBGGeneric, CLASS, L"ThreadFuncLoadTable", L"Asking to add %p to %u\n", &pRows->aRow[iCurPropRow], iCurListBoxRow);
					(void) ::SendMessage(lpListCtrl->m_hWnd, WM_MFCMAPI_THREADADDITEM, iCurListBoxRow, reinterpret_cast<LPARAM>(&pRows->aRow[iCurPropRow]));
					if (FAILED(hRes)) continue;
					iCurListBoxRow++;
				}

				// Note - we're saving the rows off, so we don't FreeProws this...we just MAPIFreeBuffer the array
				MAPIFreeBuffer(pRows);
				pRows = nullptr;

				if (ulThrottleLevel && iCurListBoxRow >= ulThrottleLevel) break; // Only render ulThrottleLevel rows if throttle is on
			}

			if (bABORTSET)
			{
				dialog::CBaseDialog::UpdateStatus(hWndHost, STATUSINFOTEXT, strings::loadstring(IDS_TABLELOADCANCELLED));
			}
			else
			{
				dialog::CBaseDialog::UpdateStatus(hWndHost, STATUSINFOTEXT, strings::loadstring(IDS_TABLELOADED));
			}

			dialog::CBaseDialog::UpdateStatus(hWndHost, STATUSDATA2, strings::emptystring);
			DebugPrintEx(DBGGeneric, CLASS, L"ThreadFuncLoadTable", L"added %u items\n", iCurListBoxRow);
			DebugPrintEx(DBGGeneric, CLASS, L"ThreadFuncLoadTable", L"Releasing pointers.\n");

			lpListCtrl->ClearLoading();

			// Bunch of cleanup
			if (pRows) FreeProws(pRows);
			if (lpContentsTable) lpContentsTable->Release();
			if (lpListCtrl) lpListCtrl->Release();
			DebugPrintEx(DBGGeneric, CLASS, L"ThreadFuncLoadTable", L"Pointers released.\n");

			MAPIUninitialize();

			delete lpThreadInfo;

			return 0;
		}

		_Check_return_ bool CContentsTableListCtrl::IsLoading() const
		{
			return m_bInLoadOp;
		}

		void CContentsTableListCtrl::ClearLoading()
		{
			m_bInLoadOp = false;
		}

		_Check_return_ HRESULT CContentsTableListCtrl::LoadContentsTableIntoView()
		{
			auto hRes = S_OK;
			CWaitCursor Wait; // Change the mouse to an hourglass while we work.

			DebugPrintEx(DBGGeneric, CLASS, L"LoadContentsTableIntoView", L"\n");

			if (m_bInLoadOp) return MAPI_E_INVALID_PARAMETER;
			if (!m_lpHostDlg) return MAPI_E_INVALID_PARAMETER;

			EC_B(DeleteAllItems());

			// whack the old thread handle if we still have it
			if (m_LoadThreadHandle) CloseHandle(m_LoadThreadHandle);
			m_LoadThreadHandle = nullptr;

			if (!m_lpContentsTable) return S_OK;
			m_bInLoadOp = true;
			// Do not call return after this point!

			auto lpThreadInfo = new ThreadLoadTableInfo;

			if (lpThreadInfo)
			{
				lpThreadInfo->hWndHost = m_lpHostDlg->m_hWnd;
				lpThreadInfo->lpbAbort = &m_bAbortLoad;
				m_bAbortLoad = false; // no need to synchronize this - the thread hasn't started yet

				lpThreadInfo->lpListCtrl = this;
				this->AddRef();

				lpThreadInfo->lpContentsTable = m_lpContentsTable;
				m_lpContentsTable->AddRef();

				DebugPrintEx(DBGGeneric, CLASS, L"LoadContentsTableIntoView", L"Creating load thread.\n");

				HANDLE hThread = nullptr;
				EC_D(hThread, reinterpret_cast<HANDLE>(_beginthreadex(nullptr, 0, ThreadFuncLoadTable, lpThreadInfo, 0, nullptr)));

				if (!hThread)
				{
					DebugPrintEx(DBGGeneric, CLASS, L"LoadContentsTableIntoView", L"Load thread creation failed.\n");
					if (lpThreadInfo->lpContentsTable) lpThreadInfo->lpContentsTable->Release();
					if (lpThreadInfo->lpListCtrl) lpThreadInfo->lpListCtrl->Release();
					delete lpThreadInfo;
				}
				else
				{
					DebugPrintEx(DBGGeneric, CLASS, L"LoadContentsTableIntoView", L"Load thread created.\n");
					m_LoadThreadHandle = hThread;
				}
			}

			return hRes;
		}

		void CContentsTableListCtrl::OnCancelTableLoad()
		{
			DebugPrintEx(DBGGeneric, CLASS, L"OnCancelTableLoad", L"Setting abort flag and waiting for thread to discover it\n");
			// Wait here until the thread we spun off has shut down
			CWaitCursor Wait; // Change the mouse to an hourglass while we work.
			DWORD dwRet = 0;
			auto bVKF5Hit = false;

			// See if the thread is still active
			while (m_LoadThreadHandle) // this won't change, but if it's NULL, we just skip the loop
			{
				MSG msg;

				InterlockedExchange(&m_bAbortLoad, true);

				// Wait for the thread to shutdown/signal, or messages posted to our queue
				dwRet = MsgWaitForMultipleObjects(
					1,
					&m_LoadThreadHandle,
					false,
					INFINITE,
					QS_ALLINPUT);
				if (dwRet == WAIT_OBJECT_0 + 0) break;

				// Read all of the messages in this next loop, removing each message as we read it.
				// If we don't do this, the thread never stops
				while (PeekMessage(&msg, m_hWnd, 0, 0, PM_REMOVE))
				{
					if (msg.message == WM_KEYDOWN && msg.wParam == VK_F5)
					{
						DebugPrintEx(DBGGeneric, CLASS, L"OnCancelTableLoad", L"Ditching refresh (F5)\n");
						bVKF5Hit = true;
					}
					else
					{
						DispatchMessage(&msg);
					}
				}
			}

			DebugPrintEx(DBGGeneric, CLASS, L"OnCancelTableLoad", L"Load thread has shut down.\n");

			if (m_LoadThreadHandle) CloseHandle(m_LoadThreadHandle);
			m_LoadThreadHandle = nullptr;
			m_bAbortLoad = false;

			if (bVKF5Hit) // If we ditched a refresh message, repost it now
			{
				DebugPrintEx(DBGGeneric, CLASS, L"OnCancelTableLoad", L"Posting skipped refresh message\n");
				PostMessage(WM_KEYDOWN, VK_F5, 0);
			}
		}

		void CContentsTableListCtrl::SetRowStrings(int iRow, _In_ LPSRow lpsRowData)
		{
			if (!lpsRowData) return;

			auto hRes = S_OK;
			const auto lpMyHeader = GetHeaderCtrl();

			if (!lpMyHeader) return;

			for (ULONG iColumn = 0; iColumn < m_ulHeaderColumns; iColumn++)
			{
				HDITEM hdItem = { 0 };
				hdItem.mask = HDI_LPARAM;
				EC_B(lpMyHeader->GetItem(iColumn, &hdItem));

				if (hdItem.lParam)
				{
					const auto ulCol = reinterpret_cast<LPHEADERDATA>(hdItem.lParam)->ulTagArrayRow;
					hRes = S_OK;

					if (ulCol < lpsRowData->cValues)
					{
						std::wstring PropString;
						const auto pProp = &lpsRowData->lpProps[ulCol];

						// If we've got a MAPI_E_NOT_FOUND error, just don't display it.
						if (RegKeys[regkeySUPPRESS_NOT_FOUND].ulCurDWORD && pProp && PT_ERROR == PROP_TYPE(pProp->ulPropTag) && MAPI_E_NOT_FOUND == pProp->Value.err)
						{
							if (0 == iColumn)
							{
								SetItemText(iRow, iColumn, L"");
							}

							continue;
						}

						interpretprop::InterpretProp(pProp, &PropString, nullptr);

						auto szFlags = smartview::InterpretNumberAsString(pProp->Value, pProp->ulPropTag, NULL, nullptr, nullptr, false);
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

#define NUMOBJTYPES 12
		static LONG _ObjTypeIcons[NUMOBJTYPES][2] =
		{
		 { MAPI_STORE, slIconMAPI_STORE },
		 { MAPI_ADDRBOOK, slIconMAPI_ADDRBOOK },
		 { MAPI_FOLDER, slIconMAPI_FOLDER },
		 { MAPI_ABCONT, slIconMAPI_ABCONT },
		 { MAPI_MESSAGE, slIconMAPI_MESSAGE },
		 { MAPI_MAILUSER, slIconMAPI_MAILUSER },
		 { MAPI_ATTACH, slIconMAPI_ATTACH },
		 { MAPI_DISTLIST, slIconMAPI_DISTLIST },
		 { MAPI_PROFSECT, slIconMAPI_PROFSECT },
		 { MAPI_STATUS, slIconMAPI_STATUS },
		 { MAPI_SESSION, slIconMAPI_SESSION },
		 { MAPI_FORMINFO, slIconMAPI_FORMINFO },
		};

		void GetDepthAndImage(_In_ LPSRow lpsRowData, _In_ ULONG* lpulDepth, _In_ ULONG* lpulImage)
		{
			if (lpulDepth) *lpulDepth = 0;
			if (lpulImage) *lpulImage = slIconDefault;
			if (!lpsRowData) return;

			ULONG ulDepth = NULL;
			ULONG ulImage = slIconDefault;
			const auto lpDepth = PpropFindProp(lpsRowData->lpProps, lpsRowData->cValues, PR_DEPTH);
			if (lpDepth && PR_DEPTH == lpDepth->ulPropTag) ulDepth = lpDepth->Value.l;
			if (ulDepth > 5) ulDepth = 5; // Just in case

			const auto lpRowType = PpropFindProp(lpsRowData->lpProps, lpsRowData->cValues, PR_ROW_TYPE);
			if (lpRowType && PR_ROW_TYPE == lpRowType->ulPropTag)
			{
				switch (lpRowType->Value.l)
				{
				case TBL_LEAF_ROW:
					break;
				case TBL_EMPTY_CATEGORY:
				case TBL_COLLAPSED_CATEGORY:
					ulImage = slIconNodeCollapsed;
					break;
				case TBL_EXPANDED_CATEGORY:
					ulImage = slIconNodeExpanded;
					break;
				}
			}

			if (slIconDefault == ulImage)
			{
				const auto lpObjType = PpropFindProp(lpsRowData->lpProps, lpsRowData->cValues, PR_OBJECT_TYPE);
				if (lpObjType && PR_OBJECT_TYPE == lpObjType->ulPropTag)
				{
					for (auto& _ObjTypeIcon : _ObjTypeIcons)
					{
						if (_ObjTypeIcon[0] == lpObjType->Value.l)
						{
							ulImage = _ObjTypeIcon[1];
							break;
						}
					}
				}
			}

			// We still don't have a good icon - make some heuristic guesses
			if (slIconDefault == ulImage)
			{
				auto lpProp = PpropFindProp(lpsRowData->lpProps, lpsRowData->cValues, PR_SERVICE_UID);
				if (!lpProp)
				{
					lpProp = PpropFindProp(lpsRowData->lpProps, lpsRowData->cValues, PR_PROVIDER_UID);
				}

				if (lpProp)
				{
					ulImage = slIconMAPI_PROFSECT;
				}
			}

			if (lpulDepth) *lpulDepth = ulDepth;
			if (lpulImage) *lpulImage = ulImage;
		}

		_Check_return_ HRESULT CContentsTableListCtrl::RefreshItem(int iRow, _In_ LPSRow lpsRowData, bool bItemExists)
		{
			const auto hRes = S_OK;
			sortlistdata::SortListData* lpData = nullptr;

			DebugPrintEx(DBGGeneric, CLASS, L"RefreshItem", L"item %d\n", iRow);

			if (bItemExists)
			{
				lpData = GetSortListData(iRow);
			}
			else
			{
				ULONG ulDepth = NULL;
				ULONG ulImage = slIconDefault;
				GetDepthAndImage(lpsRowData, &ulDepth, &ulImage);

				lpData = InsertRow(iRow, L"TempRefreshItem", ulDepth, ulImage); // STRING_OK
			}

			if (lpData)
			{
				lpData->InitializeContents(lpsRowData);

				SetRowStrings(iRow, lpsRowData);
				// Do this last so that our row can't get sorted before we're done!
				lpData->bItemFullyLoaded = true;
			}

			return hRes;
		}

		// Crack open the given SPropValue and render it to the given row in the list.
		_Check_return_ HRESULT CContentsTableListCtrl::AddItemToListBox(int iRow, _In_ LPSRow lpsRowToAdd)
		{
			auto hRes = S_OK;

			DebugPrintEx(DBGGeneric, CLASS, L"AddItemToListBox", L"item %d\n", iRow);

			EC_H(RefreshItem(iRow, lpsRowToAdd, false));

			if (m_lpHostDlg)
				m_lpHostDlg->UpdateStatusBarText(STATUSDATA1, IDS_STATUSTEXTNUMITEMS, GetItemCount());

			return hRes;
		}

		void CContentsTableListCtrl::OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags)
		{
			DebugPrintEx(DBGMenu, CLASS, L"OnKeyDown", L"0x%X\n", nChar);

			if (!m_lpHostDlg) return;
			const auto bCtrlPressed = GetKeyState(VK_CONTROL) < 0;
			const auto bShiftPressed = GetKeyState(VK_SHIFT) < 0;
			const auto bMenuPressed = GetKeyState(VK_MENU) < 0;

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

		_Check_return_ HRESULT CContentsTableListCtrl::GetSelectedItemEIDs(_Deref_out_opt_ LPENTRYLIST* lppEntryIDs) const
		{
			*lppEntryIDs = nullptr;
			auto hRes = S_OK;
			const UINT iNumItems = GetSelectedCount();

			if (!iNumItems) return S_OK;
			if (iNumItems > ULONG_MAX / sizeof(SBinary)) return MAPI_E_INVALID_PARAMETER;

			LPENTRYLIST lpTempList = nullptr;

			EC_H(MAPIAllocateBuffer(sizeof(ENTRYLIST), reinterpret_cast<LPVOID*>(&lpTempList)));

			if (lpTempList)
			{
				lpTempList->cValues = iNumItems;
				lpTempList->lpbin = nullptr;

				EC_H(MAPIAllocateMore(
					static_cast<ULONG>(sizeof(SBinary))* iNumItems,
					lpTempList,
					reinterpret_cast<LPVOID*>(&lpTempList->lpbin)));
				if (lpTempList->lpbin)
				{
					auto iSelectedItem = -1;

					for (UINT iArrayPos = 0; iArrayPos < iNumItems; iArrayPos++)
					{
						lpTempList->lpbin[iArrayPos].cb = 0;
						lpTempList->lpbin[iArrayPos].lpb = nullptr;
						iSelectedItem = GetNextItem(
							iSelectedItem,
							LVNI_SELECTED);
						if (-1 != iSelectedItem)
						{
							const auto lpData = GetSortListData(iSelectedItem);
							if (lpData && lpData->Contents() && lpData->Contents()->m_lpEntryID)
							{
								lpTempList->lpbin[iArrayPos].cb = lpData->Contents()->m_lpEntryID->cb;
								EC_H(MAPIAllocateMore(
									lpData->Contents()->m_lpEntryID->cb,
									lpTempList,
									reinterpret_cast<LPVOID *>(&lpTempList->lpbin[iArrayPos].lpb)));
								if (lpTempList->lpbin[iArrayPos].lpb)
								{
									CopyMemory(
										lpTempList->lpbin[iArrayPos].lpb,
										lpData->Contents()->m_lpEntryID->lpb,
										lpData->Contents()->m_lpEntryID->cb);
								}
							}
						}
					}
				}
			}

			*lppEntryIDs = lpTempList;
			return hRes;
		}

		_Check_return_ std::vector<int> CContentsTableListCtrl::GetSelectedItemNums() const
		{
			auto iItem = -1;
			std::vector<int> iItems;
			do
			{
				iItem = GetNextItem(
					iItem,
					LVNI_SELECTED);
				if (iItem != -1)
				{
					iItems.push_back(iItem);
					DebugPrintEx(DBGGeneric, CLASS, L"GetSelectedItemNums", L"iItem: 0x%X\n", iItem);
				}
			} while (iItem != -1);

			return iItems;
		}

		_Check_return_ std::vector<sortlistdata::SortListData*> CContentsTableListCtrl::GetSelectedItemData() const
		{
			auto iItem = -1;
			std::vector<sortlistdata::SortListData*> items;
			do
			{
				iItem = GetNextItem(
					iItem,
					LVNI_SELECTED);
				if (iItem != -1)
				{
					items.push_back(GetSortListData(iItem));
				}
			} while (iItem != -1);

			return items;
		}

		_Check_return_ sortlistdata::SortListData* CContentsTableListCtrl::GetFirstSelectedItemData() const
		{
			const auto iItem = GetNextItem(
				-1,
				LVNI_SELECTED);
			if (-1 == iItem) return nullptr;

			return GetSortListData(iItem);
		}

		// Pass iCurItem as -1 to get the primary selected item.
		// Call again with the previous iCurItem to get the next one.
		// Stop calling when iCurItem = -1 and/or lppProp is NULL
		// If iCurItem is NULL, just returns the focused item
		_Check_return_ int CContentsTableListCtrl::GetNextSelectedItemNum(
			_Inout_opt_ int *iCurItem) const
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

			DebugPrintEx(DBGGeneric, CLASS, L"GetNextSelectedItemNum", L"iItem before = 0x%X\n", iItem);

			iItem = GetNextItem(
				iItem,
				LVNI_SELECTED);

			DebugPrintEx(DBGGeneric, CLASS, L"GetNextSelectedItemNum", L"iItem after = 0x%X\n", iItem);

			if (iCurItem) *iCurItem = iItem;

			return iItem;
		}

		_Check_return_ sortlistdata::SortListData* CContentsTableListCtrl::GetSortListData(int iItem) const
		{
			return reinterpret_cast<sortlistdata::SortListData*>(GetItemData(iItem));
		}

		// Pass iCurItem as -1 to get the primary selected item.
		// Call again with the previous iCurItem to get the next one.
		// Stop calling when iCurItem = -1 and/or lppProp is NULL
		// If iCurItem is NULL, just returns the focused item
		_Check_return_ HRESULT CContentsTableListCtrl::OpenNextSelectedItemProp(
			_Inout_opt_ int *iCurItem,
			__mfcmapiModifyEnum bModify,
			_Deref_out_opt_ LPMAPIPROP* lppProp) const
		{
			auto hRes = S_OK;

			*lppProp = nullptr;

			const auto iItem = GetNextSelectedItemNum(iCurItem);
			if (-1 != iItem)
				WC_H(m_lpHostDlg->OpenItemProp(iItem, bModify, lppProp));

			return hRes;
		}

		_Check_return_ HRESULT CContentsTableListCtrl::DefaultOpenItemProp(
			int iItem,
			__mfcmapiModifyEnum bModify,
			_Deref_out_opt_ LPMAPIPROP* lppProp) const
		{
			auto hRes = S_OK;

			*lppProp = nullptr;

			if (!m_lpMapiObjects || -1 == iItem) return S_OK;

			DebugPrintEx(DBGGeneric, CLASS, L"DefaultOpenItemProp", L"iItem = %d, bModify = %d, m_ulContainerType = 0x%X\n", iItem, bModify, m_ulContainerType);

			const auto lpListData = GetSortListData(iItem);
			if (!lpListData || !lpListData->Contents()) return S_OK;

			const auto lpEID = lpListData->Contents()->m_lpEntryID;
			if (!lpEID || lpEID->cb == 0) return S_OK;

			DebugPrint(DBGGeneric, L"Item being opened:\n");
			DebugPrintBinary(DBGGeneric, *lpEID);

			// Find the highlighted item EID
			switch (m_ulContainerType)
			{
			case MAPI_ABCONT:
			{
				const auto lpAB = m_lpMapiObjects->GetAddrBook(false); // do not release
				WC_H(CallOpenEntry(
					nullptr,
					lpAB, // use AB
					nullptr,
					nullptr,
					lpEID,
					nullptr,
					bModify == mfcmapiREQUEST_MODIFY ? MAPI_MODIFY : MAPI_BEST_ACCESS,
					nullptr,
					reinterpret_cast<LPUNKNOWN*>(lppProp)));
			}

			break;
			case MAPI_FOLDER:
			{
				const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
				LPCIID lpInterface = nullptr;

				if (RegKeys[regkeyUSE_MESSAGERAW].ulCurDWORD)
				{
					lpInterface = &IID_IMessageRaw;
				}

				WC_H(CallOpenEntry(
					lpMDB, // use MDB
					nullptr,
					nullptr,
					nullptr,
					lpEID,
					lpInterface,
					bModify == mfcmapiREQUEST_MODIFY ? MAPI_MODIFY : MAPI_BEST_ACCESS,
					nullptr,
					reinterpret_cast<LPUNKNOWN*>(lppProp)));
				if (MAPI_E_INTERFACE_NOT_SUPPORTED == hRes && RegKeys[regkeyUSE_MESSAGERAW].ulCurDWORD)
				{
					error::ErrDialog(__FILE__, __LINE__, IDS_EDMESSAGERAWNOTSUPPORTED);
				}
			}

			break;
			default:
			{
				const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
				WC_H(CallOpenEntry(
					nullptr,
					nullptr,
					nullptr,
					lpMAPISession, // use session
					lpEID,
					nullptr,
					bModify == mfcmapiREQUEST_MODIFY ? MAPI_MODIFY : MAPI_BEST_ACCESS,
					nullptr,
					reinterpret_cast<LPUNKNOWN*>(lppProp)));
			}

			break;
			}

			if (!*lppProp && FAILED(hRes) && mfcmapiREQUEST_MODIFY == bModify && MAPI_E_NOT_FOUND != hRes)
			{
				DebugPrint(DBGGeneric, L"\tOpenEntry failed: 0x%X. Will try again without MAPI_MODIFY\n", hRes);
				// We got access denied when we passed MAPI_MODIFY
				// Let's try again without it.
				hRes = S_OK;
				EC_H(DefaultOpenItemProp(
					iItem,
					mfcmapiDO_NOT_REQUEST_MODIFY,
					lppProp));
			}

			if (MAPI_E_NOT_FOUND == hRes)
			{
				DebugPrint(DBGGeneric, L"\tDefaultOpenItemProp encountered an entry ID for an item that doesn't exist\n\tThis happens often when we're deleting items.\n");
				hRes = S_OK;
			}

			DebugPrintEx(DBGGeneric, CLASS, L"DefaultOpenItemProp", L"returning *lppProp = %p and hRes = 0x%X\n", *lppProp, hRes);
			return hRes;
		}

		void CContentsTableListCtrl::SelectAll()
		{
			auto hRes = S_OK;
			DebugPrintEx(DBGGeneric, CLASS, L"SelectAll", L"\n");
			CWaitCursor Wait; // Change the mouse to an hourglass while we work.
			MySetRedraw(false);
			for (auto iIndex = 0; iIndex < GetItemCount(); iIndex++)
			{
				EC_B(SetItemState(iIndex, LVIS_SELECTED, LVIS_SELECTED | LVIS_FOCUSED));
				hRes = S_OK;
			}

			MySetRedraw(true);
			if (m_lpHostDlg)
				m_lpHostDlg->OnUpdateSingleMAPIPropListCtrl(nullptr, nullptr);
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
				sortlistdata::SortListData* lpData = nullptr;
				std::wstring szTitle;
				if (1 == GetSelectedCount())
				{
					auto hRes = S_OK;

					// go get the original row for display in the prop list control
					lpData = GetSortListData(pNMListView->iItem);
					ULONG cValues = 0;
					LPSPropValue lpProps = nullptr;
					if (lpData)
					{
						cValues = lpData->cSourceProps;
						lpProps = lpData->lpSourceProps;
					}

					WC_H(m_lpHostDlg->OpenItemProp(pNMListView->iItem, mfcmapiREQUEST_MODIFY, &lpMAPIProp));

					szTitle = strings::loadstring(IDS_DISPLAYNAMENOTFOUND);

					// try to use our rowset first
					if (NODISPLAYNAME != m_ulDisplayNameColumn
						&& lpProps
						&& m_ulDisplayNameColumn < cValues)
					{
						if (CheckStringProp(&lpProps[m_ulDisplayNameColumn], PT_STRING8))
						{
							szTitle = strings::stringTowstring(lpProps[m_ulDisplayNameColumn].Value.lpszA);
						}
						else if (CheckStringProp(&lpProps[m_ulDisplayNameColumn], PT_UNICODE))
						{
							szTitle = lpProps[m_ulDisplayNameColumn].Value.lpszW;
						}
						else
						{
							szTitle = GetTitle(lpMAPIProp);
						}
					}
					else if (lpMAPIProp)
					{
						szTitle = GetTitle(lpMAPIProp);
					}
				}

				// Update the main window with our changes
				m_lpHostDlg->OnUpdateSingleMAPIPropListCtrl(lpMAPIProp, lpData);
				m_lpHostDlg->UpdateTitleBarText(szTitle);

				if (lpMAPIProp) lpMAPIProp->Release();
			}
		}

		_Check_return_ bool CContentsTableListCtrl::IsAdviseSet() const
		{
			return m_lpAdviseSink != nullptr;
		}

		_Check_return_ HRESULT CContentsTableListCtrl::NotificationOn()
		{
			auto hRes = S_OK;

			if (m_lpAdviseSink || !m_lpContentsTable) return S_OK;

			DebugPrintEx(DBGGeneric, CLASS, L"NotificationOn", L"registering table notification on %p\n", m_lpContentsTable);

			m_lpAdviseSink = new CAdviseSink(m_hWnd, nullptr);

			if (m_lpAdviseSink)
			{
				WC_MAPI(m_lpContentsTable->Advise(
					fnevTableModified,
					static_cast<IMAPIAdviseSink *>(m_lpAdviseSink),
					&m_ulAdviseConnection));
				if (MAPI_E_NO_SUPPORT == hRes) // Some tables don't support this!
				{
					if (m_lpAdviseSink) m_lpAdviseSink->Release();
					m_lpAdviseSink = nullptr;
					DebugPrint(DBGGeneric, L"This table doesn't support notifications\n");
					hRes = S_OK; // mask the error
				}
				else if (S_OK == hRes)
				{
					const auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
					if (lpMDB)
					{
						m_lpAdviseSink->SetAdviseTarget(lpMDB);
						ForceRop(lpMDB);
					}
				}
			}

			DebugPrintEx(DBGGeneric, CLASS, L"NotificationOn", L"Table notification results (Sink:%p, ulConnection:0x%X) on %p\n", m_lpAdviseSink, static_cast<int>(m_ulAdviseConnection), m_lpContentsTable);
			return hRes;
		}

		// This function gets called a lot, make sure it's ok to call it too often...:)
		// If there exists a current advise sink, unadvise it. Otherwise, don't complain.
		void CContentsTableListCtrl::NotificationOff()
		{
			if (!m_lpAdviseSink) return;
			DebugPrintEx(DBGGeneric, CLASS, L"NotificationOff", L"clearing table notification (Sink:%p, ulConnection:0x%X) on %p\n", m_lpAdviseSink, static_cast<int>(m_ulAdviseConnection), m_lpContentsTable);

			if (m_ulAdviseConnection && m_lpContentsTable)
				m_lpContentsTable->Unadvise(m_ulAdviseConnection);

			m_ulAdviseConnection = NULL;
			if (m_lpAdviseSink) m_lpAdviseSink->Release();
			m_lpAdviseSink = nullptr;
		}

		_Check_return_ HRESULT CContentsTableListCtrl::RefreshTable()
		{
			auto hRes = S_OK;
			if (!m_lpHostDlg) return MAPI_E_INVALID_PARAMETER;
			if (m_bInLoadOp)
			{
				DebugPrintEx(DBGGeneric, CLASS, L"RefreshTable", L"called during table load - ditching call\n");
				return S_OK;
			}

			DebugPrintEx(DBGGeneric, CLASS, L"RefreshTable", L"\n");

			EC_H(LoadContentsTableIntoView());

			// Reset the title while we're at it
			m_lpHostDlg->UpdateTitleBarText();

			return hRes;
		}

		// Call ExpandRow or CollapseRow as needed
		// Returns
		// S_OK - made a call to ExpandRow/CollapseRow and it succeeded
		// S_FALSE - no call was needed
		// other errors as appropriate
		_Check_return_ HRESULT CContentsTableListCtrl::DoExpandCollapse()
		{
			auto hRes = S_FALSE;

			const auto iItem = GetNextItem(
				-1,
				LVNI_SELECTED);

			// nothing selected, no work done
			if (-1 == iItem) return S_FALSE;

			const auto lpData = GetSortListData(iItem);

			// No lpData or wrong type of row - no work done
			if (!lpData ||
				!lpData->Contents() ||
				lpData->Contents()->m_ulRowType == TBL_LEAF_ROW ||
				lpData->Contents()->m_ulRowType == TBL_EMPTY_CATEGORY)
				return S_FALSE;

			auto bDidWork = false;
			LVITEM lvItem = { 0 };
			lvItem.iItem = iItem;
			lvItem.iSubItem = 0;
			lvItem.mask = LVIF_IMAGE;
			switch (lpData->Contents()->m_ulRowType)
			{
			default:
				break;
			case TBL_COLLAPSED_CATEGORY:
			{
				if (lpData->Contents()->m_lpInstanceKey)
				{
					LPSRowSet lpRowSet = nullptr;
					ULONG ulRowsAdded = 0;

					EC_MAPI(m_lpContentsTable->ExpandRow(
						lpData->Contents()->m_lpInstanceKey->cb,
						lpData->Contents()->m_lpInstanceKey->lpb,
						256,
						NULL,
						&lpRowSet,
						&ulRowsAdded));
					if (S_OK == hRes && lpRowSet)
					{
						for (ULONG i = 0; i < lpRowSet->cRows; i++)
						{
							// add the item to the NEXT slot
							EC_H(AddItemToListBox(iItem + i + 1, &lpRowSet->aRow[i]));
							// Since we handed the props off to the list box, null it out of the row set
							// so we don't free it later with FreeProws
							lpRowSet->aRow[i].lpProps = nullptr;
						}
					}

					FreeProws(lpRowSet);
					lpData->Contents()->m_ulRowType = TBL_EXPANDED_CATEGORY;
					lvItem.iImage = slIconNodeExpanded;
					bDidWork = true;
				}
			}

			break;
			case TBL_EXPANDED_CATEGORY:
				if (lpData->Contents()->m_lpInstanceKey)
				{
					ULONG ulRowsRemoved = 0;

					EC_MAPI(m_lpContentsTable->CollapseRow(
						lpData->Contents()->m_lpInstanceKey->cb,
						lpData->Contents()->m_lpInstanceKey->lpb,
						NULL,
						&ulRowsRemoved));
					if (S_OK == hRes && ulRowsRemoved)
					{
						for (int i = iItem + ulRowsRemoved; i > iItem; i--)
						{
							EC_B(DeleteItem(i));
						}
					}

					lpData->Contents()->m_ulRowType = TBL_COLLAPSED_CATEGORY;
					lvItem.iImage = slIconNodeCollapsed;
					bDidWork = true;
				}

				break;
			}

			if (bDidWork)
			{
				EC_B(SetItem(&lvItem)); // Set new image for the row

				// Save the row type (header/leaf) into lpData
				const auto lpProp = PpropFindProp(
					lpData->lpSourceProps,
					lpData->cSourceProps,
					PR_ROW_TYPE);
				if (lpProp && PR_ROW_TYPE == lpProp->ulPropTag)
				{
					lpProp->Value.l = lpData->Contents()->m_ulRowType;
				}

				SRow sRowData = { 0 };
				sRowData.cValues = lpData->cSourceProps;
				sRowData.lpProps = lpData->lpSourceProps;
				SetRowStrings(iItem, &sRowData);
			}

			return hRes;
		}

		void CContentsTableListCtrl::OnOutputTable(const std::wstring& szFileName) const
		{
			if (m_bInLoadOp) return;
			const auto fTable = MyOpenFile(szFileName, true);
			if (fTable)
			{
				OutputTableToFile(fTable, m_lpContentsTable);
				CloseFile(fTable);
			}
		}

		_Check_return_ HRESULT CContentsTableListCtrl::SetSortTable(_In_ LPSSortOrderSet lpSortOrderSet, ULONG ulFlags) const
		{
			auto hRes = S_OK;
			if (!m_lpContentsTable) return MAPI_E_INVALID_PARAMETER;

			EC_MAPI(m_lpContentsTable->SortTable(
				lpSortOrderSet,
				ulFlags));

			return hRes;
		}

		// WM_MFCMAPI_THREADADDITEM
		_Check_return_ LRESULT CContentsTableListCtrl::msgOnThreadAddItem(WPARAM wParam, LPARAM lParam)
		{
			auto hRes = S_OK;
			const auto iNewRow = static_cast<int>(wParam);
			const auto lpsRow = reinterpret_cast<LPSRow>(lParam);

			if (!lpsRow) return MAPI_E_INVALID_PARAMETER;

			DebugPrintEx(DBGGeneric, CLASS, L"msgOnThreadAddItem", L"Received message to add %p to row %d\n", lpsRow, iNewRow);
			EC_H(AddItemToListBox(iNewRow, lpsRow));

			return hRes;
		}

		// WM_MFCMAPI_ADDITEM
		_Check_return_ LRESULT CContentsTableListCtrl::msgOnAddItem(WPARAM wParam, LPARAM /*lParam*/)
		{
			auto hRes = S_OK;
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
				iNewRow = FindRow(&tab->propPrior.Value.bin) + 1;
			}

			// We make this copy here and pass it in to AddItemToListBox, where it is grabbed by SortListData::InitializeContents to be part of the item data
			// The mem will be freed when the item data is cleaned up - do not free here
			SRow NewRow = { 0 };
			NewRow.cValues = tab->row.cValues;
			NewRow.ulAdrEntryPad = tab->row.ulAdrEntryPad;
			EC_MAPI(ScDupPropset(
				tab->row.cValues,
				tab->row.lpProps,
				MAPIAllocateBuffer,
				&NewRow.lpProps));

			DebugPrintEx(DBGGeneric, CLASS, L"msgOnAddItem", L"Received message to add row to row %d\n", iNewRow);
			EC_H(AddItemToListBox(iNewRow, &NewRow));

			return hRes;
		}

		// WM_MFCMAPI_DELETEITEM
		_Check_return_ LRESULT CContentsTableListCtrl::msgOnDeleteItem(WPARAM wParam, LPARAM /*lParam*/)
		{
			auto hRes = S_OK;
			auto tab = reinterpret_cast<TABLE_NOTIFICATION*>(wParam);

			if (!tab) return MAPI_E_INVALID_PARAMETER;

			const auto iItem = FindRow(&tab->propIndex.Value.bin);

			DebugPrintEx(DBGGeneric, CLASS, L"msgOnDeleteItem", L"Received message to delete item 0x%d\n", iItem);

			if (iItem == -1) return S_OK;

			EC_B(DeleteItem(iItem));

			if (S_OK != hRes || !m_lpHostDlg) return hRes;

			const auto iCount = GetItemCount();
			if (iCount == 0)
			{
				m_lpHostDlg->OnUpdateSingleMAPIPropListCtrl(nullptr, nullptr);
			}

			m_lpHostDlg->UpdateStatusBarText(STATUSDATA1, IDS_STATUSTEXTNUMITEMS, iCount);
			return hRes;
		}

		// WM_MFCMAPI_MODIFYITEM
		_Check_return_ LRESULT CContentsTableListCtrl::msgOnModifyItem(WPARAM wParam, LPARAM /*lParam*/)
		{
			auto hRes = S_OK;
			auto tab = reinterpret_cast<TABLE_NOTIFICATION*>(wParam);

			if (!tab) return MAPI_E_INVALID_PARAMETER;

			const auto iItem = FindRow(&tab->propIndex.Value.bin);

			if (-1 != iItem)
			{
				DebugPrintEx(DBGGeneric, CLASS, L"msgOnModifyItem", L"Received message to modify row %d with %p\n", iItem, &tab->row);

				// We make this copy here and pass it in to RefreshItem, where it is grabbed by SortListData::InitializeContents to be part of the item data
				// The mem will be freed when the item data is cleaned up - do not free here
				SRow NewRow = { 0 };
				NewRow.cValues = tab->row.cValues;
				NewRow.ulAdrEntryPad = tab->row.ulAdrEntryPad;
				EC_MAPI(ScDupPropset(
					tab->row.cValues,
					tab->row.lpProps,
					MAPIAllocateBuffer,
					&NewRow.lpProps));

				EC_H(RefreshItem(iItem, &NewRow, true));
			}

			return hRes;
		}

		// WM_MFCMAPI_REFRESHTABLE
		_Check_return_ LRESULT CContentsTableListCtrl::msgOnRefreshTable(WPARAM /*wParam*/, LPARAM /*lParam*/)
		{
			auto hRes = S_OK;
			DebugPrintEx(DBGGeneric, CLASS, L"msgOnRefreshTable", L"Received message refresh table\n");
			EC_H(RefreshTable());

			return hRes;
		}

		// This function steps through the list control to find the entry with this instance key
		// return -1 if item not found
		_Check_return_ int CContentsTableListCtrl::FindRow(_In_ LPSBinary lpInstance) const
		{
			DebugPrintEx(DBGGeneric, CLASS, L"msgOnGetIndex", L"Getting index for %p\n", lpInstance);

			if (!lpInstance) return -1;

			auto iItem = 0;
			for (iItem = 0; iItem < GetItemCount(); iItem++)
			{
				const auto lpListData = GetSortListData(iItem);
				if (lpListData && lpListData->Contents())
				{
					const auto lpCurInstance = lpListData->Contents()->m_lpInstanceKey;
					if (lpCurInstance)
					{
						if (!memcmp(lpCurInstance->lpb, lpInstance->lpb, lpInstance->cb))
						{
							DebugPrintEx(DBGGeneric, CLASS, L"msgOnGetIndex", L"Matched at 0x%08X\n", iItem);
							return iItem;
						}
					}
				}
			}

			DebugPrintEx(DBGGeneric, CLASS, L"msgOnGetIndex", L"No match found: 0x%08X\n", iItem);
			return -1;
		}
	}
}