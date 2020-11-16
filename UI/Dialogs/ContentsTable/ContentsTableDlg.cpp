#include <StdAfx.h>
#include <UI/Dialogs/ContentsTable/ContentsTableDlg.h>
#include <UI/Controls/SortList/ContentsTableListCtrl.h>
#include <UI/FakeSplitter.h>
#include <UI/file/FileDialogEx.h>
#include <UI/Controls/SortList/SingleMAPIPropListCtrl.h>
#include <core/mapi/cache/mapiObjects.h>
#include <UI/Dialogs/MFCUtilityFunctions.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Dialogs/Editors/RestrictEditor.h>
#include <UI/Dialogs/Editors/PropertyTagEditor.h>
#include <UI/Dialogs/Editors/SearchEditor.h>
#include <core/mapi/mapiMemory.h>
#include <core/mapi/extraPropTags.h>
#include <UI/addinui.h>
#include <core/utility/registry.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFunctions.h>

namespace dialog
{
	static std::wstring CLASS = L"CContentsTableDlg";

	CContentsTableDlg::CContentsTableDlg(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		UINT uidTitle,
		createDialogType bCreateDialog,
		_In_opt_ LPMAPIPROP lpContainer,
		_In_opt_ LPMAPITABLE lpContentsTable,
		_In_ LPSPropTagArray sptDefaultDisplayColumnTags,
		_In_ const std::vector<columns::TagNames>& lpDefaultDisplayColumns,
		ULONG nIDContextMenu,
		ULONG ulAddInContext)
		: CBaseDialog(pParentWnd, lpMapiObjects, ulAddInContext)
	{
		TRACE_CONSTRUCTOR(CLASS);
		if (NULL != uidTitle)
		{
			m_szTitle = strings::loadstring(uidTitle);
		}
		else
		{
			m_szTitle = strings::loadstring(IDS_TABLEASCONTENTS);
		}

		m_lpContentsTableListCtrl = nullptr;
		m_lpContainer = mapi::safe_cast<LPMAPICONTAINER>(lpContainer);
		m_nIDContextMenu = nIDContextMenu;

		m_displayFlags = tableDisplayFlags::dfNormal;

		m_lpContentsTable = lpContentsTable;
		if (m_lpContentsTable) m_lpContentsTable->AddRef();

		m_sptDefaultDisplayColumnTags = sptDefaultDisplayColumnTags;
		m_lpDefaultDisplayColumns = lpDefaultDisplayColumns;

		if (createDialogType::CALL_CREATE_DIALOG == bCreateDialog)
		{
			CContentsTableDlg::CreateDialogAndMenu(NULL);
		}
	}

	CContentsTableDlg::~CContentsTableDlg()
	{
		TRACE_DESTRUCTOR(CLASS);
		if (m_lpContentsTable) m_lpContentsTable->Release();
		if (m_lpContainer) m_lpContainer->Release();
	}

	_Check_return_ bool CContentsTableDlg::HandleMenu(WORD wMenuSelect)
	{
		output::DebugPrint(
			output::dbgLevel::Menu,
			L"CContentsTableDlg::HandleMenu wMenuSelect = 0x%X = %u\n",
			wMenuSelect,
			wMenuSelect);
		switch (wMenuSelect)
		{
		case ID_APPLYFINDROW:
			SetRestrictionType(restrictionType::findrow);
			return true;
		case ID_APPLYRESTRICTION:
			SetRestrictionType(restrictionType::normal);
			return true;
		case ID_CLEARRESTRICTION:
			SetRestrictionType(restrictionType::none);
			return true;
		}

		return CBaseDialog::HandleMenu(wMenuSelect);
	}

	BOOL CContentsTableDlg::OnInitDialog()
	{
		const auto bRet = CBaseDialog::OnInitDialog();

		m_lpContentsTableListCtrl = new controls::sortlistctrl::CContentsTableListCtrl(
			&m_lpFakeSplitter,
			m_lpMapiObjects,
			m_sptDefaultDisplayColumnTags,
			m_lpDefaultDisplayColumns,
			m_nIDContextMenu,
			m_bIsAB,
			this);

		if (m_lpContentsTableListCtrl)
		{
			m_lpFakeSplitter.SetPaneOne(m_lpContentsTableListCtrl->GetSafeHwnd());
			m_lpFakeSplitter.SetPercent(static_cast<FLOAT>(0.40));
			m_lpFakeSplitter.SetSplitType(controls::splitType::vertical);
		}

		if (m_lpContainer)
		{
			// Get a property for the title bar
			m_szTitle = mapi::GetTitle(m_lpContainer);

			if (m_displayFlags && tableDisplayFlags::dfAssoc)
			{
				m_szTitle = strings::formatmessage(IDS_HIDDEN, m_szTitle.c_str());
			}

			if (m_lpContentsTable) m_lpContentsTable->Release();
			m_lpContentsTable = nullptr;

			const auto unicodeFlag = registry::preferUnicodeProps ? MAPI_UNICODE : fMapiUnicode;

			const auto ulFlags = (m_displayFlags && tableDisplayFlags::dfAssoc ? MAPI_ASSOCIATED : NULL) |
								 (m_displayFlags && tableDisplayFlags::dfDeleted ? SHOW_SOFT_DELETES : NULL) |
								 unicodeFlag;

			// Get the table of contents of the IMAPIContainer!!!
			EC_MAPI_S(m_lpContainer->GetContentsTable(ulFlags, &m_lpContentsTable));
		}

		UpdateTitleBarText();

		return bRet;
	}

	void CContentsTableDlg::CreateDialogAndMenu(UINT nIDMenuResource)
	{
		output::DebugPrintEx(
			output::dbgLevel::CreateDialog, CLASS, L"CreateDialogAndMenu", L"id = 0x%X\n", nIDMenuResource);
		CBaseDialog::CreateDialogAndMenu(nIDMenuResource, IDR_MENU_TABLE, IDS_TABLEMENU);

		if (m_lpContentsTableListCtrl && m_lpContentsTable)
		{
			const auto ulPropType = mapi::GetMAPIObjectType(m_lpContainer);

			// Pass the contents table to the list control, but don't render yet - call BuildUIForContentsTable from CreateDialogAndMenu for that
			m_lpContentsTableListCtrl->SetContentsTable(m_lpContentsTable, m_displayFlags, ulPropType);
		}
	}

	BEGIN_MESSAGE_MAP(CContentsTableDlg, CBaseDialog)
	ON_COMMAND(ID_DISPLAYSELECTEDITEM, OnDisplayItem)
	ON_COMMAND(ID_CANCELTABLELOAD, OnEscHit)
	ON_COMMAND(ID_REFRESHVIEW, OnRefreshView)
	ON_COMMAND(ID_CREATEPROPERTYSTRINGRESTRICTION, OnCreatePropertyStringRestriction)
	ON_COMMAND(ID_CREATERANGERESTRICTION, OnCreateRangeRestriction)
	ON_COMMAND(ID_EDITRESTRICTION, OnEditRestriction)
	ON_COMMAND(ID_GETSTATUS, OnGetStatus)
	ON_COMMAND(ID_OUTPUTTABLE, OnOutputTable)
	ON_COMMAND(ID_SETCOLUMNS, OnSetColumns)
	ON_COMMAND(ID_SORTTABLE, OnSortTable)
	ON_COMMAND(ID_TABLENOTIFICATIONON, OnNotificationOn)
	ON_COMMAND(ID_TABLENOTIFICATIONOFF, OnNotificationOff)
	ON_MESSAGE(WM_MFCMAPI_RESETCOLUMNS, msgOnResetColumns)
	END_MESSAGE_MAP()

	void CContentsTableDlg::OnInitMenu(_In_opt_ CMenu* pMenu)
	{
		if (m_lpContentsTableListCtrl && pMenu)
		{
			const int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();

			pMenu->EnableMenuItem(ID_CANCELTABLELOAD, DIM(m_lpContentsTableListCtrl->IsLoading()));

			pMenu->EnableMenuItem(ID_DISPLAYSELECTEDITEM, DIMMSOK(iNumSel));

			const auto RestrictionType = m_lpContentsTableListCtrl->GetRestrictionType();
			pMenu->CheckMenuItem(ID_APPLYFINDROW, CHECK(RestrictionType == restrictionType::findrow));
			pMenu->CheckMenuItem(ID_APPLYRESTRICTION, CHECK(RestrictionType == restrictionType::normal));
			pMenu->CheckMenuItem(ID_CLEARRESTRICTION, CHECK(RestrictionType == restrictionType::none));
			pMenu->EnableMenuItem(
				ID_TABLENOTIFICATIONON,
				DIM(m_lpContentsTableListCtrl->IsContentsTableSet() && !m_lpContentsTableListCtrl->IsAdviseSet()));
			pMenu->CheckMenuItem(ID_TABLENOTIFICATIONON, CHECK(m_lpContentsTableListCtrl->IsAdviseSet()));
			pMenu->EnableMenuItem(ID_TABLENOTIFICATIONOFF, DIM(m_lpContentsTableListCtrl->IsAdviseSet()));
			pMenu->EnableMenuItem(
				ID_OUTPUTTABLE,
				DIM(m_lpContentsTableListCtrl->IsContentsTableSet() && !m_lpContentsTableListCtrl->IsLoading()));

			pMenu->EnableMenuItem(ID_SETCOLUMNS, DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
			pMenu->EnableMenuItem(ID_SORTTABLE, DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
			pMenu->EnableMenuItem(ID_GETSTATUS, DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
			pMenu->EnableMenuItem(
				ID_CREATEPROPERTYSTRINGRESTRICTION, DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
			pMenu->EnableMenuItem(ID_CREATERANGERESTRICTION, DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
			pMenu->EnableMenuItem(ID_EDITRESTRICTION, DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
			pMenu->EnableMenuItem(ID_APPLYFINDROW, DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
			pMenu->EnableMenuItem(ID_APPLYRESTRICTION, DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
			pMenu->EnableMenuItem(ID_CLEARRESTRICTION, DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
			pMenu->EnableMenuItem(ID_REFRESHVIEW, DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));

			for (ULONG ulMenu = ID_ADDINMENU; ulMenu < ID_ADDINMENU + m_ulAddInMenuItems; ulMenu++)
			{
				const auto lpAddInMenu = ui::addinui::GetAddinMenuItem(m_hWnd, ulMenu);
				if (!lpAddInMenu) continue;

				const auto ulFlags = lpAddInMenu->ulFlags;
				UINT uiEnable = MF_ENABLED;

				if ((ulFlags & MENU_FLAGS_SINGLESELECT) && iNumSel != 1) uiEnable = MF_GRAYED;
				if ((ulFlags & MENU_FLAGS_MULTISELECT) && !iNumSel) uiEnable = MF_GRAYED;
				EnableAddInMenus(pMenu->m_hMenu, ulMenu, lpAddInMenu, uiEnable);
			}
		}
		CBaseDialog::OnInitMenu(pMenu);
	}

	void CContentsTableDlg::OnCancel()
	{
		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"OnCancel", L"\n");
		// get rid of the window before we start our cleanup
		ShowWindow(SW_HIDE);

		if (m_lpContentsTableListCtrl) m_lpContentsTableListCtrl->OnCancelTableLoad();

		if (m_lpContentsTableListCtrl) m_lpContentsTableListCtrl->Release();
		m_lpContentsTableListCtrl = nullptr;
		CBaseDialog::OnCancel();
	}

	void CContentsTableDlg::OnEscHit()
	{
		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"OnEscHit", L"\n");
		if (m_lpContentsTableListCtrl)
		{
			m_lpContentsTableListCtrl->OnCancelTableLoad();
		}
	}

	void CContentsTableDlg::SetRestrictionType(restrictionType RestrictionType)
	{
		if (m_lpContentsTableListCtrl) m_lpContentsTableListCtrl->SetRestrictionType(RestrictionType);
		OnRefreshView();
	}

	void CContentsTableDlg::OnDisplayItem()
	{
		auto iItem = -1;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		do
		{
			auto lpMAPIProp = m_lpContentsTableListCtrl->OpenNextSelectedItemProp(&iItem, modifyType::REQUEST_MODIFY);

			if (lpMAPIProp)
			{
				EC_H_S(DisplayObject(lpMAPIProp, NULL, objectType::hierarchy, this));
				lpMAPIProp->Release();
			}
		} while (iItem != -1);
	}

	// Clear the current list and get a new one with whatever code we've got in LoadMAPIPropList
	void CContentsTableDlg::OnRefreshView()
	{
		if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;
		output::DebugPrintEx(output::dbgLevel::Generic, CLASS, L"OnRefreshView", L"\n");
		if (m_lpContentsTableListCtrl->IsLoading()) m_lpContentsTableListCtrl->OnCancelTableLoad();
		m_lpContentsTableListCtrl->RefreshTable();
	}

	void CContentsTableDlg::OnNotificationOn()
	{
		if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;
		m_lpContentsTableListCtrl->NotificationOn();
	}

	void CContentsTableDlg::OnNotificationOff()
	{
		if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;
		m_lpContentsTableListCtrl->NotificationOff();
	}

	void CContentsTableDlg::OnCreatePropertyStringRestriction()
	{
		if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;

		editor::CSearchEditor SearchEditor(PR_SUBJECT_W, m_lpContainer, this);
		if (SearchEditor.DisplayDialog())
		{
			const auto lpRes = SearchEditor.GetRestriction();
			if (lpRes)
			{
				m_lpContentsTableListCtrl->SetRestriction(lpRes);

				if (SearchEditor.GetCheck(editor::CSearchEditor::SearchFields::FINDROW))
				{
					SetRestrictionType(restrictionType::findrow);
				}
				else
				{
					SetRestrictionType(restrictionType::normal);
				}
			}
		}
	}

	void CContentsTableDlg::OnCreateRangeRestriction()
	{
		if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;

		editor::CPropertyTagEditor MyPropertyTag(
			IDS_CREATERANGERES,
			NULL, // prompt
			PR_SUBJECT,
			m_bIsAB,
			m_lpContainer,
			this);

		if (!MyPropertyTag.DisplayDialog()) return;
		editor::CEditor MyData(
			this, IDS_SEARCHCRITERIA, IDS_RANGESEARCHCRITERIAPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_SUBSTRING, false));
		MyData.AddPane(viewpane::CheckPane::Create(1, IDS_APPLYUSINGFINDROW, false, false));

		if (!MyData.DisplayDialog()) return;

		const auto szString = MyData.GetStringW(0);
		// Allocate and create our SRestriction
		const auto lpRes = mapi::CreateRangeRestriction(
			CHANGE_PROP_TYPE(MyPropertyTag.GetPropertyTag(), PT_UNICODE), szString, nullptr);

		m_lpContentsTableListCtrl->SetRestriction(lpRes);

		if (MyData.GetCheck(1))
		{
			SetRestrictionType(restrictionType::findrow);
		}
		else
		{
			SetRestrictionType(restrictionType::normal);
		}
	}

	void CContentsTableDlg::OnEditRestriction()
	{
		if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;

		editor::CRestrictEditor MyRestrict(
			this,
			nullptr, // No alloc parent - we must MAPIFreeBuffer the result
			m_lpContentsTableListCtrl->GetRestriction());

		if (!MyRestrict.DisplayDialog()) return;

		m_lpContentsTableListCtrl->SetRestriction(MyRestrict.DetachModifiedSRestriction());
	}

	void CContentsTableDlg::OnOutputTable()
	{
		if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;
		auto szFileName = file::CFileDialogExW::SaveAs(
			L"txt", // STRING_OK
			L"table.txt", // STRING_OK
			OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
			strings::loadstring(IDS_TEXTFILES),
			this);
		if (!szFileName.empty())
		{
			output::DebugPrintEx(
				output::dbgLevel::Generic, CLASS, L"OnOutputTable", L"saving to %ws\n", szFileName.c_str());
			m_lpContentsTableListCtrl->OnOutputTable(szFileName);
		}
	}

	void CContentsTableDlg::OnSetColumns()
	{
		if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;
		m_lpContentsTableListCtrl->DoSetColumns(false, true);
	}

	void CContentsTableDlg::OnGetStatus()
	{
		if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;
		m_lpContentsTableListCtrl->GetStatus();
	}

	void CContentsTableDlg::OnSortTable()
	{
		if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;

		editor::CEditor MyData(this, IDS_SORTTABLE, IDS_SORTTABLEPROMPT1, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_CSORTS, false));
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_CCATS, false));
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(2, IDS_CEXPANDED, false));
		MyData.AddPane(viewpane::CheckPane::Create(3, IDS_TBLASYNC, false, false));
		MyData.AddPane(viewpane::CheckPane::Create(4, IDS_TBLBATCH, false, false));
		MyData.AddPane(viewpane::CheckPane::Create(5, IDS_REFRESHAFTERSORT, true, false));

		if (!MyData.DisplayDialog()) return;

		const auto cSorts = MyData.GetDecimal(0);
		const auto cCategories = MyData.GetDecimal(1);
		const auto cExpanded = MyData.GetDecimal(2);

		if (cSorts < cCategories || cCategories < cExpanded)
		{
			error::ErrDialog(__FILE__, __LINE__, IDS_EDBADSORTS, cSorts, cCategories, cExpanded);
			return;
		}

		auto lpMySortOrders = mapi::allocate<LPSSortOrderSet>(CbNewSSortOrderSet(cSorts));
		if (lpMySortOrders)
		{
			lpMySortOrders->cSorts = cSorts;
			lpMySortOrders->cCategories = cCategories;
			lpMySortOrders->cExpanded = cExpanded;
			auto bNoError = true;

			for (ULONG i = 0; i < cSorts; i++)
			{
				editor::CPropertyTagEditor MyPropertyTag(
					NULL, // title
					NULL, // prompt
					NULL,
					m_bIsAB,
					m_lpContainer,
					this);

				if (MyPropertyTag.DisplayDialog())
				{
					lpMySortOrders->aSort[i].ulPropTag = MyPropertyTag.GetPropertyTag();
					editor::CEditor MySortOrderDlg(
						this, IDS_SORTORDER, IDS_SORTORDERPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
					UINT uidDropDown[] = {
						IDS_DDTABLESORTASCEND,
						IDS_DDTABLESORTDESCEND,
						IDS_DDTABLESORTCOMBINE,
						IDS_DDTABLESORTCATEGMAX,
						IDS_DDTABLESORTCATEGMIN};
					MySortOrderDlg.AddPane(
						viewpane::DropDownPane::Create(0, IDS_SORTORDER, _countof(uidDropDown), uidDropDown, true));

					if (MySortOrderDlg.DisplayDialog())
					{
						switch (MySortOrderDlg.GetDropDown(0))
						{
						default:
						case 0:
							lpMySortOrders->aSort[i].ulOrder = TABLE_SORT_ASCEND;
							break;
						case 1:
							lpMySortOrders->aSort[i].ulOrder = TABLE_SORT_DESCEND;
							break;
						case 2:
							lpMySortOrders->aSort[i].ulOrder = TABLE_SORT_COMBINE;
							break;
						case 3:
							lpMySortOrders->aSort[i].ulOrder = TABLE_SORT_CATEG_MAX;
							break;
						case 4:
							lpMySortOrders->aSort[i].ulOrder = TABLE_SORT_CATEG_MIN;
							break;
						}
					}
					else
						bNoError = false;
				}
				else
					bNoError = false;
			}

			if (bNoError)
			{
				m_lpContentsTableListCtrl->SetSortTable(
					lpMySortOrders,
					(MyData.GetCheck(3) ? TBL_ASYNC : 0) | (MyData.GetCheck(4) ? TBL_BATCH : 0) // flags
				);
			}
		}

		MAPIFreeBuffer(lpMySortOrders);

		if (MyData.GetCheck(5)) m_lpContentsTableListCtrl->RefreshTable();
	}

	// Since the strategy for opening the selected property may vary depending on the table we're displaying,
	// this virtual function allows us to override the default method with the method used by the table we've written a special class for.
	_Check_return_ LPMAPIPROP CContentsTableDlg::OpenItemProp(int iSelectedItem, modifyType bModify)
	{
		if (!m_lpContentsTableListCtrl) return nullptr;
		output::DebugPrintEx(
			output::dbgLevel::OpenItemProp, CLASS, L"OpenItemProp", L"iSelectedItem = 0x%X\n", iSelectedItem);

		if (-1 == iSelectedItem)
		{
			// Get the first selected item
			return m_lpContentsTableListCtrl->OpenNextSelectedItemProp(nullptr, bModify);
		}
		else
		{
			return m_lpContentsTableListCtrl->DefaultOpenItemProp(iSelectedItem, bModify);
		}
	}

	_Check_return_ HRESULT CContentsTableDlg::OpenAttachmentsFromMessage(_In_ LPMESSAGE lpMessage)
	{
		if (lpMessage == nullptr) return MAPI_E_INVALID_PARAMETER;

		return DisplayTable(lpMessage, PR_MESSAGE_ATTACHMENTS, objectType::otDefault, this);
	}

	_Check_return_ HRESULT CContentsTableDlg::OpenRecipientsFromMessage(_In_ LPMESSAGE lpMessage)
	{
		return DisplayTable(lpMessage, PR_MESSAGE_RECIPIENTS, objectType::otDefault, this);
	}

	_Check_return_ bool CContentsTableDlg::HandleAddInMenu(WORD wMenuSelect)
	{
		if (wMenuSelect < ID_ADDINMENU || ID_ADDINMENU + m_ulAddInMenuItems < wMenuSelect) return false;
		if (!m_lpContentsTableListCtrl) return false;
		LPMAPIPROP lpMAPIProp = nullptr;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		const auto lpAddInMenu = ui::addinui::GetAddinMenuItem(m_hWnd, wMenuSelect);
		if (!lpAddInMenu) return false;

		const auto ulFlags = lpAddInMenu->ulFlags;

		const auto fRequestModify =
			(ulFlags & MENU_FLAGS_REQUESTMODIFY) ? modifyType::REQUEST_MODIFY : modifyType::DO_NOT_REQUEST_MODIFY;

		// Get the stuff we need for any case
		_AddInMenuParams MyAddInMenuParams = {nullptr};
		MyAddInMenuParams.lpAddInMenu = lpAddInMenu;
		MyAddInMenuParams.ulAddInContext = m_ulAddInContext;
		MyAddInMenuParams.hWndParent = m_hWnd;
		if (m_lpMapiObjects)
		{
			MyAddInMenuParams.lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
			MyAddInMenuParams.lpMDB = m_lpMapiObjects->GetMDB(); // do not release
			MyAddInMenuParams.lpAdrBook = m_lpMapiObjects->GetAddrBook(false); // do not release
		}

		if (m_lpPropDisplay)
		{
			m_lpPropDisplay->GetSelectedPropTag(&MyAddInMenuParams.ulPropTag);
		}

		// MENU_FLAGS_SINGLESELECT and MENU_FLAGS_MULTISELECT can't both be set, so we can ignore this case
		if (!(ulFlags & (MENU_FLAGS_SINGLESELECT | MENU_FLAGS_MULTISELECT)))
		{
			HandleAddInMenuSingle(&MyAddInMenuParams, nullptr, nullptr);
		}
		else
		{
			// Add appropriate flag to context
			MyAddInMenuParams.ulCurrentFlags |= ulFlags & (MENU_FLAGS_SINGLESELECT | MENU_FLAGS_MULTISELECT);
			auto items = m_lpContentsTableListCtrl->GetSelectedItemNums();
			for (const auto& item : items)
			{
				SRow MyRow = {0};

				const auto lpData = m_lpContentsTableListCtrl->GetSortListData(item);
				// If we have a row to give, give it - it's free
				if (lpData)
				{
					MyRow.cValues = lpData->cSourceProps;
					MyRow.lpProps = lpData->lpSourceProps;
					MyAddInMenuParams.lpRow = &MyRow;
					MyAddInMenuParams.ulCurrentFlags |= MENU_FLAGS_ROW;
				}

				if (!(ulFlags & MENU_FLAGS_ROW))
				{
					lpMAPIProp = OpenItemProp(item, fRequestModify);
				}

				HandleAddInMenuSingle(&MyAddInMenuParams, lpMAPIProp, nullptr);
				if (lpMAPIProp) lpMAPIProp->Release();

				// If we're not doing multiselect, then we're done after a single pass
				if (!(ulFlags & MENU_FLAGS_MULTISELECT)) break;
			}
		}
		return true;
	}

	void CContentsTableDlg::HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_opt_ LPMAPIPROP lpMAPIProp,
		_In_opt_ LPMAPICONTAINER /*lpContainer*/)
	{
		if (lpParams)
		{
			lpParams->lpTable = m_lpContentsTable;
			switch (lpParams->ulAddInContext)
			{
			case MENU_CONTEXT_RECIEVE_FOLDER_TABLE:
				lpParams->lpFolder = mapi::safe_cast<LPMAPIFOLDER>(lpMAPIProp);
				break;
			case MENU_CONTEXT_HIER_TABLE:
				lpParams->lpFolder = mapi::safe_cast<LPMAPIFOLDER>(lpMAPIProp);
				break;
			}
		}

		ui::addinui::InvokeAddInMenu(lpParams);

		if (lpParams && lpParams->lpFolder)
		{
			lpParams->lpFolder->Release();
			lpParams->lpFolder = nullptr;
		}
	}

	// WM_MFCMAPI_RESETCOLUMNS
	// Returns true if we reset columns, false otherwise
	_Check_return_ LRESULT CContentsTableDlg::msgOnResetColumns(WPARAM /*wParam*/, LPARAM /*lParam*/)
	{
		output::DebugPrintEx(
			output::dbgLevel::Generic, CLASS, L"msgOnResetColumns", L"Received message reset columns\n");

		if (m_lpContentsTableListCtrl)
		{
			m_lpContentsTableListCtrl->DoSetColumns(true, false);
			return true;
		}

		return false;
	}
} // namespace dialog
