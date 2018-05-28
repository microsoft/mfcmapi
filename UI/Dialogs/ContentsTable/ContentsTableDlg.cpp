#include "StdAfx.h"
#include <UI/Dialogs/ContentsTable/ContentsTableDlg.h>
#include <UI/Controls/ContentsTableListCtrl.h>
#include <UI/FakeSplitter.h>
#include <UI/FileDialogEx.h>
#include <UI/Controls/SingleMAPIPropListCtrl.h>
#include <MAPI/MAPIFunctions.h>
#include <MAPI/MapiObjects.h>
#include <UI/MFCUtilityFunctions.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Dialogs/Editors/RestrictEditor.h>
#include <UI/Dialogs/Editors/PropertyTagEditor.h>
#include <UI/Dialogs/Editors/SearchEditor.h>

static std::wstring CLASS = L"CContentsTableDlg";

CContentsTableDlg::CContentsTableDlg(
	_In_ CParentWnd* pParentWnd,
	_In_ CMapiObjects* lpMapiObjects,
	UINT uidTitle,
	__mfcmapiCreateDialogEnum bCreateDialog,
	_In_opt_ LPMAPITABLE lpContentsTable,
	_In_ LPSPropTagArray sptExtraColumnTags,
	_In_ const std::vector<TagNames>& lpExtraDisplayColumns,
	ULONG nIDContextMenu,
	ULONG ulAddInContext
) :
	CBaseDialog(
		pParentWnd,
		lpMapiObjects,
		ulAddInContext)
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
	m_lpContainer = nullptr;
	m_nIDContextMenu = nIDContextMenu;

	m_ulDisplayFlags = dfNormal;

	m_lpContentsTable = lpContentsTable;
	if (m_lpContentsTable) m_lpContentsTable->AddRef();

	m_sptExtraColumnTags = sptExtraColumnTags;
	m_lpExtraDisplayColumns = lpExtraDisplayColumns;

	if (mfcmapiCALL_CREATE_DIALOG == bCreateDialog)
	{
		CContentsTableDlg::CreateDialogAndMenu(NULL);
	}
}

CContentsTableDlg::~CContentsTableDlg()
{
	TRACE_DESTRUCTOR(CLASS);
	if (m_lpContentsTable) m_lpContentsTable->Release();
	m_lpContentsTable = nullptr;
}

_Check_return_ bool CContentsTableDlg::HandleMenu(WORD wMenuSelect)
{
	DebugPrint(DBGMenu, L"CContentsTableDlg::HandleMenu wMenuSelect = 0x%X = %u\n", wMenuSelect, wMenuSelect);
	switch (wMenuSelect)
	{
	case ID_APPLYFINDROW: SetRestrictionType(mfcmapiFINDROW_RESTRICTION); return true;
	case ID_APPLYRESTRICTION: SetRestrictionType(mfcmapiNORMAL_RESTRICTION); return true;
	case ID_CLEARRESTRICTION: SetRestrictionType(mfcmapiNO_RESTRICTION); return true;
	}

	return CBaseDialog::HandleMenu(wMenuSelect);
}

BOOL CContentsTableDlg::OnInitDialog()
{
	const auto bRet = CBaseDialog::OnInitDialog();

	m_lpContentsTableListCtrl = new CContentsTableListCtrl(
		m_lpFakeSplitter,
		m_lpMapiObjects,
		m_sptExtraColumnTags,
		m_lpExtraDisplayColumns,
		m_nIDContextMenu,
		m_bIsAB,
		this);

	if (m_lpContentsTableListCtrl && m_lpFakeSplitter)
	{
		m_lpFakeSplitter->SetPaneOne(m_lpContentsTableListCtrl);
		m_lpFakeSplitter->SetPercent(static_cast<FLOAT>(0.40));
		m_lpFakeSplitter->SetSplitType(SplitVertical);
	}

	if (m_lpContainer)
	{
		// Get a property for the title bar
		m_szTitle = GetTitle(m_lpContainer);

		if (m_ulDisplayFlags & dfAssoc)
		{
			m_szTitle = strings::formatmessage(IDS_HIDDEN, m_szTitle.c_str());
		}

		if (m_lpContentsTable) m_lpContentsTable->Release();
		m_lpContentsTable = nullptr;

		const auto ulFlags =
			(m_ulDisplayFlags & dfAssoc ? MAPI_ASSOCIATED : NULL) |
			(m_ulDisplayFlags & dfDeleted ? SHOW_SOFT_DELETES : NULL) |
			fMapiUnicode;

		auto hRes = S_OK;
		// Get the table of contents of the IMAPIContainer!!!
		EC_MAPI(m_lpContainer->GetContentsTable(
			ulFlags,
			&m_lpContentsTable));
	}

	UpdateTitleBarText();

	return bRet;
}

void CContentsTableDlg::CreateDialogAndMenu(UINT nIDMenuResource)
{
	auto hRes = S_OK;

	DebugPrintEx(DBGCreateDialog, CLASS, L"CreateDialogAndMenu", L"id = 0x%X\n", nIDMenuResource);
	CBaseDialog::CreateDialogAndMenu(nIDMenuResource, IDR_MENU_TABLE, IDS_TABLEMENU);

	if (m_lpContentsTableListCtrl && m_lpContentsTable)
	{
		const auto ulPropType = GetMAPIObjectType(m_lpContainer);

		// Pass the contents table to the list control, but don't render yet - call BuildUIForContentsTable from CreateDialogAndMenu for that
		WC_H(m_lpContentsTableListCtrl->SetContentsTable(
			m_lpContentsTable,
			m_ulDisplayFlags,
			ulPropType));
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
		pMenu->CheckMenuItem(ID_APPLYFINDROW, CHECK(mfcmapiFINDROW_RESTRICTION == RestrictionType));
		pMenu->CheckMenuItem(ID_APPLYRESTRICTION, CHECK(mfcmapiNORMAL_RESTRICTION == RestrictionType));
		pMenu->CheckMenuItem(ID_CLEARRESTRICTION, CHECK(mfcmapiNO_RESTRICTION == RestrictionType));
		pMenu->EnableMenuItem(ID_TABLENOTIFICATIONON, DIM(m_lpContentsTableListCtrl->IsContentsTableSet() && !m_lpContentsTableListCtrl->IsAdviseSet()));
		pMenu->CheckMenuItem(ID_TABLENOTIFICATIONON, CHECK(m_lpContentsTableListCtrl->IsAdviseSet()));
		pMenu->EnableMenuItem(ID_TABLENOTIFICATIONOFF, DIM(m_lpContentsTableListCtrl->IsAdviseSet()));
		pMenu->EnableMenuItem(ID_OUTPUTTABLE, DIM(m_lpContentsTableListCtrl->IsContentsTableSet() && !m_lpContentsTableListCtrl->IsLoading()));

		pMenu->EnableMenuItem(ID_SETCOLUMNS, DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
		pMenu->EnableMenuItem(ID_SORTTABLE, DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
		pMenu->EnableMenuItem(ID_GETSTATUS, DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
		pMenu->EnableMenuItem(ID_CREATEPROPERTYSTRINGRESTRICTION, DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
		pMenu->EnableMenuItem(ID_CREATERANGERESTRICTION, DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
		pMenu->EnableMenuItem(ID_EDITRESTRICTION, DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
		pMenu->EnableMenuItem(ID_APPLYFINDROW, DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
		pMenu->EnableMenuItem(ID_APPLYRESTRICTION, DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
		pMenu->EnableMenuItem(ID_CLEARRESTRICTION, DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));
		pMenu->EnableMenuItem(ID_REFRESHVIEW, DIM(m_lpContentsTableListCtrl->IsContentsTableSet()));

		for (ULONG ulMenu = ID_ADDINMENU; ulMenu < ID_ADDINMENU + m_ulAddInMenuItems; ulMenu++)
		{
			const auto lpAddInMenu = GetAddinMenuItem(m_hWnd, ulMenu);
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
	DebugPrintEx(DBGGeneric, CLASS, L"OnCancel", L"\n");
	// get rid of the window before we start our cleanup
	ShowWindow(SW_HIDE);

	if (m_lpContentsTableListCtrl) m_lpContentsTableListCtrl->OnCancelTableLoad();

	if (m_lpContentsTableListCtrl) m_lpContentsTableListCtrl->Release();
	m_lpContentsTableListCtrl = nullptr;
	CBaseDialog::OnCancel();
}

void CContentsTableDlg::OnEscHit()
{
	DebugPrintEx(DBGGeneric, CLASS, L"OnEscHit", L"\n");
	if (m_lpContentsTableListCtrl)
	{
		m_lpContentsTableListCtrl->OnCancelTableLoad();
	}
}

void CContentsTableDlg::SetRestrictionType(__mfcmapiRestrictionTypeEnum RestrictionType)
{
	if (m_lpContentsTableListCtrl)
		m_lpContentsTableListCtrl->SetRestrictionType(RestrictionType);
	OnRefreshView();
}

void CContentsTableDlg::OnDisplayItem()
{
	LPMAPIPROP lpMAPIProp = nullptr;
	auto iItem = -1;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	do
	{
		auto hRes = S_OK;
		EC_H(m_lpContentsTableListCtrl->OpenNextSelectedItemProp(
			&iItem,
			mfcmapiREQUEST_MODIFY,
			&lpMAPIProp));

		if (lpMAPIProp)
		{
			EC_H(DisplayObject(
				lpMAPIProp,
				NULL,
				otHierarchy,
				this));
			lpMAPIProp->Release();
			lpMAPIProp = nullptr;
		}
	} while (iItem != -1);
}

// Clear the current list and get a new one with whatever code we've got in LoadMAPIPropList
void CContentsTableDlg::OnRefreshView()
{
	auto hRes = S_OK;
	if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;
	DebugPrintEx(DBGGeneric, CLASS, L"OnRefreshView", L"\n");
	if (m_lpContentsTableListCtrl->IsLoading()) m_lpContentsTableListCtrl->OnCancelTableLoad();
	EC_H(m_lpContentsTableListCtrl->RefreshTable());
}

void CContentsTableDlg::OnNotificationOn()
{
	if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;
	auto hRes = S_OK;
	EC_H(m_lpContentsTableListCtrl->NotificationOn());
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
	auto hRes = S_OK;
	WC_H(SearchEditor.DisplayDialog());
	if (S_OK == hRes)
	{
		const _SRestriction* lpRes = SearchEditor.GetRestriction();
		if (lpRes)
		{
			m_lpContentsTableListCtrl->SetRestriction(lpRes);

			if (SearchEditor.GetCheck(editor::CSearchEditor::SearchFields::FINDROW))
			{
				SetRestrictionType(mfcmapiFINDROW_RESTRICTION);
			}
			else
			{
				SetRestrictionType(mfcmapiNORMAL_RESTRICTION);
			}
		}
	}
}

void CContentsTableDlg::OnCreateRangeRestriction()
{
	auto hRes = S_OK;
	LPSRestriction lpRes = nullptr;

	if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;

	editor::CPropertyTagEditor MyPropertyTag(
		IDS_CREATERANGERES,
		NULL, // prompt
		PR_SUBJECT,
		m_bIsAB,
		m_lpContainer,
		this);

	WC_H(MyPropertyTag.DisplayDialog());
	if (S_OK == hRes)
	{
		editor::CEditor MyData(
			this,
			IDS_SEARCHCRITERIA,
			IDS_RANGESEARCHCRITERIAPROMPT,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyData.InitPane(0, TextPane::CreateSingleLinePane(IDS_SUBSTRING, false));
		MyData.InitPane(1, CheckPane::Create(IDS_APPLYUSINGFINDROW, false, false));

		WC_H(MyData.DisplayDialog());
		if (S_OK != hRes) return;

		const auto szString = MyData.GetStringW(0);
		// Allocate and create our SRestriction
		EC_H(CreateRangeRestriction(
			CHANGE_PROP_TYPE(MyPropertyTag.GetPropertyTag(), PT_UNICODE),
			szString,
			nullptr,
			&lpRes));
		if (S_OK != hRes)
		{
			MAPIFreeBuffer(lpRes);
			lpRes = nullptr;
		}

		m_lpContentsTableListCtrl->SetRestriction(lpRes);

		if (MyData.GetCheck(1))
		{
			SetRestrictionType(mfcmapiFINDROW_RESTRICTION);
		}
		else
		{
			SetRestrictionType(mfcmapiNORMAL_RESTRICTION);
		}
	}
}

void CContentsTableDlg::OnEditRestriction()
{
	auto hRes = S_OK;

	if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;

	editor::CRestrictEditor MyRestrict(
		this,
		nullptr, // No alloc parent - we must MAPIFreeBuffer the result
		m_lpContentsTableListCtrl->GetRestriction());

	WC_H(MyRestrict.DisplayDialog());
	if (S_OK != hRes) return;

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
		DebugPrintEx(DBGGeneric, CLASS, L"OnOutputTable", L"saving to %ws\n", szFileName.c_str());
		m_lpContentsTableListCtrl->OnOutputTable(szFileName);
	}
}

void CContentsTableDlg::OnSetColumns()
{
	if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;
	m_lpContentsTableListCtrl->DoSetColumns(
		false,
		true);
}

void CContentsTableDlg::OnGetStatus()
{
	if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;
	m_lpContentsTableListCtrl->GetStatus();
}

void CContentsTableDlg::OnSortTable()
{
	auto hRes = S_OK;

	if (!m_lpContentsTableListCtrl || !m_lpContentsTableListCtrl->IsContentsTableSet()) return;

	editor::CEditor MyData(
		this,
		IDS_SORTTABLE,
		IDS_SORTTABLEPROMPT1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	MyData.InitPane(0, TextPane::CreateSingleLinePane(IDS_CSORTS, false));
	MyData.InitPane(1, TextPane::CreateSingleLinePane(IDS_CCATS, false));
	MyData.InitPane(2, TextPane::CreateSingleLinePane(IDS_CEXPANDED, false));
	MyData.InitPane(3, CheckPane::Create(IDS_TBLASYNC, false, false));
	MyData.InitPane(4, CheckPane::Create(IDS_TBLBATCH, false, false));
	MyData.InitPane(5, CheckPane::Create(IDS_REFRESHAFTERSORT, true, false));

	WC_H(MyData.DisplayDialog());
	if (S_OK != hRes) return;

	const auto cSorts = MyData.GetDecimal(0);
	const auto cCategories = MyData.GetDecimal(1);
	const auto cExpanded = MyData.GetDecimal(2);

	if (cSorts < cCategories || cCategories < cExpanded)
	{
		error::ErrDialog(__FILE__, __LINE__, IDS_EDBADSORTS, cSorts, cCategories, cExpanded);
		return;
	}

	LPSSortOrderSet lpMySortOrders = nullptr;

	EC_H(MAPIAllocateBuffer(CbNewSSortOrderSet(cSorts), reinterpret_cast<LPVOID*>(&lpMySortOrders)));

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

			WC_H(MyPropertyTag.DisplayDialog());
			if (S_OK == hRes)
			{
				lpMySortOrders->aSort[i].ulPropTag = MyPropertyTag.GetPropertyTag();
				editor::CEditor MySortOrderDlg(
					this,
					IDS_SORTORDER,
					IDS_SORTORDERPROMPT,
					CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
				UINT uidDropDown[] = {
				IDS_DDTABLESORTASCEND,
				IDS_DDTABLESORTDESCEND,
				IDS_DDTABLESORTCOMBINE,
				IDS_DDTABLESORTCATEGMAX,
				IDS_DDTABLESORTCATEGMIN
				};
				MySortOrderDlg.InitPane(0, DropDownPane::Create(IDS_SORTORDER, _countof(uidDropDown), uidDropDown, true));

				WC_H(MySortOrderDlg.DisplayDialog());
				if (S_OK == hRes)
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
				else bNoError = false;
			}
			else bNoError = false;
		}

		if (bNoError)
		{
			EC_MAPI(m_lpContentsTableListCtrl->SetSortTable(
				lpMySortOrders,
				(MyData.GetCheck(3) ? TBL_ASYNC : 0) | (MyData.GetCheck(4) ? TBL_BATCH : 0) // flags
			));
		}
	}
	MAPIFreeBuffer(lpMySortOrders);

	if (MyData.GetCheck(5)) EC_H(m_lpContentsTableListCtrl->RefreshTable());
}

// Since the strategy for opening the selected property may vary depending on the table we're displaying,
// this virtual function allows us to override the default method with the method used by the table we've written a special class for.
_Check_return_ HRESULT CContentsTableDlg::OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp)
{
	auto hRes = S_OK;
	DebugPrintEx(DBGOpenItemProp, CLASS, L"OpenItemProp", L"iSelectedItem = 0x%X\n", iSelectedItem);

	if (!lppMAPIProp || !m_lpContentsTableListCtrl) return MAPI_E_INVALID_PARAMETER;

	if (-1 == iSelectedItem)
	{
		// Get the first selected item
		EC_H(m_lpContentsTableListCtrl->OpenNextSelectedItemProp(
			nullptr,
			bModify,
			lppMAPIProp));
	}
	else
	{
		EC_H(m_lpContentsTableListCtrl->DefaultOpenItemProp(
			iSelectedItem,
			bModify,
			lppMAPIProp));
	}

	return hRes;
}

_Check_return_ HRESULT CContentsTableDlg::OpenAttachmentsFromMessage(_In_ LPMESSAGE lpMessage)
{
	auto hRes = S_OK;

	if (lpMessage == nullptr) return MAPI_E_INVALID_PARAMETER;

	EC_H(DisplayTable(lpMessage, PR_MESSAGE_ATTACHMENTS, otDefault, this));

	return hRes;
}

_Check_return_ HRESULT CContentsTableDlg::OpenRecipientsFromMessage(_In_ LPMESSAGE lpMessage)
{
	auto hRes = S_OK;

	EC_H(DisplayTable(lpMessage, PR_MESSAGE_RECIPIENTS, otDefault, this));

	return hRes;
}

_Check_return_ bool CContentsTableDlg::HandleAddInMenu(WORD wMenuSelect)
{
	if (wMenuSelect < ID_ADDINMENU || ID_ADDINMENU + m_ulAddInMenuItems < wMenuSelect) return false;
	if (!m_lpContentsTableListCtrl) return false;
	LPMAPIPROP lpMAPIProp = nullptr;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	const auto lpAddInMenu = GetAddinMenuItem(m_hWnd, wMenuSelect);
	if (!lpAddInMenu) return false;

	const auto ulFlags = lpAddInMenu->ulFlags;

	const auto fRequestModify =
		(ulFlags & MENU_FLAGS_REQUESTMODIFY) ? mfcmapiREQUEST_MODIFY : mfcmapiDO_NOT_REQUEST_MODIFY;

	// Get the stuff we need for any case
	_AddInMenuParams MyAddInMenuParams = { nullptr };
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
		HandleAddInMenuSingle(
			&MyAddInMenuParams,
			nullptr,
			nullptr);
	}
	else
	{
		// Add appropriate flag to context
		MyAddInMenuParams.ulCurrentFlags |= ulFlags & (MENU_FLAGS_SINGLESELECT | MENU_FLAGS_MULTISELECT);
		auto items = m_lpContentsTableListCtrl->GetSelectedItemNums();
		for (const auto& item : items)
		{
			SRow MyRow = { 0 };

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
				if (FAILED(OpenItemProp(item, fRequestModify, &lpMAPIProp)))
				{
					lpMAPIProp = nullptr;
				}
			}

			HandleAddInMenuSingle(
				&MyAddInMenuParams,
				lpMAPIProp,
				nullptr);
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
			lpParams->lpFolder = dynamic_cast<LPMAPIFOLDER>(lpMAPIProp); // OpenItemProp returns LPMAPIFOLDER
			break;
		case MENU_CONTEXT_HIER_TABLE:
			lpParams->lpFolder = dynamic_cast<LPMAPIFOLDER>(lpMAPIProp); // OpenItemProp returns LPMAPIFOLDER
			break;
		}
	}

	InvokeAddInMenu(lpParams);
}

// WM_MFCMAPI_RESETCOLUMNS
// Returns true if we reset columns, false otherwise
_Check_return_ LRESULT CContentsTableDlg::msgOnResetColumns(WPARAM /*wParam*/, LPARAM /*lParam*/)
{
	DebugPrintEx(DBGGeneric, CLASS, L"msgOnResetColumns", L"Received message reset columns\n");

	if (m_lpContentsTableListCtrl)
	{
		m_lpContentsTableListCtrl->DoSetColumns(
			true,
			false);
		return true;
	}

	return false;
}