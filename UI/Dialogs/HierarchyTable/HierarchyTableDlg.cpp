#include <StdAfx.h>
#include <UI/Dialogs/HierarchyTable/HierarchyTableDlg.h>
#include <UI/Controls/HierarchyTableTreeCtrl.h>
#include <UI/FakeSplitter.h>
#include <UI/Controls/SingleMAPIPropListCtrl.h>
#include <MAPI/MapiObjects.h>
#include <UI/MFCUtilityFunctions.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Dialogs/Editors/RestrictEditor.h>

static std::wstring CLASS = L"CHierarchyTableDlg";

CHierarchyTableDlg::CHierarchyTableDlg(
	_In_ CParentWnd* pParentWnd,
	_In_ CMapiObjects* lpMapiObjects,
	UINT uidTitle,
	_In_opt_ LPUNKNOWN lpRootContainer,
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
		m_szTitle = strings::loadstring(IDS_TABLEASHIERARCHY);
	}

	m_nIDContextMenu = nIDContextMenu;

	m_ulDisplayFlags = dfNormal;
	m_lpHierarchyTableTreeCtrl = nullptr;
	m_lpContainer = nullptr;
	// need to make sure whatever gets passed to us is really a container
	if (lpRootContainer)
	{
		auto hRes = S_OK;
		LPMAPICONTAINER lpTemp = nullptr;
		EC_MAPI(lpRootContainer->QueryInterface(IID_IMAPIContainer, reinterpret_cast<LPVOID*>(&lpTemp)));
		if (lpTemp)
		{
			m_lpContainer = lpTemp;
		}
	}
}

CHierarchyTableDlg::~CHierarchyTableDlg()
{
	TRACE_DESTRUCTOR(CLASS);
	if (m_lpContainer) m_lpContainer->Release();
}

BEGIN_MESSAGE_MAP(CHierarchyTableDlg, CBaseDialog)
	ON_COMMAND(ID_DISPLAYSELECTEDITEM, OnDisplayItem)
	ON_COMMAND(ID_REFRESHVIEW, OnRefreshView)
	ON_COMMAND(ID_DISPLAYHIERARCHYTABLE, OnDisplayHierarchyTable)
	ON_COMMAND(ID_EDITSEARCHCRITERIA, OnEditSearchCriteria)
END_MESSAGE_MAP()

void CHierarchyTableDlg::OnInitMenu(_In_ CMenu* pMenu)
{
	if (pMenu)
	{
		if (m_lpHierarchyTableTreeCtrl)
		{
			const auto bItemSelected = m_lpHierarchyTableTreeCtrl->IsItemSelected();
			pMenu->EnableMenuItem(ID_DISPLAYSELECTEDITEM, DIM(bItemSelected));
			pMenu->EnableMenuItem(ID_DISPLAYHIERARCHYTABLE, DIM(bItemSelected));
			pMenu->EnableMenuItem(ID_EDITSEARCHCRITERIA, DIM(bItemSelected));
			for (ULONG ulMenu = ID_ADDINMENU; ulMenu < ID_ADDINMENU + m_ulAddInMenuItems; ulMenu++)
			{
				const auto lpAddInMenu = GetAddinMenuItem(m_hWnd, ulMenu);
				if (!lpAddInMenu) continue;

				const auto ulFlags = lpAddInMenu->ulFlags;
				UINT uiEnable = MF_ENABLED;

				if (ulFlags & MENU_FLAGS_MULTISELECT && !bItemSelected) uiEnable = MF_GRAYED;
				if (ulFlags & MENU_FLAGS_SINGLESELECT && !bItemSelected) uiEnable = MF_GRAYED;
				EnableAddInMenus(pMenu->m_hMenu, ulMenu, lpAddInMenu, uiEnable);
			}
		}
	}

	CBaseDialog::OnInitMenu(pMenu);
}

void CHierarchyTableDlg::OnCancel()
{
	ShowWindow(SW_HIDE);
	if (m_lpHierarchyTableTreeCtrl) m_lpHierarchyTableTreeCtrl->Release();
	m_lpHierarchyTableTreeCtrl = nullptr;
	CBaseDialog::OnCancel();
}

void CHierarchyTableDlg::OnDisplayItem()
{
	auto hRes = S_OK;

	auto lpMAPIContainer = m_lpHierarchyTableTreeCtrl->GetSelectedContainer(
		mfcmapiREQUEST_MODIFY);
	if (!lpMAPIContainer)
	{
		WARNHRESMSG(MAPI_E_NOT_FOUND, IDS_NOITEMSELECTED);
		return;
	}

	EC_H(DisplayObject(
		lpMAPIContainer,
		NULL,
		otContents,
		this));

	lpMAPIContainer->Release();
}

void CHierarchyTableDlg::OnDisplayHierarchyTable()
{
	auto hRes = S_OK;
	LPMAPITABLE lpMAPITable = nullptr;

	if (!m_lpHierarchyTableTreeCtrl) return;

	auto lpContainer = m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY);

	if (lpContainer)
	{
		editor::CEditor MyData(
			this,
			IDS_DISPLAYHIEARCHYTABLE,
			IDS_DISPLAYHIEARCHYTABLEPROMPT,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, viewpane::CheckPane::Create(IDS_CONVENIENTDEPTH, false, false));

		WC_H(MyData.DisplayDialog());

		EC_MAPI(lpContainer->GetHierarchyTable(
			MyData.GetCheck(0) ? CONVENIENT_DEPTH : 0
			| fMapiUnicode,
			&lpMAPITable));

		if (lpMAPITable)
		{
			EC_H(DisplayTable(
				lpMAPITable,
				otHierarchy,
				this));
			lpMAPITable->Release();
		}
		lpContainer->Release();
	}
}

void CHierarchyTableDlg::OnEditSearchCriteria()
{
	auto hRes = S_OK;

	if (!m_lpHierarchyTableTreeCtrl) return;

	// Find the highlighted item
	auto lpMAPIFolder = static_cast<LPMAPIFOLDER>(m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY));

	if (lpMAPIFolder)
	{
		DebugPrintEx(DBGGeneric, CLASS, L"OnEditSearchCriteria", L"Calling GetSearchCriteria on %p.\n", lpMAPIFolder);

		LPSRestriction lpRes = nullptr;
		LPENTRYLIST lpEntryList = nullptr;
		ULONG ulSearchState = 0;

		WC_MAPI(lpMAPIFolder->GetSearchCriteria(
			fMapiUnicode,
			&lpRes,
			&lpEntryList,
			&ulSearchState));
		if (MAPI_E_NOT_INITIALIZED == hRes)
		{
			DebugPrint(DBGGeneric, L"No search criteria has been set on this folder.\n");
			hRes = S_OK;
		}
		else CHECKHRESMSG(hRes, IDS_GETSEARCHCRITERIAFAILED);

		editor::CCriteriaEditor MyCriteria(
			this,
			lpRes,
			lpEntryList,
			ulSearchState);

		WC_H(MyCriteria.DisplayDialog());
		if (S_OK == hRes)
		{
			// make sure the user really wants to call SetSearchCriteria
			// hard to detect 'dirty' on this dialog so easier just to ask
			editor::CEditor MyYesNoDialog(
				this,
				IDS_CALLSETSEARCHCRITERIA,
				IDS_CALLSETSEARCHCRITERIAPROMPT,
				CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			WC_H(MyYesNoDialog.DisplayDialog());
			if (S_OK == hRes)
			{
				// do the set search criteria
				const auto lpNewRes = MyCriteria.DetachModifiedSRestriction();
				const auto lpNewEntryList = MyCriteria.DetachModifiedEntryList();
				const auto ulSearchFlags = MyCriteria.GetSearchFlags();
				EC_MAPI(lpMAPIFolder->SetSearchCriteria(
					lpNewRes,
					lpNewEntryList,
					ulSearchFlags));
				MAPIFreeBuffer(lpNewRes);
				MAPIFreeBuffer(lpNewEntryList);
			}
		}
		lpMAPIFolder->Release();
	}
}

BOOL CHierarchyTableDlg::OnInitDialog()
{
	const auto bRet = CBaseDialog::OnInitDialog();

	if (m_lpFakeSplitter)
	{
		m_lpHierarchyTableTreeCtrl = new CHierarchyTableTreeCtrl(
			m_lpFakeSplitter,
			m_lpMapiObjects,
			this,
			m_ulDisplayFlags,
			m_nIDContextMenu);

		if (m_lpHierarchyTableTreeCtrl)
		{
			m_lpFakeSplitter->SetPaneOne(m_lpHierarchyTableTreeCtrl);

			m_lpFakeSplitter->SetPercent(0.25);
		}
	}

	UpdateTitleBarText();

	return bRet;
}

void CHierarchyTableDlg::CreateDialogAndMenu(UINT nIDMenuResource)
{
	auto hRes = S_OK;

	DebugPrintEx(DBGCreateDialog, CLASS, L"CreateDialogAndMenu", L"id = 0x%X\n", nIDMenuResource);
	CBaseDialog::CreateDialogAndMenu(nIDMenuResource, IDR_MENU_HIERARCHY_TABLE, IDS_HIERARCHYTABLE);

	if (m_lpHierarchyTableTreeCtrl)
	{
		EC_H(m_lpHierarchyTableTreeCtrl->LoadHierarchyTable(m_lpContainer));
	}
}

// Per Q167960 BUG: ESC/ENTER Keys Do Not Work When Editing CTreeCtrl Labels
BOOL CHierarchyTableDlg::PreTranslateMessage(MSG* pMsg)
{
	// If edit control is visible in tree view control, when you send a
	// WM_KEYDOWN message to the edit control it will dismiss the edit
	// control. When the ENTER key was sent to the edit control, the
	// parent window of the tree view control is responsible for updating
	// the item's label in TVN_ENDLABELEDIT notification code.
	if (m_lpHierarchyTableTreeCtrl && pMsg &&
		pMsg->message == WM_KEYDOWN &&
		(pMsg->wParam == VK_RETURN || pMsg->wParam == VK_ESCAPE))
	{
		const auto edit = m_lpHierarchyTableTreeCtrl->GetEditControl();
		if (edit)
		{
			edit->SendMessage(WM_KEYDOWN, pMsg->wParam, pMsg->lParam);
			return true;
		}
	}
	return CMyDialog::PreTranslateMessage(pMsg);
}

void CHierarchyTableDlg::OnRefreshView()
{
	auto hRes = S_OK;

	DebugPrintEx(DBGGeneric, CLASS, L"OnRefreshView", L"\n");
	if (m_lpHierarchyTableTreeCtrl)
		EC_H(m_lpHierarchyTableTreeCtrl->RefreshHierarchyTable());
}

_Check_return_ bool CHierarchyTableDlg::HandleAddInMenu(WORD wMenuSelect)
{
	if (wMenuSelect < ID_ADDINMENU || ID_ADDINMENU + m_ulAddInMenuItems < wMenuSelect) return false;
	if (!m_lpHierarchyTableTreeCtrl) return false;

	LPMAPICONTAINER lpContainer = nullptr;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	const auto lpAddInMenu = GetAddinMenuItem(m_hWnd, wMenuSelect);
	if (!lpAddInMenu) return false;

	const auto ulFlags = lpAddInMenu->ulFlags;

	const auto fRequestModify =
		ulFlags & MENU_FLAGS_REQUESTMODIFY ? mfcmapiREQUEST_MODIFY : mfcmapiDO_NOT_REQUEST_MODIFY;

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
		SRow MyRow = { 0 };

		// If we have a row to give, give it - it's free
		const auto lpData = m_lpHierarchyTableTreeCtrl->GetSelectedItemData();
		if (lpData)
		{
			MyRow.cValues = lpData->cSourceProps;
			MyRow.lpProps = lpData->lpSourceProps;
			MyAddInMenuParams.lpRow = &MyRow;
			MyAddInMenuParams.ulCurrentFlags |= MENU_FLAGS_ROW;
		}

		if (!(ulFlags & MENU_FLAGS_ROW))
		{
			lpContainer = m_lpHierarchyTableTreeCtrl->GetSelectedContainer(fRequestModify);
		}

		HandleAddInMenuSingle(
			&MyAddInMenuParams,
			nullptr,
			lpContainer);
		if (lpContainer) lpContainer->Release();
	}
	return true;
}

void CHierarchyTableDlg::HandleAddInMenuSingle(
	_In_ LPADDINMENUPARAMS lpParams,
	_In_opt_ LPMAPIPROP /*lpMAPIProp*/,
	_In_opt_ LPMAPICONTAINER /*lpContainer*/)
{
	InvokeAddInMenu(lpParams);
}