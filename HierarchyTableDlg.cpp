// HierarchyTableDlg.cpp : implementation file
//

#include "stdafx.h"
#include "HierarchyTableDlg.h"
#include "HierarchyTableTreeCtrl.h"
#include "FakeSplitter.h"
#include "SingleMAPIPropListCtrl.h"
#include "MAPIObjects.h"
#include "MFCUtilityFunctions.h"
#include "Editor.h"
#include "InterpretProp.h"
#include "RestrictEditor.h"

static TCHAR* CLASS = _T("CHierarchyTableDlg");

/////////////////////////////////////////////////////////////////////////////
// CHierarchyTableDlg dialog


CHierarchyTableDlg::CHierarchyTableDlg(
									   _In_ CParentWnd* pParentWnd,
									   _In_ CMapiObjects* lpMapiObjects,
									   UINT uidTitle,
									   _In_opt_ LPUNKNOWN lpRootContainer,
									   ULONG nIDContextMenu,
									   ULONG ulAddInContext
									   ):
CBaseDialog(
			pParentWnd,
			lpMapiObjects,
			ulAddInContext)
{
	TRACE_CONSTRUCTOR(CLASS);
	HRESULT hRes = S_OK;
	if (NULL != uidTitle)
	{
		EC_B(m_szTitle.LoadString(uidTitle));
	}
	else
	{
		EC_B(m_szTitle.LoadString(IDS_TABLEASHIERARCHY));
	}

	m_nIDContextMenu = nIDContextMenu;

	m_ulDisplayFlags = dfNormal;
	m_lpHierarchyTableTreeCtrl = NULL;
	m_lpContainer = NULL;
	// need to make sure whatever gets passed to us is really a container
	if (lpRootContainer)
	{
		LPMAPICONTAINER lpTemp = NULL;
		EC_MAPI(lpRootContainer->QueryInterface(IID_IMAPIContainer,(LPVOID*) &lpTemp));
		if (lpTemp)
		{
			m_lpContainer = lpTemp;
		}
	}
} // CHierarchyTableDlg::CHierarchyTableDlg

CHierarchyTableDlg::~CHierarchyTableDlg()
{
	TRACE_DESTRUCTOR(CLASS);
	if (m_lpContainer) m_lpContainer->Release();
} // CHierarchyTableDlg::~CHierarchyTableDlg

BEGIN_MESSAGE_MAP(CHierarchyTableDlg, CBaseDialog)
	ON_COMMAND(ID_DISPLAYSELECTEDITEM, OnDisplayItem)
	ON_COMMAND(ID_REFRESHVIEW, OnRefreshView)
	ON_COMMAND(ID_DISPLAYHIERARCHYTABLE,OnDisplayHierarchyTable)
	ON_COMMAND(ID_EDITSEARCHCRITERIA, OnEditSearchCriteria)
END_MESSAGE_MAP()

void CHierarchyTableDlg::OnInitMenu(_In_ CMenu* pMenu)
{
	if (pMenu)
	{
		if (m_lpHierarchyTableTreeCtrl)
		{
			bool bItemSelected = m_lpHierarchyTableTreeCtrl->IsItemSelected();
			pMenu->EnableMenuItem(ID_DISPLAYSELECTEDITEM,DIM(bItemSelected));
			pMenu->EnableMenuItem(ID_DISPLAYHIERARCHYTABLE,DIM(bItemSelected));
			pMenu->EnableMenuItem(ID_EDITSEARCHCRITERIA,DIM(bItemSelected));
			ULONG ulMenu = ID_ADDINMENU;
			for (ulMenu = ID_ADDINMENU; ulMenu < ID_ADDINMENU+m_ulAddInMenuItems ; ulMenu++)
			{
				LPMENUITEM lpAddInMenu = GetAddinMenuItem(m_hWnd,ulMenu);
				if (!lpAddInMenu) continue;

				ULONG ulFlags = lpAddInMenu->ulFlags;
				UINT uiEnable = MF_ENABLED;

				if ((ulFlags & MENU_FLAGS_MULTISELECT) && !bItemSelected) uiEnable = MF_GRAYED;
				if ((ulFlags & MENU_FLAGS_SINGLESELECT) && !bItemSelected) uiEnable = MF_GRAYED;
				EnableAddInMenus(pMenu->m_hMenu, ulMenu, lpAddInMenu, uiEnable);
			}
		}

	}
	CBaseDialog::OnInitMenu(pMenu);
} // CHierarchyTableDlg::OnInitMenu

void CHierarchyTableDlg::OnCancel()
{
	ShowWindow(SW_HIDE);
	if (m_lpHierarchyTableTreeCtrl) m_lpHierarchyTableTreeCtrl->Release();
	m_lpHierarchyTableTreeCtrl = NULL;
	CBaseDialog::OnCancel();
} // CHierarchyTableDlg::OnCancel

void CHierarchyTableDlg::OnDisplayItem()
{
	HRESULT			hRes = S_OK;

	LPMAPICONTAINER	lpMAPIContainer = m_lpHierarchyTableTreeCtrl->GetSelectedContainer(
		mfcmapiREQUEST_MODIFY);
	if (!lpMAPIContainer)
	{
		WARNHRESMSG(MAPI_E_NOT_FOUND,IDS_NOITEMSELECTED);
		return;
	}

	EC_H(DisplayObject(
		lpMAPIContainer,
		NULL,
		otContents,
		this));

	lpMAPIContainer->Release();
} // CHierarchyTableDlg::OnDisplayItem

void CHierarchyTableDlg::OnDisplayHierarchyTable()
{
	HRESULT			hRes = S_OK;
	LPMAPITABLE		lpMAPITable = NULL;

	if (!m_lpHierarchyTableTreeCtrl) return;

	LPMAPICONTAINER	lpContainer = m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY);

	if (lpContainer)
	{
		CEditor MyData(
			this,
			IDS_DISPLAYHIEARCHYTABLE,
			IDS_DISPLAYHIEARCHYTABLEPROMPT,
			1,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, CreateCheckPane(IDS_CONVENIENTDEPTH, false, false));

		WC_H(MyData.DisplayDialog());

		EC_MAPI(lpContainer->GetHierarchyTable(
			MyData.GetCheck(0)?CONVENIENT_DEPTH:0
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
} // CHierarchyTableDlg::OnDisplayHierarchyTable

void CHierarchyTableDlg::OnEditSearchCriteria()
{
	HRESULT			hRes = S_OK;

	if (!m_lpHierarchyTableTreeCtrl) return;

	// Find the highlighted item
	LPMAPIFOLDER lpMAPIFolder = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY);

	if (lpMAPIFolder)
	{
		DebugPrintEx(DBGGeneric,CLASS,_T("OnEditSearchCriteria"),_T("Calling GetSearchCriteria on %p.\n"),lpMAPIFolder);

		LPSRestriction lpRes = NULL;
		LPENTRYLIST lpEntryList = NULL;
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
		else CHECKHRESMSG(hRes,IDS_GETSEARCHCRITERIAFAILED);

		CCriteriaEditor MyCriteria(
			this,
			lpRes,
			lpEntryList,
			ulSearchState);

		WC_H(MyCriteria.DisplayDialog());
		if (S_OK == hRes)
		{
			// make sure the user really wants to call SetSearchCriteria
			// hard to detect 'dirty' on this dialog so easier just to ask
			CEditor MyYesNoDialog(
				this,
				IDS_CALLSETSEARCHCRITERIA,
				IDS_CALLSETSEARCHCRITERIAPROMPT,
				0,
				CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
			WC_H(MyYesNoDialog.DisplayDialog());
			if (S_OK == hRes)
			{
				// do the set search criteria
				LPSRestriction lpNewRes = MyCriteria.DetachModifiedSRestriction();
				LPENTRYLIST lpNewEntryList = MyCriteria.DetachModifiedEntryList();
				ULONG ulSearchFlags = MyCriteria.GetSearchFlags();
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
} // CHierarchyTableDlg::OnEditSearchCriteria

BOOL CHierarchyTableDlg::OnInitDialog()
{
	BOOL bRet = CBaseDialog::OnInitDialog();

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

	UpdateTitleBarText(NULL);

	return bRet;
} // CHierarchyTableDlg::OnInitDialog

void CHierarchyTableDlg::CreateDialogAndMenu(UINT nIDMenuResource)
{
	HRESULT hRes = S_OK;

	DebugPrintEx(DBGCreateDialog,CLASS,_T("CreateDialogAndMenu"),_T("id = 0x%X\n"),nIDMenuResource);
	CBaseDialog::CreateDialogAndMenu(nIDMenuResource,IDR_MENU_HIERARCHY_TABLE,IDS_HIERARCHYTABLE);

	if (m_lpHierarchyTableTreeCtrl)
	{
		EC_H(m_lpHierarchyTableTreeCtrl->LoadHierarchyTable(m_lpContainer));
	}
} // CHierarchyTableDlg::CreateDialogAndMenu

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
		CEdit* edit = m_lpHierarchyTableTreeCtrl->GetEditControl();
		if (edit)
		{
			edit->SendMessage(WM_KEYDOWN, pMsg->wParam, pMsg->lParam);
			return true;
		}
	}
	return CMyDialog::PreTranslateMessage(pMsg);
} // CHierarchyTableDlg::PreTranslateMessage(MSG* pMsg)

void CHierarchyTableDlg::OnRefreshView()
{
	HRESULT hRes = S_OK;

	DebugPrintEx(DBGGeneric,CLASS,_T("OnRefreshView"),_T("\n"));
	if (m_lpHierarchyTableTreeCtrl)
		EC_H(m_lpHierarchyTableTreeCtrl->RefreshHierarchyTable());
} // CHierarchyTableDlg::OnRefreshView

_Check_return_ bool CHierarchyTableDlg::HandleAddInMenu(WORD wMenuSelect)
{
	if (wMenuSelect < ID_ADDINMENU || ID_ADDINMENU+m_ulAddInMenuItems < wMenuSelect) return false;
	if (!m_lpHierarchyTableTreeCtrl) return false;

	LPMAPICONTAINER	lpContainer = NULL;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	LPMENUITEM lpAddInMenu = GetAddinMenuItem(m_hWnd,wMenuSelect);
	if (!lpAddInMenu) return false;

	ULONG ulFlags = lpAddInMenu->ulFlags;

	__mfcmapiModifyEnum	fRequestModify =
		(ulFlags & MENU_FLAGS_REQUESTMODIFY)?mfcmapiREQUEST_MODIFY:mfcmapiDO_NOT_REQUEST_MODIFY;

	// Get the stuff we need for any case
	_AddInMenuParams MyAddInMenuParams = {0};
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
	if (!(ulFlags & (MENU_FLAGS_SINGLESELECT|MENU_FLAGS_MULTISELECT)))
	{
		HandleAddInMenuSingle(
			&MyAddInMenuParams,
			NULL,
			NULL);
	}
	else
	{
		SRow MyRow = {0};

		// If we have a row to give, give it - it's free
		SortListData* lpData = (SortListData*) m_lpHierarchyTableTreeCtrl->GetSelectedItemData();
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
			NULL,
			lpContainer);
		if (lpContainer) lpContainer->Release();
	}
	return true;
} // CHierarchyTableDlg::HandleAddInMenu

void CHierarchyTableDlg::HandleAddInMenuSingle(
	_In_ LPADDINMENUPARAMS lpParams,
	_In_opt_ LPMAPIPROP /*lpMAPIProp*/,
	_In_opt_ LPMAPICONTAINER /*lpContainer*/)
{
	InvokeAddInMenu(lpParams);
} // CHierarchyTableDlg::HandleAddInMenuSingle