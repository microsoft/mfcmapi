// HierarchyTableDlg.cpp : implementation file
//

#include "stdafx.h"
#include "Error.h"

#include "HierarchyTableDlg.h"

#include "HierarchyTableTreeCtrl.h"
#include "FakeSplitter.h"
#include "SingleMAPIPropListCtrl.h"
#include "MAPIObjects.h"
#include "MFCUtilityFunctions.h"
#include "Editor.h"
#include "InterpretProp.h"
#include "RestrictEditor.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

static TCHAR* CLASS = _T("CHierarchyTableDlg");

/////////////////////////////////////////////////////////////////////////////
// CHierarchyTableDlg dialog


CHierarchyTableDlg::CHierarchyTableDlg(
													 CParentWnd* pParentWnd,
													 CMapiObjects *lpMapiObjects,
													 UINT uidTitle,
													 LPUNKNOWN lpRootContainer,
													 ULONG nIDContextMenu,
													 ULONG ulAddInContext
													 ):
CBaseDialog(
		pParentWnd,
		lpMapiObjects,
		ulAddInContext)
{
	TRACE_CONSTRUCTOR(CLASS);
	if (NULL != uidTitle)
	{
		m_szTitle.LoadString(uidTitle);
	}
	else
	{
		m_szTitle.LoadString(IDS_TABLEASHIERARCHY);
	}

	m_nIDContextMenu = nIDContextMenu;

	m_bShowingDeletedFolders = mfcmapiDO_NOT_SHOW_DELETED_ITEMS;
	m_lpHierarchyTableTreeCtrl = NULL;
	m_lpContainer = NULL;
	//need to make sure whatever gets passed to us is really a container
	if (lpRootContainer)
	{
		HRESULT hRes = S_OK;
		LPMAPICONTAINER lpTemp = NULL;
		EC_H(lpRootContainer->QueryInterface(IID_IMAPIContainer,(LPVOID*) &lpTemp));
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
//{{AFX_MSG_MAP(CHierarchyTableDlg)
	ON_COMMAND(ID_DISPLAYSELECTEDITEM, OnDisplayItem)
	ON_COMMAND(ID_REFRESHVIEW, OnRefreshView)
	ON_COMMAND(ID_DISPLAYHIERARCHYTABLE,OnDisplayHierarchyTable)
	ON_COMMAND(ID_EDITSEARCHCRITERIA, OnEditSearchCriteria)
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

void CHierarchyTableDlg::OnInitMenu(CMenu* pMenu)
{
	if (pMenu)
	{
		BOOL	bItemSelected = m_lpHierarchyTableTreeCtrl && m_lpHierarchyTableTreeCtrl->m_bItemSelected;
		if (m_lpHierarchyTableTreeCtrl)
		{
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
				EnableAddInMenus(pMenu, ulMenu, lpAddInMenu, uiEnable);
			}
		}

	}
	CBaseDialog::OnInitMenu(pMenu);
}

void CHierarchyTableDlg::OnCancel()
{
	ShowWindow(SW_HIDE);
	if (m_lpHierarchyTableTreeCtrl) m_lpHierarchyTableTreeCtrl->Release();
	m_lpHierarchyTableTreeCtrl = NULL;
	CBaseDialog::OnCancel();
}

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
		this))

	lpMAPIContainer->Release();
	return;
}

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
		MyData.InitCheck(0,IDS_CONVENIENTDEPTH,false,false);

		WC_H(MyData.DisplayDialog());

		EC_H(lpContainer->GetHierarchyTable(
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
	return;
}//CHierarchyTableDlg::OnDisplayHierarchyTable

void CHierarchyTableDlg::OnEditSearchCriteria()
{
	HRESULT			hRes = S_OK;

	if (!m_lpHierarchyTableTreeCtrl) return;

	//Find the highlighted item
	LPMAPIFOLDER lpMAPIFolder = (LPMAPIFOLDER) m_lpHierarchyTableTreeCtrl->GetSelectedContainer(mfcmapiREQUEST_MODIFY);

	if (lpMAPIFolder)
	{
		DebugPrintEx(DBGGeneric,CLASS,_T("OnEditSearchCriteria"),_T("Calling GetSearchCriteria on 0x%08X.\n"),lpMAPIFolder);

		LPSRestriction lpRes = NULL;
		LPENTRYLIST lpEntryList = NULL;
		ULONG ulSearchState = 0;

		WC_H(lpMAPIFolder->GetSearchCriteria(
			fMapiUnicode,
			&lpRes,
			&lpEntryList,
			&ulSearchState));
		if (MAPI_E_NOT_INITIALIZED == hRes)
		{
			DebugPrint(DBGGeneric,_T("No search criteria has been set on this folder.\n"));
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
			//make sure the user really wants to call SetSearchCriteria
			//hard to detect 'dirty' on this dialog so easier just to ask
			CEditor MyYesNoDialog(
				this,
				IDS_CALLSETSEARCHCRITERIA,
				IDS_CALLSETSEARCHCRITERIAPROMPT,
				0,
				CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
			WC_H(MyYesNoDialog.DisplayDialog());
			if (S_OK == hRes)
			{
				//do the set search criteria
				LPSRestriction lpNewRes = MyCriteria.DetachModifiedSRestriction();
				LPENTRYLIST lpNewEntryList = MyCriteria.DetachModifiedEntryList();
				ULONG ulSearchFlags = MyCriteria.GetSearchFlags();
				EC_H(lpMAPIFolder->SetSearchCriteria(
					lpNewRes,
					lpNewEntryList,
					ulSearchFlags));
				MAPIFreeBuffer(lpNewRes);
				MAPIFreeBuffer(lpNewEntryList);
			}
		}
		lpMAPIFolder->Release();
	}
	return;
}//CHierarchyTableDlg::OnEditSearchCriteria

BOOL CHierarchyTableDlg::OnInitDialog()
{
	CBaseDialog::OnInitDialog();

	if (m_lpFakeSplitter)
	{
		m_lpHierarchyTableTreeCtrl = new CHierarchyTableTreeCtrl(
			m_lpFakeSplitter,
			m_lpMapiObjects,
			this,
			m_bShowingDeletedFolders);

		if (m_lpHierarchyTableTreeCtrl)
		{
			m_lpFakeSplitter->SetPaneOne(m_lpHierarchyTableTreeCtrl);

			m_lpFakeSplitter->SetPercent(0.25);
		}
	}

	UpdateTitleBarText(NULL);

	return TRUE;
}//CHierarchyTableDlg::OnInitDialog

BOOL CHierarchyTableDlg::CreateDialogAndMenu(UINT nIDMenuResource)
{
	HRESULT hRes = S_OK;

	DebugPrintEx(DBGCreateDialog,CLASS,_T("CreateDialogAndMenu"),_T("id = 0x%X\n"),nIDMenuResource);
	EC_B(CBaseDialog::CreateDialogAndMenu(nIDMenuResource));

	EC_B(AddMenu(IDR_MENU_HIERARCHY_TABLE,IDS_HIERARCHYTABLE,(UINT)-1));

	if (m_lpHierarchyTableTreeCtrl)
	{
		EC_H(m_lpHierarchyTableTreeCtrl->LoadHierarchyTable(m_lpContainer));
	}

	return HRES_TO_BOOL(hRes);
}//CHierarchyTableDlg::CreateDialogAndMenu

//Per Q167960 BUG: ESC/ENTER Keys Do Not Work When Editing CTreeCtrl Labels
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
			return TRUE;
		}
	}
	return CDialog::PreTranslateMessage(pMsg);
}//CHierarchyTableDlg::PreTranslateMessage(MSG* pMsg)

void CHierarchyTableDlg::OnRefreshView()
{
	HRESULT hRes = S_OK;

	DebugPrintEx(DBGGeneric,CLASS,_T("OnRefreshView"),_T("\n"));
	if (m_lpHierarchyTableTreeCtrl)
		EC_H(m_lpHierarchyTableTreeCtrl->RefreshHierarchyTable());
}

BOOL CHierarchyTableDlg::HandleAddInMenu(WORD wMenuSelect)
{
	if (wMenuSelect < ID_ADDINMENU || ID_ADDINMENU+m_ulAddInMenuItems < wMenuSelect) return false;
	if (!m_lpHierarchyTableTreeCtrl) return false;

	LPMAPICONTAINER	lpContainer = NULL;
	CWaitCursor	Wait;//Change the mouse to an hourglass while we work.

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
		MyAddInMenuParams.lpMAPISession = m_lpMapiObjects->GetSession();//do not release
		MyAddInMenuParams.lpMDB = m_lpMapiObjects->GetMDB();//do not release
		MyAddInMenuParams.lpAdrBook = m_lpMapiObjects->GetAddrBook(false);//do not release
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
}//CHierarchyTableDlg::HandleAddInMenu

void CHierarchyTableDlg::HandleAddInMenuSingle(
									   LPADDINMENUPARAMS lpParams,
									   LPMAPIPROP /*lpMAPIProp*/,
									   LPMAPICONTAINER /*lpContainer*/)
{
	InvokeAddInMenu(lpParams);
} // CHierarchyTableDlg::HandleAddInMenuSingle
