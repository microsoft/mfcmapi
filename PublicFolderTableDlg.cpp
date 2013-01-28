// PublicFolderTableDlg.cpp : implementation file
//

#include "stdafx.h"
#include "PublicFolderTableDlg.h"
#include "ContentsTableListCtrl.h"
#include "MapiObjects.h"
#include "MAPIFunctions.h"
#include "MAPIStoreFunctions.h"
#include "SingleMAPIPropListCtrl.h"
#include "ColumnTags.h"
#include "UIFunctions.h"
#include "Editor.h"
#include "PropertyTagEditor.h"
#include "InterpretProp2.h"

static TCHAR* CLASS = _T("CPublicFolderTableDlg");

/////////////////////////////////////////////////////////////////////////////
// CPublicFolderTableDlg dialog

CPublicFolderTableDlg::CPublicFolderTableDlg(
	_In_ CParentWnd* pParentWnd,
	_In_ CMapiObjects* lpMapiObjects,
	_In_ LPCTSTR lpszServerName,
	_In_ LPMAPITABLE lpMAPITable
	):
CContentsTableDlg(
				  pParentWnd,
				  lpMapiObjects,
				  IDS_PUBLICFOLDERTABLE,
				  mfcmapiDO_NOT_CALL_CREATE_DIALOG,
				  lpMAPITable,
				  (LPSPropTagArray) &sptPFCols,
				  NUMPFCOLUMNS,
				  PFColumns,
				  NULL,
				  MENU_CONTEXT_PUBLIC_FOLDER_TABLE)
{
	TRACE_CONSTRUCTOR(CLASS);
	HRESULT hRes = S_OK;

	EC_H(CopyString(&m_lpszServerName,lpszServerName,NULL));

	CreateDialogAndMenu(NULL);
} // CPublicFolderTableDlg::CPublicFolderTableDlg

CPublicFolderTableDlg::~CPublicFolderTableDlg()
{
	TRACE_DESTRUCTOR(CLASS);
	MAPIFreeBuffer(m_lpszServerName);
} // CPublicFolderTableDlg::~CPublicFolderTableDlg

void CPublicFolderTableDlg::CreateDialogAndMenu(UINT nIDMenuResource)
{
	DebugPrintEx(DBGCreateDialog,CLASS,_T("CreateDialogAndMenu"),_T("id = 0x%X\n"),nIDMenuResource);
	CContentsTableDlg::CreateDialogAndMenu(nIDMenuResource);

	UpdateMenuString(
		m_hWnd,
		ID_CREATEPROPERTYSTRINGRESTRICTION,
		IDS_PFRESMENU);
} // CPublicFolderTableDlg::CreateDialogAndMenu

void CPublicFolderTableDlg::OnDisplayItem()
{
/*	HRESULT			hRes = S_OK;
	LPMAPIPROP		lpMAPIProp = NULL;
	LPMAPISESSION	lpMAPISession	= NULL;
	LPMDB			lpMDB = NULL;
	TCHAR*			szMailboxDN = NULL;
	int				iItem = -1;
	SortListData*	lpListData = NULL;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB)
	{
		EC_H(OpenMessageStoreGUID(lpMAPISession, pbExchangeProviderPrimaryUserGuid, &lpMDB));
	}
	if (lpMDB)
	{
		do
		{
			hRes = S_OK;
			lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
			if (lpListData)
			{
				szMailboxDN = lpListData->data.Contents.szDN;

				if (szMailboxDN)
				{
					EC_H(OpenOtherUsersMailbox(
						lpMAPISession,
						lpMDB,
						m_lpszServerName,
						szMailboxDN,
						OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_TAKE_OWNERSHIP,
						(LPMDB*) &lpMAPIProp));

					if (lpMAPIProp)
					{
						if (m_lpPropDisplay)
							EC_H(m_lpPropDisplay->SetDataSource(lpMAPIProp, NULL, false));

						EC_H(DisplayObject(
							lpMAPIProp,
							NULL,
							otStore,
							this));
						lpMAPIProp->Release();
						lpMAPIProp = NULL;
					}
				}
			}
		}
		while (iItem != -1);
	}*/
} // CPublicFolderTableDlg::OnDisplayItem

_Check_return_ HRESULT CPublicFolderTableDlg::OpenItemProp(int /*iSelectedItem*/, __mfcmapiModifyEnum /*bModify*/, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp)
{
	if (lppMAPIProp) *lppMAPIProp = NULL;
	HRESULT			hRes = S_OK;
/*	LPSBinary		lpProviderUID = NULL;
	SortListData*	lpListData = NULL;
	LPMAPISESSION	lpMAPISession	= NULL;
	LPMDB			lpMDB = NULL;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	DebugPrintEx(DBGOpenItemProp,CLASS,_T("OpenItemProp"),_T("iSelectedItem = 0x%X\n"),iSelectedItem);

	if (!lppMAPIProp || !m_lpContentsTableListCtrl) return MAPI_E_INVALID_PARAMETER;

	lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return MAPI_E_INVALID_PARAMETER;

	lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	if (!lpMDB)
	{
		EC_H(OpenMessageStoreGUID(lpMAPISession, pbExchangeProviderPrimaryUserGuid, &lpMDB));
	}

	lpListData = (SortListData*) m_lpContentsTableListCtrl->GetItemData(iSelectedItem);
	if (lpListData)
	{
		TCHAR* szFolderDN = NULL;
		szFolderDN = lpListData->data.Contents.szDN;

		if (szFolderDN)
		{
			EC_H(OpenPublicMessageStoreFolderAdminPriv(
				lpMAPISession,
				lpMDB,
				m_lpszServerName,
				szFolderDN,
				(LPMDB*) lppMAPIProp));
		}
	}*/
	return hRes;
} // CPublicFolderTableDlg::OpenItemProp

void CPublicFolderTableDlg::OnCreatePropertyStringRestriction()
{
	HRESULT			hRes = S_OK;
	LPSRestriction	lpRes = NULL;

	CPropertyTagEditor MyPropertyTag(
		IDS_PROPRES, // title
		NULL, // prompt
		PR_DISPLAY_NAME,
		m_bIsAB,
		m_lpContainer,
		this);

	WC_H(MyPropertyTag.DisplayDialog());
	if (S_OK == hRes)
	{
		CEditor MyData(
			this,
			IDS_SEARCHCRITERIA,
			IDS_PFSEARCHCRITERIAPROMPT,
			2,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
		MyData.SetPromptPostFix(AllFlagsToString(flagFuzzyLevel,true));

		MyData.InitSingleLine(0,IDS_NAME,NULL,false);
		MyData.InitSingleLine(1,IDS_ULFUZZYLEVEL,NULL,false);
		MyData.SetHex(1,FL_IGNORECASE|FL_PREFIX);

		WC_H(MyData.DisplayDialog());
		if (S_OK != hRes) return;

		// Allocate and create our SRestriction
		EC_H(CreatePropertyStringRestriction(
			CHANGE_PROP_TYPE(MyPropertyTag.GetPropertyTag(),PT_TSTRING),
			MyData.GetString(0),
			MyData.GetHex(1),
			NULL,
			&lpRes));
		if (S_OK != hRes && lpRes)
		{
			MAPIFreeBuffer(lpRes);
			lpRes = NULL;
		}

		m_lpContentsTableListCtrl->SetRestriction(lpRes);

		SetRestrictionType(mfcmapiNORMAL_RESTRICTION);
	}
} // CPublicFolderTableDlg::OnCreatePropertyStringRestriction
