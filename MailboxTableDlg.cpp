// MailboxTableDlg.cpp : implementation file
//

#include "stdafx.h"
#include "MailboxTableDlg.h"
#include "ContentsTableListCtrl.h"
#include "MapiObjects.h"
#include "MAPIFunctions.h"
#include "MAPIStoreFunctions.h"
#include "SingleMAPIPropListCtrl.h"
#include "ColumnTags.h"
#include "MFCUtilityFunctions.h"
#include "Editor.h"
#include "PropertyTagEditor.h"
#include "InterpretProp2.h"
#include "PropTagArray.h"

static TCHAR* CLASS = _T("CMailboxTableDlg");

/////////////////////////////////////////////////////////////////////////////
// CMailboxTableDlg dialog

CMailboxTableDlg::CMailboxTableDlg(
								   _In_ CParentWnd* pParentWnd,
								   _In_ CMapiObjects* lpMapiObjects,
								   _In_z_ LPCTSTR lpszServerName,
								   _In_ LPMAPITABLE lpMAPITable
								   ):
CContentsTableDlg(
				  pParentWnd,
				  lpMapiObjects,
				  IDS_MAILBOXTABLE,
				  mfcmapiDO_NOT_CALL_CREATE_DIALOG,
				  lpMAPITable,
				  (LPSPropTagArray) &sptMBXCols,
				  NUMMBXCOLUMNS,
				  MBXColumns,
				  IDR_MENU_MAILBOX_TABLE_POPUP,
				  MENU_CONTEXT_MAILBOX_TABLE
				  )
{
	TRACE_CONSTRUCTOR(CLASS);
	HRESULT hRes = S_OK;

	EC_H(CopyString(&m_lpszServerName,lpszServerName,NULL));

	CreateDialogAndMenu(IDR_MENU_MAILBOX_TABLE);
} // CMailboxTableDlg::CMailboxTableDlg

CMailboxTableDlg::~CMailboxTableDlg()
{
	TRACE_DESTRUCTOR(CLASS);
	MAPIFreeBuffer(m_lpszServerName);
} // CMailboxTableDlg::~CMailboxTableDlg

BEGIN_MESSAGE_MAP(CMailboxTableDlg, CContentsTableDlg)
	ON_COMMAND(ID_OPENWITHFLAGS, OnOpenWithFlags)
END_MESSAGE_MAP()

void CMailboxTableDlg::OnInitMenu(_In_ CMenu* pMenu)
{
	if (pMenu)
	{
		if (m_lpContentsTableListCtrl)
		{
			int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
			pMenu->EnableMenuItem(ID_OPENWITHFLAGS,DIMMSOK(iNumSel));
		}
	}
	CContentsTableDlg::OnInitMenu(pMenu);
} // CMailboxTableDlg::OnInitMenu

void CMailboxTableDlg::CreateDialogAndMenu(UINT nIDMenuResource)
{
	DebugPrintEx(DBGCreateDialog,CLASS,_T("CreateDialogAndMenu"),_T("id = 0x%X\n"),nIDMenuResource);
	CContentsTableDlg::CreateDialogAndMenu(nIDMenuResource);

	UpdateMenuString(
		this,
		ID_CREATEPROPERTYSTRINGRESTRICTION,
		IDS_MBRESMENU);
} // CMailboxTableDlg::CreateDialogAndMenu

void CMailboxTableDlg::DisplayItem(ULONG ulFlags)
{
	HRESULT		hRes = S_OK;
	CWaitCursor	Wait; // Change the mouse to an hourglass while we work.

	LPMAPISESSION	lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	LPMDB lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	LPMDB lpGUIDMDB = NULL;
	if (!lpMDB)
	{
		EC_H(OpenMessageStoreGUID(lpMAPISession, pbExchangeProviderPrimaryUserGuid, &lpGUIDMDB));
	}
	LPMDB lpSourceMDB = NULL; // do not release

	lpSourceMDB = lpMDB?lpMDB:lpGUIDMDB;

	if (lpSourceMDB)
	{
		LPMDB			lpNewMDB = NULL;
		TCHAR*			szMailboxDN = NULL;
		int				iItem = -1;
		SortListData*	lpListData = NULL;
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
						lpSourceMDB,
						m_lpszServerName,
						szMailboxDN,
						ulFlags,
						&lpNewMDB));

					if (lpNewMDB)
					{
						EC_H(DisplayObject(
							(LPMAPIPROP) lpNewMDB,
							NULL,
							otStore,
							this));
						lpNewMDB->Release();
						lpNewMDB = NULL;
					}
				}
			}
		}
		while (iItem != -1);

	}
	if (lpGUIDMDB) lpGUIDMDB->Release();
} // CMailboxTableDlg::DisplayItem

void CMailboxTableDlg::OnDisplayItem()
{
	DisplayItem(OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_TAKE_OWNERSHIP);
} // CMailboxTableDlg::OnDisplayItem

void CMailboxTableDlg::OnOpenWithFlags()
{
	HRESULT		hRes = S_OK;

	CEditor MyPrompt(
		this,
		IDS_OPENWITHFLAGS,
		IDS_OPENWITHFLAGSPROMPT,
		1,
		CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
	MyPrompt.SetPromptPostFix(AllFlagsToString(PROP_ID(PR_PROFILE_OPEN_FLAGS),true));
	MyPrompt.InitSingleLine(0,IDS_CREATESTORENTRYIDFLAGS,NULL,false);
	MyPrompt.SetHex(0,OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_TAKE_OWNERSHIP);
	WC_H(MyPrompt.DisplayDialog());
	if (S_OK == hRes)
	{
		DisplayItem(MyPrompt.GetHex(0));
	}
} // CMailboxTableDlg::OnOpenWithFlags

void CMailboxTableDlg::OnCreatePropertyStringRestriction()
{
	HRESULT			hRes = S_OK;
	LPSRestriction	lpRes = NULL;

	CPropertyTagEditor MyPropertyTag(
		IDS_PROPRES,
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
			IDS_MBSEARCHCRITERIAPROMPT,
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
		if (hRes != S_OK)
		{
			MAPIFreeBuffer(lpRes);
			lpRes = NULL;
		}

		m_lpContentsTableListCtrl->SetRestriction(lpRes);

		SetRestrictionType(mfcmapiNORMAL_RESTRICTION);
	}
} // CMailboxTableDlg::OnCreatePropertyStringRestriction