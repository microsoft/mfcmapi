#include "stdafx.h"
#include "MailboxTableDlg.h"
#include <UI/Controls/ContentsTableListCtrl.h>
#include <MAPI/MapiObjects.h>
#include <MAPI/MAPIFunctions.h>
#include <MAPI/MAPIStoreFunctions.h>
#include <MAPI/ColumnTags.h>
#include <UI/MFCUtilityFunctions.h>
#include <UI/UIFunctions.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Dialogs/Editors/PropertyTagEditor.h>
#include <Interpret/InterpretProp2.h>
#include <UI/Controls/SortList/ContentsData.h>

static wstring CLASS = L"CMailboxTableDlg";

CMailboxTableDlg::CMailboxTableDlg(
	_In_ CParentWnd* pParentWnd,
	_In_ CMapiObjects* lpMapiObjects,
	_In_ const wstring& lpszServerName,
	_In_ LPMAPITABLE lpMAPITable
) :
	CContentsTableDlg(
		pParentWnd,
		lpMapiObjects,
		IDS_MAILBOXTABLE,
		mfcmapiDO_NOT_CALL_CREATE_DIALOG,
		lpMAPITable,
		LPSPropTagArray(&sptMBXCols),
		MBXColumns,
		IDR_MENU_MAILBOX_TABLE_POPUP,
		MENU_CONTEXT_MAILBOX_TABLE
	)
{
	TRACE_CONSTRUCTOR(CLASS);
	m_lpszServerName = lpszServerName;
	CMailboxTableDlg::CreateDialogAndMenu(IDR_MENU_MAILBOX_TABLE);
}

CMailboxTableDlg::~CMailboxTableDlg()
{
	TRACE_DESTRUCTOR(CLASS);
}

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
			pMenu->EnableMenuItem(ID_OPENWITHFLAGS, DIMMSOK(iNumSel));
		}
	}

	CContentsTableDlg::OnInitMenu(pMenu);
}

void CMailboxTableDlg::CreateDialogAndMenu(UINT nIDMenuResource)
{
	DebugPrintEx(DBGCreateDialog, CLASS, L"CreateDialogAndMenu", L"id = 0x%X\n", nIDMenuResource);
	CContentsTableDlg::CreateDialogAndMenu(nIDMenuResource);

	UpdateMenuString(
		m_hWnd,
		ID_CREATEPROPERTYSTRINGRESTRICTION,
		IDS_MBRESMENU);
}

void CMailboxTableDlg::DisplayItem(ULONG ulFlags)
{
	auto hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
	LPMDB lpGUIDMDB = nullptr;
	if (!lpMDB)
	{
		EC_H(OpenMessageStoreGUID(lpMAPISession, pbExchangeProviderPrimaryUserGuid, &lpGUIDMDB));
	}

	auto lpSourceMDB = lpMDB ? lpMDB : lpGUIDMDB; // do not release

	if (lpSourceMDB)
	{
		LPMDB lpNewMDB = nullptr;
		auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
		for (const auto& lpListData : items)
		{
			hRes = S_OK;
			if (lpListData && lpListData->Contents())
			{
				if (!lpListData->Contents()->m_szDN.empty())
				{
					EC_H(OpenOtherUsersMailbox(
						lpMAPISession,
						lpSourceMDB,
						strings::wstringTostring(m_lpszServerName),
						strings::wstringTostring(lpListData->Contents()->m_szDN),
						strings::emptystring,
						ulFlags,
						false,
						&lpNewMDB));

					if (lpNewMDB)
					{
						EC_H(DisplayObject(
							static_cast<LPMAPIPROP>(lpNewMDB),
							NULL,
							otStore,
							this));
						lpNewMDB->Release();
						lpNewMDB = nullptr;
					}
				}
			}
		}
	}

	if (lpGUIDMDB) lpGUIDMDB->Release();
}

void CMailboxTableDlg::OnDisplayItem()
{
	DisplayItem(OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_TAKE_OWNERSHIP);
}

void CMailboxTableDlg::OnOpenWithFlags()
{
	auto hRes = S_OK;

	CEditor MyPrompt(
		this,
		IDS_OPENWITHFLAGS,
		IDS_OPENWITHFLAGSPROMPT,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyPrompt.SetPromptPostFix(AllFlagsToString(PROP_ID(PR_PROFILE_OPEN_FLAGS), true));
	MyPrompt.InitPane(0, TextPane::CreateSingleLinePane(IDS_CREATESTORENTRYIDFLAGS, false));
	MyPrompt.SetHex(0, OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_TAKE_OWNERSHIP);
	WC_H(MyPrompt.DisplayDialog());
	if (S_OK == hRes)
	{
		DisplayItem(MyPrompt.GetHex(0));
	}
}

void CMailboxTableDlg::OnCreatePropertyStringRestriction()
{
	auto hRes = S_OK;
	LPSRestriction lpRes = nullptr;

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
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.SetPromptPostFix(AllFlagsToString(flagFuzzyLevel, true));

		MyData.InitPane(0, TextPane::CreateSingleLinePane(IDS_NAME, false));
		MyData.InitPane(1, TextPane::CreateSingleLinePane(IDS_ULFUZZYLEVEL, false));
		MyData.SetHex(1, FL_IGNORECASE | FL_PREFIX);

		WC_H(MyData.DisplayDialog());
		if (S_OK != hRes) return;

		auto szString = MyData.GetStringW(0);
		// Allocate and create our SRestriction
		EC_H(CreatePropertyStringRestriction(
			CHANGE_PROP_TYPE(MyPropertyTag.GetPropertyTag(), PT_UNICODE),
			szString,
			MyData.GetHex(1),
			NULL,
			&lpRes));
		if (hRes != S_OK)
		{
			MAPIFreeBuffer(lpRes);
			lpRes = nullptr;
		}

		m_lpContentsTableListCtrl->SetRestriction(lpRes);

		SetRestrictionType(mfcmapiNORMAL_RESTRICTION);
	}
}