#include <StdAfx.h>
#include <UI/Dialogs/ContentsTable/MailboxTableDlg.h>
#include <UI/Controls/SortList/ContentsTableListCtrl.h>
#include <core/mapi/cache/mapiObjects.h>
#include <core/mapi/mapiFunctions.h>
#include <core/mapi/mapiStoreFunctions.h>
#include <core/mapi/columnTags.h>
#include <UI/Dialogs/MFCUtilityFunctions.h>
#include <UI/UIFunctions.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/Dialogs/Editors/PropertyTagEditor.h>
#include <core/sortlistdata/contentsData.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/interpret/flags.h>

namespace dialog
{
	static std::wstring CLASS = L"CMailboxTableDlg";

	CMailboxTableDlg::CMailboxTableDlg(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		_In_ const std::wstring& lpszServerName,
		_In_ LPMAPITABLE lpMAPITable)
		: CContentsTableDlg(
			  pParentWnd,
			  lpMapiObjects,
			  IDS_MAILBOXTABLE,
			  createDialogType::DO_NOT_CALL_CREATE_DIALOG,
			  nullptr,
			  lpMAPITable,
			  &columns::sptMBXCols.tags,
			  columns::MBXColumns,
			  IDR_MENU_MAILBOX_TABLE_POPUP,
			  MENU_CONTEXT_MAILBOX_TABLE)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_lpszServerName = lpszServerName;
		CMailboxTableDlg::CreateDialogAndMenu(IDR_MENU_MAILBOX_TABLE);
	}

	CMailboxTableDlg::~CMailboxTableDlg() { TRACE_DESTRUCTOR(CLASS); }

	BEGIN_MESSAGE_MAP(CMailboxTableDlg, CContentsTableDlg)
	ON_COMMAND(ID_OPENWITHFLAGS, OnOpenWithFlags)
	END_MESSAGE_MAP()

	void CMailboxTableDlg::OnInitMenu(_In_ CMenu* pMenu)
	{
		if (pMenu)
		{
			if (m_lpContentsTableListCtrl)
			{
				const int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
				pMenu->EnableMenuItem(ID_OPENWITHFLAGS, DIMMSOK(iNumSel));
			}
		}

		CContentsTableDlg::OnInitMenu(pMenu);
	}

	void CMailboxTableDlg::CreateDialogAndMenu(UINT nIDMenuResource)
	{
		output::DebugPrintEx(
			output::dbgLevel::CreateDialog, CLASS, L"CreateDialogAndMenu", L"id = 0x%X\n", nIDMenuResource);
		CContentsTableDlg::CreateDialogAndMenu(nIDMenuResource);

		ui::UpdateMenuString(m_hWnd, ID_CREATEPROPERTYSTRINGRESTRICTION, IDS_MBRESMENU);
	}

	void CMailboxTableDlg::DisplayItem(ULONG ulFlags)
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		auto lpMDB = m_lpMapiObjects->GetMDB(); // do not release
		LPMDB lpGUIDMDB = nullptr;
		if (!lpMDB)
		{
			lpGUIDMDB = mapi::store::OpenMessageStoreGUID(lpMAPISession, pbExchangeProviderPrimaryUserGuid);
		}

		const auto lpSourceMDB = lpMDB ? lpMDB : lpGUIDMDB; // do not release

		if (lpSourceMDB)
		{
			auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
			for (const auto& lpListData : items)
			{
				if (lpListData)
				{
					const auto contents = lpListData->cast<sortlistdata::contentsData>();
					if (contents)
					{
						if (!contents->m_szDN.empty())
						{
							auto lpNewMDB = mapi::store::OpenOtherUsersMailbox(
								lpMAPISession,
								lpSourceMDB,
								m_lpszServerName,
								contents->m_szDN,
								strings::emptystring,
								ulFlags,
								false);
							if (lpNewMDB)
							{
								EC_H_S(
									DisplayObject(static_cast<LPMAPIPROP>(lpNewMDB), NULL, ObjectType::otStore, this));
								lpNewMDB->Release();
								lpNewMDB = nullptr;
							}
						}
					}
				}
			}
		}

		if (lpGUIDMDB) lpGUIDMDB->Release();
	}

	void CMailboxTableDlg::OnDisplayItem() { DisplayItem(OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_TAKE_OWNERSHIP); }

	void CMailboxTableDlg::OnOpenWithFlags()
	{
		editor::CEditor MyPrompt(
			this, IDS_OPENWITHFLAGS, IDS_OPENWITHFLAGSPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyPrompt.SetPromptPostFix(flags::AllFlagsToString(PROP_ID(PR_PROFILE_OPEN_FLAGS), true));
		MyPrompt.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_CREATESTORENTRYIDFLAGS, false));
		MyPrompt.SetHex(0, OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_TAKE_OWNERSHIP);
		if (MyPrompt.DisplayDialog())
		{
			DisplayItem(MyPrompt.GetHex(0));
		}
	}

	void CMailboxTableDlg::OnCreatePropertyStringRestriction()
	{
		editor::CPropertyTagEditor MyPropertyTag(
			IDS_PROPRES,
			NULL, // prompt
			PR_DISPLAY_NAME,
			m_bIsAB,
			nullptr,
			this);

		if (!MyPropertyTag.DisplayDialog()) return;

		editor::CEditor MyData(
			this, IDS_SEARCHCRITERIA, IDS_MBSEARCHCRITERIAPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.SetPromptPostFix(flags::AllFlagsToString(flagFuzzyLevel, true));

		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_NAME, false));
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_ULFUZZYLEVEL, false));
		MyData.SetHex(1, FL_IGNORECASE | FL_PREFIX);

		if (!MyData.DisplayDialog()) return;

		const auto szString = MyData.GetStringW(0);
		// Allocate and create our SRestriction
		const auto lpRes = mapi::CreatePropertyStringRestriction(
			CHANGE_PROP_TYPE(MyPropertyTag.GetPropertyTag(), PT_UNICODE), szString, MyData.GetHex(1), nullptr);
		m_lpContentsTableListCtrl->SetRestriction(lpRes);

		SetRestrictionType(restrictionType::normal);
	}
} // namespace dialog