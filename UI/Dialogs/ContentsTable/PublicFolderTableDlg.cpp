#include <StdAfx.h>
#include <UI/Dialogs/ContentsTable/PublicFolderTableDlg.h>
#include <UI/Controls/ContentsTableListCtrl.h>
#include <core/mapi/mapiFunctions.h>
#include <core/mapi/columnTags.h>
#include <UI/UIFunctions.h>
#include <UI/Dialogs/Editors/PropertyTagEditor.h>
#include <core/utility/output.h>
#include <core/interpret/flags.h>

namespace dialog
{
	static std::wstring CLASS = L"CPublicFolderTableDlg";

	CPublicFolderTableDlg::CPublicFolderTableDlg(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ cache::CMapiObjects* lpMapiObjects,
		_In_ const std::wstring& lpszServerName,
		_In_ LPMAPITABLE lpMAPITable)
		: CContentsTableDlg(
			  pParentWnd,
			  lpMapiObjects,
			  IDS_PUBLICFOLDERTABLE,
			  mfcmapiDO_NOT_CALL_CREATE_DIALOG,
			  nullptr,
			  lpMAPITable,
			  &columns::sptPFCols.tags,
			  columns::PFColumns,
			  NULL,
			  MENU_CONTEXT_PUBLIC_FOLDER_TABLE)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_lpszServerName = lpszServerName;
		CPublicFolderTableDlg::CreateDialogAndMenu(NULL);
	}

	CPublicFolderTableDlg::~CPublicFolderTableDlg() { TRACE_DESTRUCTOR(CLASS); }

	void CPublicFolderTableDlg::CreateDialogAndMenu(UINT nIDMenuResource)
	{
		output::DebugPrintEx(output::DBGCreateDialog, CLASS, L"CreateDialogAndMenu", L"id = 0x%X\n", nIDMenuResource);
		CContentsTableDlg::CreateDialogAndMenu(nIDMenuResource);

		ui::UpdateMenuString(m_hWnd, ID_CREATEPROPERTYSTRINGRESTRICTION, IDS_PFRESMENU);
	}

	void CPublicFolderTableDlg::OnDisplayItem() {}

	_Check_return_ LPMAPIPROP
	CPublicFolderTableDlg::OpenItemProp(int /*iSelectedItem*/, __mfcmapiModifyEnum /*bModify*/)
	{
		return nullptr;
	}

	void CPublicFolderTableDlg::OnCreatePropertyStringRestriction()
	{
		editor::CPropertyTagEditor MyPropertyTag(
			IDS_PROPRES, // title
			NULL, // prompt
			PR_DISPLAY_NAME,
			m_bIsAB,
			nullptr,
			this);

		if (!MyPropertyTag.DisplayDialog()) return;

		editor::CEditor MyData(
			this, IDS_SEARCHCRITERIA, IDS_PFSEARCHCRITERIAPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
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

		SetRestrictionType(mfcmapiNORMAL_RESTRICTION);
	}
} // namespace dialog