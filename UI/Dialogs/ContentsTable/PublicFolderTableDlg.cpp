#include <StdAfx.h>
#include "PublicFolderTableDlg.h"
#include <UI/Controls/ContentsTableListCtrl.h>
#include <MAPI/MAPIFunctions.h>
#include <MAPI/ColumnTags.h>
#include <UI/UIFunctions.h>
#include <UI/Dialogs/Editors/PropertyTagEditor.h>
#include <Interpret/InterpretProp.h>

namespace dialog
{
	static std::wstring CLASS = L"CPublicFolderTableDlg";

	CPublicFolderTableDlg::CPublicFolderTableDlg(
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		_In_ const std::wstring& lpszServerName,
		_In_ LPMAPITABLE lpMAPITable
	) :
		CContentsTableDlg(
			pParentWnd,
			lpMapiObjects,
			IDS_PUBLICFOLDERTABLE,
			mfcmapiDO_NOT_CALL_CREATE_DIALOG,
			lpMAPITable,
			LPSPropTagArray(&sptPFCols),
			PFColumns,
			NULL,
			MENU_CONTEXT_PUBLIC_FOLDER_TABLE)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_lpszServerName = lpszServerName;
		CPublicFolderTableDlg::CreateDialogAndMenu(NULL);
	}

	CPublicFolderTableDlg::~CPublicFolderTableDlg()
	{
		TRACE_DESTRUCTOR(CLASS);
	}

	void CPublicFolderTableDlg::CreateDialogAndMenu(UINT nIDMenuResource)
	{
		DebugPrintEx(DBGCreateDialog, CLASS, L"CreateDialogAndMenu", L"id = 0x%X\n", nIDMenuResource);
		CContentsTableDlg::CreateDialogAndMenu(nIDMenuResource);

		UpdateMenuString(
			m_hWnd,
			ID_CREATEPROPERTYSTRINGRESTRICTION,
			IDS_PFRESMENU);
	}

	void CPublicFolderTableDlg::OnDisplayItem()
	{
	}

	_Check_return_ HRESULT CPublicFolderTableDlg::OpenItemProp(int /*iSelectedItem*/, __mfcmapiModifyEnum /*bModify*/, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp)
	{
		if (lppMAPIProp) *lppMAPIProp = nullptr;
		return S_OK;
	}

	void CPublicFolderTableDlg::OnCreatePropertyStringRestriction()
	{
		auto hRes = S_OK;
		LPSRestriction lpRes = nullptr;

		editor::CPropertyTagEditor MyPropertyTag(
			IDS_PROPRES, // title
			NULL, // prompt
			PR_DISPLAY_NAME,
			m_bIsAB,
			m_lpContainer,
			this);

		WC_H(MyPropertyTag.DisplayDialog());
		if (S_OK == hRes)
		{
			editor::CEditor MyData(
				this,
				IDS_SEARCHCRITERIA,
				IDS_PFSEARCHCRITERIAPROMPT,
				CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyData.SetPromptPostFix(interpretprop::AllFlagsToString(flagFuzzyLevel, true));

			MyData.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_NAME, false));
			MyData.InitPane(1, viewpane::TextPane::CreateSingleLinePane(IDS_ULFUZZYLEVEL, false));
			MyData.SetHex(1, FL_IGNORECASE | FL_PREFIX);

			WC_H(MyData.DisplayDialog());
			if (S_OK != hRes) return;

			const auto szString = MyData.GetStringW(0);
			// Allocate and create our SRestriction
			EC_H(CreatePropertyStringRestriction(
				CHANGE_PROP_TYPE(MyPropertyTag.GetPropertyTag(), PT_UNICODE),
				szString,
				MyData.GetHex(1),
				NULL,
				&lpRes));
			if (S_OK != hRes && lpRes)
			{
				MAPIFreeBuffer(lpRes);
				lpRes = nullptr;
			}

			m_lpContentsTableListCtrl->SetRestriction(lpRes);

			SetRestrictionType(mfcmapiNORMAL_RESTRICTION);
		}
	}
}