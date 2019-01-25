#include <StdAfx.h>
#include <UI/Dialogs/SingleRecipientDialog.h>
#include <UI/Controls/SingleMAPIPropListCtrl.h>
#include <UI/Dialogs/MFCUtilityFunctions.h>
#include <core/utility/strings.h>
#include <core/mapi/mapiFunctions.h>
#include <core/utility/output.h>

namespace dialog
{
	static std::wstring CLASS = L"SingleRecipientDialog";

	SingleRecipientDialog::SingleRecipientDialog(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ cache::CMapiObjects* lpMapiObjects,
		_In_opt_ LPMAPIPROP lpMAPIProp)
		: CBaseDialog(pParentWnd, lpMapiObjects, NULL)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_szTitle = strings::loadstring(IDS_ADDRESS_BOOK_ENTRY);

		m_lpMailUser = mapi::safe_cast<LPMAILUSER>(lpMAPIProp);

		CBaseDialog::CreateDialogAndMenu(NULL, NULL, NULL);
	}

	SingleRecipientDialog::~SingleRecipientDialog()
	{
		TRACE_DESTRUCTOR(CLASS);
		if (m_lpMailUser) m_lpMailUser->Release();
		m_lpMailUser = nullptr;
	}

	BOOL SingleRecipientDialog::OnInitDialog()
	{
		const auto bRet = CBaseDialog::OnInitDialog();

		if (m_lpMailUser)
		{
			// Get a property for the title bar
			m_szTitle = mapi::GetTitle(m_lpMailUser);
		}

		UpdateTitleBarText();

		EC_H_S(m_lpPropDisplay->SetDataSource(m_lpMailUser, NULL, true));

		return bRet;
	}

	BEGIN_MESSAGE_MAP(SingleRecipientDialog, CBaseDialog)
	ON_COMMAND(ID_REFRESHVIEW, OnRefreshView)
	END_MESSAGE_MAP()

	// Clear the current list and get a new one with whatever code we've got in LoadMAPIPropList
	void SingleRecipientDialog::OnRefreshView()
	{
		if (!m_lpPropDisplay) return;
		(void) m_lpPropDisplay->RefreshMAPIPropList();
	}
} // namespace dialog
