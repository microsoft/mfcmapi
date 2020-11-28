#include <StdAfx.h>
#include <UI/Dialogs/propList/AccountsDialog.h>
#include <UI/Controls/SortList/SingleMAPIPropListCtrl.h>
#include <UI/Dialogs/MFCUtilityFunctions.h>
#include <core/utility/strings.h>
#include <core/mapi/mapiFunctions.h>
#include <core/utility/output.h>
#include <core/propertyBag/accountPropertyBag.h>

namespace dialog
{
	static std::wstring CLASS = L"AccountsDialog";

	AccountsDialog::AccountsDialog(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		_In_opt_ LPMAPISESSION lpMAPISession)
		: CBaseDialog(pParentWnd, lpMapiObjects, NULL)
	{
		TRACE_CONSTRUCTOR(CLASS);
		// TODO: make resource from this
		m_szTitle = L"Account dialog"; //strings::loadstring(IDS_ADDRESS_BOOK_ENTRY);

		m_lpMAPISession = mapi::safe_cast<LPMAPISESSION>(lpMAPISession);

		CBaseDialog::CreateDialogAndMenu(NULL, NULL, NULL);
	}

	AccountsDialog::~AccountsDialog()
	{
		TRACE_DESTRUCTOR(CLASS);
		if (m_lpMAPISession) m_lpMAPISession->Release();
		m_lpMAPISession = nullptr;
	}

	BOOL AccountsDialog::OnInitDialog()
	{
		const auto bRet = CBaseDialog::OnInitDialog();

		if (m_lpMAPISession)
		{
			// Get a property for the title bar
			// TODO: get a good title for account
			//			m_szTitle = mapi::GetTitle(m_lpMAPISession);
			m_szTitle = L"Custom title here";
		}

		UpdateTitleBarText();

		// TODO: custom data source for accounts here
		EC_H_S(
			m_lpPropDisplay->SetDataSource(std::make_shared<propertybag::accountPropertyBag>(m_lpMAPISession), false));

		return bRet;
	}

	BEGIN_MESSAGE_MAP(AccountsDialog, CBaseDialog)
	ON_COMMAND(ID_REFRESHVIEW, OnRefreshView)
	END_MESSAGE_MAP()

	// Clear the current list and get a new one with whatever code we've got in LoadMAPIPropList
	void AccountsDialog::OnRefreshView()
	{
		if (!m_lpPropDisplay) return;
		static_cast<void>(m_lpPropDisplay->RefreshMAPIPropList());
	}
} // namespace dialog
