#include "stdafx.h"
#include "SingleRecipientDialog.h"
#include <UI/Controls/SingleMAPIPropListCtrl.h>
#include <MAPI/MAPIFunctions.h>
#include <UI/MFCUtilityFunctions.h>

namespace dialog
{
	static std::wstring CLASS = L"SingleRecipientDialog";

	SingleRecipientDialog::SingleRecipientDialog(
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		_In_opt_ LPMAILUSER lpMAPIProp) :
		CBaseDialog(
			pParentWnd,
			lpMapiObjects,
			NULL)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_szTitle = strings::loadstring(IDS_ADDRESS_BOOK_ENTRY);

		m_lpMailUser = lpMAPIProp;
		if (m_lpMailUser) m_lpMailUser->AddRef();

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
		auto bRet = CBaseDialog::OnInitDialog();

		if (m_lpMailUser)
		{
			// Get a property for the title bar
			m_szTitle = GetTitle(m_lpMailUser);
		}

		UpdateTitleBarText();

		auto hRes = S_OK;
		EC_H(m_lpPropDisplay->SetDataSource(m_lpMailUser, NULL, true));

		return bRet;
	}

	BEGIN_MESSAGE_MAP(SingleRecipientDialog, CBaseDialog)
		ON_COMMAND(ID_REFRESHVIEW, OnRefreshView)
	END_MESSAGE_MAP()

	// Clear the current list and get a new one with whatever code we've got in LoadMAPIPropList
	void SingleRecipientDialog::OnRefreshView()
	{
		if (!m_lpPropDisplay) return;
		(void)m_lpPropDisplay->RefreshMAPIPropList();
	}
}