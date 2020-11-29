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

	std::wstring GetProfileName(LPMAPISESSION lpSession)
	{
		LPPROFSECT lpProfSect = nullptr;
		std::wstring profileName;

		if (!lpSession) return profileName;

		EC_H_S(lpSession->OpenProfileSection(LPMAPIUID(pbGlobalProfileSectionGuid), nullptr, 0, &lpProfSect));
		if (lpProfSect)
		{
			LPSPropValue lpProfileName = nullptr;

			EC_H_S(HrGetOneProp(lpProfSect, PR_PROFILE_NAME_W, &lpProfileName));
			if (lpProfileName && lpProfileName->ulPropTag == PR_PROFILE_NAME_W)
			{
				profileName = std::wstring(lpProfileName->Value.lpszW);
			}

			MAPIFreeBuffer(lpProfileName);
		}

		if (lpProfSect) lpProfSect->Release();

		return profileName;
	}

	AccountsDialog::AccountsDialog(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
		_In_opt_ LPMAPISESSION lpMAPISession)
		: CBaseDialog(pParentWnd, lpMapiObjects, NULL)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_szTitle = strings::loadstring(IDS_ACCOUNT_DIALOG);

		m_lpMAPISession = mapi::safe_cast<LPMAPISESSION>(lpMAPISession);

		CBaseDialog::CreateDialogAndMenu(NULL, NULL, NULL);
	}

	AccountsDialog::~AccountsDialog()
	{
		TRACE_DESTRUCTOR(CLASS);
		m_lpAccountsList.DestroyWindow();

		if (m_lpMyAcctHelper) m_lpMyAcctHelper->Release();
		m_lpMyAcctHelper = nullptr;

		if (m_lpAcctMgr) m_lpAcctMgr->Release();
		m_lpAcctMgr = nullptr;

		if (m_lpMAPISession) m_lpMAPISession->Release();
		m_lpMAPISession = nullptr;
	}

	void AccountsDialog::InitAccountManager()
	{
		if (m_lpMAPISession)
		{
			m_lpwszProfile = GetProfileName(m_lpMAPISession);
			m_szTitle = m_lpwszProfile;
			m_lpMAPISession = mapi::safe_cast<LPMAPISESSION>(m_lpMAPISession);
			EC_H_S(CoCreateInstance(
				CLSID_OlkAccountManager,
				nullptr,
				CLSCTX_INPROC_SERVER,
				IID_IOlkAccountManager,
				reinterpret_cast<LPVOID*>(&m_lpAcctMgr)));
		}

		if (m_lpAcctMgr)
		{
			m_lpMyAcctHelper = new CAccountHelper(m_lpwszProfile.c_str(), m_lpMAPISession);
		}

		if (m_lpMyAcctHelper)
		{
			EC_H_S(m_lpAcctMgr->Init(m_lpMyAcctHelper, ACCT_INIT_NOSYNCH_MAPI_ACCTS));
		}
	}

	BOOL AccountsDialog::OnInitDialog()
	{
		const auto bRet = CBaseDialog::OnInitDialog();

		InitAccountManager();

		UpdateTitleBarText();

		if (m_lpFakeSplitter)
		{
			m_lpAccountsList.Create(&m_lpFakeSplitter, true);
			m_lpFakeSplitter.SetPaneOne(m_lpAccountsList.GetSafeHwnd());
			m_lpAccountsList.FreeNodeDataCallback = [&](auto lpAccount) {
				if (lpAccount) reinterpret_cast<LPOLKACCOUNT>(lpAccount)->Release();
			};
			m_lpAccountsList.ItemSelectedCallback = [&](auto hItem) {
				auto lpAccount = reinterpret_cast<LPOLKACCOUNT>(m_lpAccountsList.GetItemData(hItem));
				EC_H_S(m_lpPropDisplay->SetDataSource(
					std::make_shared<propertybag::accountPropertyBag>(m_lpwszProfile, lpAccount), false));
			};

			EnumAccounts();

			m_lpFakeSplitter.SetPercent(0.25);
		}

		return bRet;
	}

	void AccountsDialog::EnumAccounts()
	{
		LPOLKENUM lpAcctEnum = nullptr;

		EC_H_S(m_lpAcctMgr->EnumerateAccounts(&CLSID_OlkMail, nullptr, OLK_ACCOUNT_NO_FLAGS, &lpAcctEnum));
		if (lpAcctEnum)
		{
			DWORD cAccounts = 0;

			EC_H_S(lpAcctEnum->GetCount(&cAccounts));
			if (cAccounts)
			{
				EC_H_S(lpAcctEnum->Reset());
				DWORD iAcct = 0;
				for (iAcct = 0; iAcct < cAccounts; iAcct++)
				{
					LPUNKNOWN lpUnk = nullptr;

					EC_H_S(lpAcctEnum->GetNext(&lpUnk));
					if (lpUnk)
					{
						auto lpAccount = mapi::safe_cast<LPOLKACCOUNT>(lpUnk);
						if (lpAccount)
						{
							// TODO: Add real identifier here
							auto nodeName = strings::format(L"Account %i", iAcct);
							m_lpAccountsList.AddChildNode(nodeName, TVI_ROOT, lpAccount, nullptr);
						}
					}

					if (lpUnk) lpUnk->Release();
					lpUnk = nullptr;
				}
			}
		}

		if (lpAcctEnum) lpAcctEnum->Release();
	}

	BEGIN_MESSAGE_MAP(AccountsDialog, CBaseDialog)
	ON_COMMAND(ID_REFRESHVIEW, OnRefreshView)
	END_MESSAGE_MAP()

	// Clear the current list and get a new one with whatever code we've got in LoadMAPIPropList
	// TODO: This is probably wrong
	void AccountsDialog::OnRefreshView()
	{
		if (!m_lpPropDisplay) return;
		static_cast<void>(m_lpPropDisplay->RefreshMAPIPropList());
	}
} // namespace dialog
