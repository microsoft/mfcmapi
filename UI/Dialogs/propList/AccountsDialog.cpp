#include <StdAfx.h>
#include <UI/Dialogs/propList/AccountsDialog.h>
#include <UI/Controls/SortList/SingleMAPIPropListCtrl.h>
#include <core/mapi/mapiFunctions.h>
#include <core/utility/output.h>
#include <core/propertyBag/accountPropertyBag.h>
#include <UI/UIFunctions.h>

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
			m_lpwszProfile = mapi::GetProfileName(m_lpMAPISession);
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
				m_lpPropDisplay->SetDataSource(
					std::make_shared<propertybag::accountPropertyBag>(m_lpwszProfile, lpAccount));
			};

			EnumCategories();

			m_lpFakeSplitter.SetPercent(0.25);
		}

		return bRet;
	}

	void AccountsDialog::OnInitMenu(_In_ CMenu* pMenu)
	{
		CBaseDialog::OnInitMenu(pMenu);
		if (pMenu)
		{
			ui::DeleteMenu(pMenu->m_hMenu, ID_DISPLAYPROPERTYASSECURITYDESCRIPTORPROPSHEET);
			ui::DeleteMenu(pMenu->m_hMenu, ID_ATTACHMENTPROPERTIES);
			ui::DeleteMenu(pMenu->m_hMenu, ID_RECIPIENTPROPERTIES);
			ui::DeleteMenu(pMenu->m_hMenu, ID_RTFSYNC);
			ui::DeleteMenu(pMenu->m_hMenu, ID_TESTEDITBODY);
			ui::DeleteMenu(pMenu->m_hMenu, ID_TESTEDITHTML);
			ui::DeleteMenu(pMenu->m_hMenu, ID_TESTEDITRTF);

			// Locate and delete "Edit as stream" menu by searching for an item on it
			WC_B_S(ui::DeleteSubmenu(pMenu->m_hMenu, ID_EDITPROPERTYASASCIISTREAM));
		}
	}

	void AccountsDialog::EnumAccounts(const std::wstring& szCat, const CLSID* pclsidCategory)
	{
		const auto root = m_lpAccountsList.AddChildNode(szCat.c_str(), TVI_ROOT, nullptr, nullptr);
		LPOLKENUM lpAcctEnum = nullptr;

		EC_H_S(m_lpAcctMgr->EnumerateAccounts(pclsidCategory, nullptr, OLK_ACCOUNT_NO_FLAGS, &lpAcctEnum));
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
							auto acctName = ACCT_VARIANT{};
							auto nodeName = strings::format(L"Account %i", iAcct);
							const auto hRes = WC_H(lpAccount->GetProp(PROP_ACCT_NAME, &acctName));
							if (SUCCEEDED(hRes))
							{
								nodeName = acctName.Val.pwsz;
								WC_H_S(lpAccount->FreeMemory(reinterpret_cast<LPBYTE>(acctName.Val.pwsz)));
							}

							static_cast<void>(m_lpAccountsList.AddChildNode(nodeName, root, lpAccount, nullptr));
						}
					}

					if (lpUnk) lpUnk->Release();
					lpUnk = nullptr;
				}
			}

			lpAcctEnum->Release();
		}
	}

	void AccountsDialog::EnumCategories()
	{
		EnumAccounts(L"Mail", &CLSID_OlkMail);
		EnumAccounts(L"Address Book", &CLSID_OlkAddressBook);
		EnumAccounts(L"Store", &CLSID_OlkStore);
	}

	BEGIN_MESSAGE_MAP(AccountsDialog, CBaseDialog)
	END_MESSAGE_MAP()
} // namespace dialog
