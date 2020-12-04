#include <StdAfx.h>
#include <UI/Dialogs/propList/RegistryDialog.h>
//#include <UI/Controls/SortList/SingleMAPIPropListCtrl.h>
//#include <UI/Dialogs/MFCUtilityFunctions.h>
//#include <core/utility/strings.h>
#include <core/mapi/mapiFunctions.h>
#include <core/utility/output.h>

namespace dialog
{
	static std::wstring CLASS = L"RegistryDialog";

	RegistryDialog::RegistryDialog(
		_In_ ui::CParentWnd* pParentWnd,
		_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects)
		: CBaseDialog(pParentWnd, lpMapiObjects, NULL)
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_szTitle = strings::loadstring(IDS_REGISTRY_DIALOG);

		CBaseDialog::CreateDialogAndMenu(NULL, NULL, NULL);
	}

	RegistryDialog::~RegistryDialog() { TRACE_DESTRUCTOR(CLASS); }

	//void RegistryDialog::InitAccountManager()
	//{
	//}

	BOOL RegistryDialog::OnInitDialog()
	{
		const auto bRet = CBaseDialog::OnInitDialog();

		UpdateTitleBarText();

		if (m_lpFakeSplitter)
		{
			m_lpRegKeyList.Create(&m_lpFakeSplitter, true);
			m_lpFakeSplitter.SetPaneOne(m_lpRegKeyList.GetSafeHwnd());
			//m_lpRegKeyList.FreeNodeDataCallback = [&](auto lpAccount) {
			//	if (lpAccount) reinterpret_cast<LPOLKACCOUNT>(lpAccount)->Release();
			//};
			m_lpRegKeyList.ItemSelectedCallback = [&](auto /*hItem*/) {
				//auto lpAccount = reinterpret_cast<LPOLKACCOUNT>(m_lpRegKeyList.GetItemData(hItem));
				//EC_H_S(m_lpPropDisplay->SetDataSource(
				//	std::make_shared<propertybag::accountPropertyBag>(m_lpwszProfile, lpAccount), false));
			};

			EnumRegistry();

			m_lpFakeSplitter.SetPercent(0.25);
		}

		return bRet;
	}

	//void RegistryDialog::EnumRegistry(const std::wstring& szCat, const CLSID* pclsidCategory)
	//{
	//	const auto root = m_lpAccountsList.AddChildNode(szCat.c_str(), TVI_ROOT, nullptr, nullptr);
	//	LPOLKENUM lpAcctEnum = nullptr;

	//	EC_H_S(m_lpAcctMgr->EnumerateAccounts(pclsidCategory, nullptr, OLK_ACCOUNT_NO_FLAGS, &lpAcctEnum));
	//	if (lpAcctEnum)
	//	{
	//		DWORD cAccounts = 0;

	//		EC_H_S(lpAcctEnum->GetCount(&cAccounts));
	//		if (cAccounts)
	//		{
	//			EC_H_S(lpAcctEnum->Reset());
	//			DWORD iAcct = 0;
	//			for (iAcct = 0; iAcct < cAccounts; iAcct++)
	//			{
	//				LPUNKNOWN lpUnk = nullptr;

	//				EC_H_S(lpAcctEnum->GetNext(&lpUnk));
	//				if (lpUnk)
	//				{
	//					auto lpAccount = mapi::safe_cast<LPOLKACCOUNT>(lpUnk);
	//					if (lpAccount)
	//					{
	//						auto acctName = ACCT_VARIANT{};
	//						auto nodeName = strings::format(L"Account %i", iAcct);
	//						const auto hRes = WC_H(lpAccount->GetProp(PROP_ACCT_NAME, &acctName));
	//						if (SUCCEEDED(hRes))
	//						{
	//							nodeName = acctName.Val.pwsz;
	//							WC_H_S(lpAccount->FreeMemory(reinterpret_cast<LPBYTE>(acctName.Val.pwsz)));
	//						}

	//						static_cast<void>(m_lpAccountsList.AddChildNode(nodeName, root, lpAccount, nullptr));
	//					}
	//				}

	//				if (lpUnk) lpUnk->Release();
	//				lpUnk = nullptr;
	//			}
	//		}

	//		lpAcctEnum->Release();
	//	}
	//}

	void RegistryDialog::EnumRegistry()
	{
		//EnumAccounts(L"Mail", &CLSID_OlkMail);
		//EnumAccounts(L"Address Book", &CLSID_OlkAddressBook);
		//EnumAccounts(L"Store", &CLSID_OlkStore);
	}

	BEGIN_MESSAGE_MAP(RegistryDialog, CBaseDialog)
	END_MESSAGE_MAP()
} // namespace dialog
