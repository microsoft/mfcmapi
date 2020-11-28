#include <core/stdafx.h>
#include <core/propertyBag/accountPropertyBag.h>
#include <core/mapi/mapiFunctions.h>
#include <core/mapi/mapiMemory.h>
#include <core/mapi/account/actMgmt.h>
#include <core/mapi/account/accountHelper.h>

namespace propertybag
{
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

	accountPropertyBag::accountPropertyBag(LPMAPISESSION lpMAPISession)
	{
		if (lpMAPISession)
		{
			m_lpwszProfile = GetProfileName(lpMAPISession);
			m_lpMAPISession = mapi::safe_cast<LPMAPISESSION>(lpMAPISession);
			EC_H_S(CoCreateInstance(
				CLSID_OlkAccountManager,
				nullptr,
				CLSCTX_INPROC_SERVER,
				IID_IOlkAccountManager,
				reinterpret_cast<LPVOID*>(&m_lpAcctMgr)));
		}

		if (m_lpAcctMgr)
		{
			m_lpMyAcctHelper = new CAccountHelper(m_lpwszProfile.c_str(), lpMAPISession);
		}

		if (m_lpMyAcctHelper)
		{
			EC_H_S(m_lpMyAcctHelper->QueryInterface(IID_IOlkAccountHelper, reinterpret_cast<LPVOID*>(&m_lpAcctHelper)));
		}

		if (m_lpAcctHelper)
		{
			EC_H_S(m_lpAcctMgr->Init(m_lpAcctHelper, ACCT_INIT_NOSYNCH_MAPI_ACCTS));
		}
	}

	accountPropertyBag ::~accountPropertyBag()
	{
		if (m_lpAcctHelper) m_lpAcctHelper->Release();
		m_lpAcctHelper = nullptr;

		if (m_lpMyAcctHelper) m_lpMyAcctHelper->Release();
		m_lpMyAcctHelper = nullptr;

		if (m_lpAcctMgr) m_lpAcctMgr->Release();
		m_lpAcctMgr = nullptr;

		if (m_lpMAPISession) m_lpMAPISession->Release();
		m_lpMAPISession = nullptr;
	}

	propBagFlags accountPropertyBag::GetFlags() const
	{
		auto ulFlags = propBagFlags::None;
		return ulFlags;
	}

	bool accountPropertyBag::IsEqual(const std::shared_ptr<IMAPIPropertyBag> lpPropBag) const
	{
		if (!lpPropBag) return false;
		if (GetType() != lpPropBag->GetType()) return false;

		// Two accounts are the same if their profile is the same
		const auto lpOther = std::dynamic_pointer_cast<accountPropertyBag>(lpPropBag);
		if (lpOther)
		{
			if (m_lpwszProfile != lpOther->m_lpwszProfile) return false;
			return true;
		}

		return false;
	}

	_Check_return_ HRESULT accountPropertyBag::GetAllProps(ULONG FAR* lpcValues, LPSPropValue FAR* lppPropArray)
	{
		if (!lpcValues || !lppPropArray) return MAPI_E_INVALID_PARAMETER;

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
				// TODO: For now we just get the first account - later we need a content list type view of accounts
				for (iAcct = 0; iAcct < 1; iAcct++)
				{
					LPUNKNOWN lpUnk = nullptr;

					EC_H_S(lpAcctEnum->GetNext(&lpUnk));
					if (lpUnk)
					{
						LPOLKACCOUNT lpAccount = nullptr;

						EC_H_S(lpUnk->QueryInterface(IID_IOlkAccount, reinterpret_cast<LPVOID*>(&lpAccount)));
						if (lpAccount)
						{
							// Just return what we have
							auto hRes = S_OK;
							ULONG i = 0;
							std::vector<std::pair<ULONG, ACCT_VARIANT>> props = {};

							for (i = 0; i < 0x8000; i++)
							{
								ACCT_VARIANT pProp = {};
								hRes = lpAccount->GetProp(PROP_TAG(PT_LONG, i), &pProp);
								if (SUCCEEDED(hRes))
								{
									props.emplace_back(PROP_TAG(PT_LONG, i), pProp);
								}

								hRes = lpAccount->GetProp(PROP_TAG(PT_UNICODE, i), &pProp);
								if (SUCCEEDED(hRes))
								{
									props.emplace_back(PROP_TAG(PT_UNICODE, i), pProp);
								}

								hRes = lpAccount->GetProp(PROP_TAG(PT_BINARY, i), &pProp);
								if (SUCCEEDED(hRes))
								{
									props.emplace_back(PROP_TAG(PT_BINARY, i), pProp);
								}
							}

							if (props.size() > 0)
							{
								*lpcValues = props.size();
								*lppPropArray = mapi::allocate<LPSPropValue>(props.size() * sizeof(SPropValue));
								auto iProp = 0;

								for (const auto prop : props)
								{
									(*lppPropArray)[iProp].ulPropTag = prop.first;
									switch (PROP_TYPE(prop.first))
									{
									case PT_LONG:
										(*lppPropArray)[iProp].Value.l = prop.second.Val.dw;
										break;
									case PT_UNICODE:
										(*lppPropArray)[iProp].Value.lpszW =
											mapi::CopyStringW(prop.second.Val.pwsz, *lppPropArray);
										(void) lpAccount->FreeMemory(reinterpret_cast<LPBYTE>(prop.second.Val.pwsz));
										break;
									case PT_BINARY:
										auto bin = SBinary{prop.second.Val.bin.cb, prop.second.Val.bin.pb};
										(*lppPropArray)[iProp].Value.bin = mapi::CopySBinary(bin, *lppPropArray);
										(void) lpAccount->FreeMemory(reinterpret_cast<LPBYTE>(prop.second.Val.bin.pb));
										break;
									}

									iProp++;
								}
							}

							lpAccount->Release();
						}
					}

					if (lpUnk) lpUnk->Release();
					lpUnk = nullptr;
				}
			}
		}

		if (lpAcctEnum) lpAcctEnum->Release();
		return S_OK;
	}

	_Check_return_ HRESULT accountPropertyBag::GetProps(
		LPSPropTagArray /*lpPropTagArray*/,
		ULONG /*ulFlags*/,
		ULONG FAR* /*lpcValues*/,
		LPSPropValue FAR* /*lppPropArray*/)
	{
		// TODO: See if accounts could support this
		// This is only called from the Extra Props code. We can't support Extra Props from a row
		// so we don't need to implement this.
		return E_NOTIMPL;
	}

	_Check_return_ HRESULT accountPropertyBag::GetProp(ULONG ulPropTag, LPSPropValue FAR* lppProp)
	{
		if (!lppProp) return MAPI_E_INVALID_PARAMETER;

		//*lppProp = PpropFindProp(m_lpProps, m_cValues, ulPropTag);
		return S_OK;
	}

	_Check_return_ HRESULT accountPropertyBag::SetProp(LPSPropValue lpProp)
	{
		return E_NOTIMPL;
		//ULONG ulNewArray = NULL;
		//LPSPropValue lpNewArray = nullptr;

		//const auto hRes = EC_H(ConcatLPSPropValue(1, lpProp, m_cValues, m_lpProps, &ulNewArray, &lpNewArray));
		//if (SUCCEEDED(hRes))
		//{
		//	MAPIFreeBuffer(m_lpListData->lpSourceProps);
		//	m_lpListData->cSourceProps = ulNewArray;
		//	m_lpListData->lpSourceProps = lpNewArray;

		//	m_cValues = ulNewArray;
		//	m_lpProps = lpNewArray;
		//	m_bRowModified = true;
		//}
		//return hRes;
	}
} // namespace propertybag
