#include <core/stdafx.h>
#include <core/propertyBag/accountPropertyBag.h>
#include <core/mapi/mapiFunctions.h>
#include <core/mapi/mapiMemory.h>
#include <core/mapi/account/actMgmt.h>

namespace propertybag
{
	accountPropertyBag::accountPropertyBag(std::wstring lpwszProfile, LPOLKACCOUNT lpAccount)
	{
		m_lpwszProfile = lpwszProfile;
		m_lpAccount = mapi::safe_cast<LPOLKACCOUNT>(lpAccount);
	}

	accountPropertyBag ::~accountPropertyBag()
	{
		if (m_lpAccount) m_lpAccount->Release();
		m_lpAccount = nullptr;
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

	// Convert an ACCT_VARIANT to SPropValue allocated off of pParent
	// Frees any memory associated with the ACCT_VARIANT
	SPropValue accountPropertyBag::convertVarToMAPI(ULONG ulPropTag, ACCT_VARIANT var, _In_opt_ const VOID* pParent)
	{
		SPropValue sProp = {};
		sProp.ulPropTag = ulPropTag;
		switch (PROP_TYPE(ulPropTag))
		{
		case PT_LONG:
			sProp.Value.l = var.Val.dw;
			break;
		case PT_UNICODE:
			sProp.Value.lpszW = mapi::CopyStringW(var.Val.pwsz, pParent);
			(void) m_lpAccount->FreeMemory(reinterpret_cast<LPBYTE>(var.Val.pwsz));
			break;
		case PT_BINARY:
			auto bin = SBinary{var.Val.bin.cb, var.Val.bin.pb};
			sProp.Value.bin = mapi::CopySBinary(bin, pParent);
			(void) m_lpAccount->FreeMemory(reinterpret_cast<LPBYTE>(var.Val.bin.pb));
			break;
		}

		return sProp;
	}

	_Check_return_ HRESULT accountPropertyBag::GetAllProps(ULONG FAR* lpcValues, LPSPropValue FAR* lppPropArray)
	{
		if (!lpcValues || !lppPropArray) return MAPI_E_INVALID_PARAMETER;
		if (!m_lpAccount)
		{
			*lpcValues = 0;
			*lppPropArray = nullptr;
			return S_OK;
		}

		// Just return what we have
		auto hRes = S_OK;
		ULONG i = 0;
		std::vector<std::pair<ULONG, ACCT_VARIANT>> props = {};

		for (i = 0; i < 0x8000; i++)
		{
			ACCT_VARIANT pProp = {};
			hRes = m_lpAccount->GetProp(PROP_TAG(PT_LONG, i), &pProp);
			if (SUCCEEDED(hRes))
			{
				props.emplace_back(PROP_TAG(PT_LONG, i), pProp);
			}

			hRes = m_lpAccount->GetProp(PROP_TAG(PT_UNICODE, i), &pProp);
			if (SUCCEEDED(hRes))
			{
				props.emplace_back(PROP_TAG(PT_UNICODE, i), pProp);
			}

			hRes = m_lpAccount->GetProp(PROP_TAG(PT_BINARY, i), &pProp);
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
				(*lppPropArray)[iProp] = convertVarToMAPI(prop.first, prop.second, *lppPropArray);
				iProp++;
			}
		}

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

		auto pProp = ACCT_VARIANT{};
		const auto hRes = m_lpAccount->GetProp(ulPropTag, &pProp);
		if (SUCCEEDED(hRes))
		{
			*lppProp = mapi::allocate<LPSPropValue>(sizeof(SPropValue));
			(*lppProp)[0] = convertVarToMAPI(ulPropTag, pProp, *lppProp);
		}

		return S_OK;
	}

	_Check_return_ HRESULT accountPropertyBag::SetProp(LPSPropValue lpProp) { return E_NOTIMPL; }
} // namespace propertybag
