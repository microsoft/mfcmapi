#include <core/stdafx.h>
#include <core/propertyBag/accountPropertyBag.h>
#include <core/mapi/mapiFunctions.h>
#include <core/mapi/mapiMemory.h>

namespace propertybag
{
	accountPropertyBag::accountPropertyBag(_In_ std::wstring lpwszProfile, _In_ LPOLKACCOUNT lpAccount)
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
		const auto ulFlags = propBagFlags::None;
		return ulFlags;
	}

	bool accountPropertyBag::IsEqual(_In_ const std::shared_ptr<IMAPIPropertyBag> lpPropBag) const
	{
		if (!lpPropBag) return false;
		if (GetType() != lpPropBag->GetType()) return false;

		// Two accounts are the same if their profile is the same
		// And PROP_ACCT_MINI_UID is the same
		const auto lpOther = std::dynamic_pointer_cast<accountPropertyBag>(lpPropBag);
		if (lpOther)
		{
			if (m_lpwszProfile != lpOther->m_lpwszProfile) return false;
			if (!lpOther->m_lpAccount && !m_lpAccount) return true; // If neither has an account they're the same
			if (!lpOther->m_lpAccount || !m_lpAccount) return false; // If just one has an account they're different
			auto uidMe = ACCT_VARIANT{};
			auto uidOther = ACCT_VARIANT{};
			auto hRes = WC_H(m_lpAccount->GetProp(PROP_ACCT_MINI_UID, &uidMe));
			if (FAILED(hRes)) return false;
			hRes = WC_H(lpOther->m_lpAccount->GetProp(PROP_ACCT_MINI_UID, &uidOther));
			if (FAILED(hRes)) return false;

			if (uidMe.dwType != PT_LONG || uidOther.dwType != PT_LONG) return false;
			if (uidMe.Val.dw != uidOther.Val.dw) return false;
			return true;
		}

		return false;
	}

	// Convert an ACCT_VARIANT to SPropValue allocated off of pParent
	// Frees any memory associated with the ACCT_VARIANT
	SPropValue accountPropertyBag::convertVarToMAPI(
		_In_ ULONG ulPropTag,
		_In_ const ACCT_VARIANT& var,
		_In_opt_ const VOID* pParent)
	{
		auto sProp = SPropValue{ulPropTag, 0};
		switch (var.dwType)
		{
		case PT_LONG:
			sProp.Value.l = var.Val.dw;
			break;
		case PT_UNICODE:
			sProp.Value.lpszW = mapi::CopyStringW(var.Val.pwsz, pParent);
			WC_H_S(m_lpAccount->FreeMemory(reinterpret_cast<LPBYTE>(var.Val.pwsz)));
			break;
		case PT_BINARY:
			auto bin = SBinary{var.Val.bin.cb, var.Val.bin.pb};
			sProp.Value.bin = mapi::CopySBinary(bin, pParent);
			WC_H_S(m_lpAccount->FreeMemory(reinterpret_cast<LPBYTE>(var.Val.bin.pb)));
			break;
		}

		return sProp;
	}

	// Always returns a propval, even in errors, unless we fail allocating memory
	_Check_return_ LPSPropValue accountPropertyBag::GetOneProp(_In_ ULONG ulPropTag)
	{
		LPSPropValue lpPropVal = mapi::allocate<LPSPropValue>(sizeof(SPropValue));
		if (lpPropVal)
		{
			auto pProp = ACCT_VARIANT{};
			const auto hRes = WC_H(m_lpAccount->GetProp(ulPropTag, &pProp));
			if (SUCCEEDED(hRes))
			{
				*lpPropVal = convertVarToMAPI(ulPropTag, pProp, lpPropVal);
			}
			else
			{
				lpPropVal->ulPropTag = CHANGE_PROP_TYPE(ulPropTag, PT_ERROR);
				lpPropVal->Value.err = hRes;
			}
		}

		return lpPropVal;
	}

	_Check_return_ HRESULT accountPropertyBag::SetProp(_In_ LPSPropValue lpProp, _In_ bool saveChanges)
	{
		if (!lpProp) return MAPI_E_INVALID_PARAMETER;

		auto pProp = ACCT_VARIANT{PROP_TYPE(lpProp->ulPropTag), 0};

		switch (pProp.dwType)
		{
		case PT_LONG:
			pProp.Val.dw = lpProp->Value.l;
			break;
		case PT_UNICODE:
			pProp.Val.pwsz = lpProp->Value.lpszW;
			break;
		case PT_BINARY:
			pProp.Val.bin = {lpProp->Value.bin.cb, lpProp->Value.bin.lpb};
			break;
		}

		auto hRes = WC_H(m_lpAccount->SetProp(lpProp->ulPropTag, &pProp));
		if (SUCCEEDED(hRes) && saveChanges)
		{
			hRes = WC_H(m_lpAccount->SaveChanges(OLK_ACCOUNT_NO_FLAGS));
		}

		return hRes;
	}

	_Check_return_ HRESULT accountPropertyBag::SetProps(_In_ ULONG cValues, _In_ LPSPropValue lpPropArray)
	{
		if (!cValues || !lpPropArray) return MAPI_E_INVALID_PARAMETER;

		for (ULONG i = 0; i < cValues; i++)
		{
			const auto hRes = WC_H(SetProp(&lpPropArray[i], false));
			if (FAILED(hRes)) return hRes;
		}

		return WC_H(m_lpAccount->SaveChanges(OLK_ACCOUNT_NO_FLAGS));
	}

	_Check_return_ HRESULT accountPropertyBag::SetProp(_In_ LPSPropValue lpProp) { return SetProp(lpProp, true); }

	_Check_return_ std::vector<std::shared_ptr<model::mapiRowModel>> accountPropertyBag::GetAllModels()
	{
		if (!m_lpAccount)
		{
			return {};
		}

		auto hRes = S_OK;
		std::vector<std::pair<ULONG, ACCT_VARIANT>> props = {};
		const auto _ignore = std::list<HRESULT>{static_cast<HRESULT>(E_ACCT_NOT_FOUND)};
		for (auto i = 0; i < 0x8000; i++)
		{
			auto pProp = ACCT_VARIANT{};
			hRes = WC_H_IGNORE_RET(_ignore, m_lpAccount->GetProp(PROP_TAG(PT_LONG, i), &pProp));
			if (SUCCEEDED(hRes))
			{
				props.emplace_back(PROP_TAG(PT_LONG, i), pProp);
			}

			hRes = WC_H_IGNORE_RET(_ignore, m_lpAccount->GetProp(PROP_TAG(PT_UNICODE, i), &pProp));
			if (SUCCEEDED(hRes))
			{
				props.emplace_back(PROP_TAG(PT_UNICODE, i), pProp);
			}

			hRes = WC_H_IGNORE_RET(_ignore, m_lpAccount->GetProp(PROP_TAG(PT_BINARY, i), &pProp));
			if (SUCCEEDED(hRes))
			{
				props.emplace_back(PROP_TAG(PT_BINARY, i), pProp);
			}
		}

		if (props.size() > 0)
		{
			auto models = std::vector<std::shared_ptr<model::mapiRowModel>>{};

			for (const auto& prop : props)
			{
				models.push_back(convertVarToModel(prop.second, prop.first));
			}

			return models;
		}

		return {};
	}

	_Check_return_ std::shared_ptr<model::mapiRowModel> accountPropertyBag::GetOneModel(_In_ ULONG ulPropTag)
	{
		auto pProp = ACCT_VARIANT{};
		const auto hRes = WC_H(m_lpAccount->GetProp(ulPropTag, &pProp));
		if (SUCCEEDED(hRes))
		{
			return convertVarToModel(pProp, ulPropTag);
		}
		else
		{
			auto propVal = SPropValue{CHANGE_PROP_TYPE(ulPropTag, PT_ERROR), 0};
			propVal.Value.err = hRes;
			return model::propToModel(&propVal, ulPropTag, nullptr, false);
		}
	}

	_Check_return_ std::shared_ptr<model::mapiRowModel>
	accountPropertyBag::convertVarToModel(_In_ const ACCT_VARIANT& var, _In_ ULONG ulPropTag)
	{
		auto sProp = SPropValue{ulPropTag, 0};
		switch (var.dwType)
		{
		case PT_LONG:
			sProp.Value.l = var.Val.dw;
			break;
		case PT_UNICODE:
			sProp.Value.lpszW = var.Val.pwsz;
			break;
		case PT_BINARY:
			sProp.Value.bin = SBinary{var.Val.bin.cb, var.Val.bin.pb};
			break;
		}

		const auto model = model::propToModel(&sProp, ulPropTag, nullptr, false);
		if (var.dwType == PT_UNICODE)
		{
			WC_H_S(m_lpAccount->FreeMemory(reinterpret_cast<LPBYTE>(var.Val.pwsz)));
		}
		else if (var.dwType == PT_BINARY)
		{
			WC_H_S(m_lpAccount->FreeMemory(reinterpret_cast<LPBYTE>(var.Val.bin.pb)));
		}

		return model;
	}
} // namespace propertybag
