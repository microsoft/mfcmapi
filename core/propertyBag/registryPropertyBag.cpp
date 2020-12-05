#include <core/stdafx.h>
#include <core/propertyBag/registryPropertyBag.h>
#include <core/mapi/mapiFunctions.h>
#include <core/mapi/mapiMemory.h>

namespace propertybag
{
	registryPropertyBag::registryPropertyBag(std::wstring /*lpwszProfile*/, LPOLKACCOUNT /*lpAccount*/) {}

	registryPropertyBag ::~registryPropertyBag() {}

	propBagFlags registryPropertyBag::GetFlags() const
	{
		const auto ulFlags = propBagFlags::None;
		return ulFlags;
	}

	bool registryPropertyBag::IsEqual(const std::shared_ptr<IMAPIPropertyBag> lpPropBag) const
	{
		if (!lpPropBag) return false;
		if (GetType() != lpPropBag->GetType()) return false;

		const auto lpOther = std::dynamic_pointer_cast<registryPropertyBag>(lpPropBag);
		if (lpOther)
		{
			//if (m_lpwszProfile != lpOther->m_lpwszProfile) return false;
			return true;
		}

		return false;
	}

	_Check_return_ HRESULT registryPropertyBag::GetAllProps(ULONG FAR* lpcValues, LPSPropValue FAR* lppPropArray)
	{
		if (!lpcValues || !lppPropArray) return MAPI_E_INVALID_PARAMETER;

		return E_NOTIMPL;
	}

	_Check_return_ HRESULT registryPropertyBag::GetProps(
		LPSPropTagArray lpPropTagArray,
		ULONG /*ulFlags*/,
		ULONG FAR* lpcValues,
		LPSPropValue FAR* lppPropArray)
	{
		if (!lpcValues || !lppPropArray) return MAPI_E_INVALID_PARAMETER;

		return E_NOTIMPL;
	}

	_Check_return_ HRESULT registryPropertyBag::GetProp(ULONG ulPropTag, LPSPropValue FAR* lppProp)
	{
		if (!lppProp) return MAPI_E_INVALID_PARAMETER;

		return E_NOTIMPL;
	}

	_Check_return_ HRESULT registryPropertyBag::SetProps(ULONG cValues, LPSPropValue lpPropArray)
	{
		if (!cValues || !lpPropArray) return MAPI_E_INVALID_PARAMETER;

		return E_NOTIMPL;
	}

	_Check_return_ HRESULT registryPropertyBag::SetProp(LPSPropValue lpProp) { return E_NOTIMPL; }
} // namespace propertybag
