#include <core/stdafx.h>
#include <core/propertyBag/registryPropertyBag.h>
#include <core/mapi/mapiFunctions.h>
#include <core/mapi/mapiMemory.h>
#include <core/utility/strings.h>

namespace propertybag
{
	registryPropertyBag::registryPropertyBag(HKEY hKey) { m_hKey = hKey; }

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
			// Can two different hKeys be the same reg key?
			if (m_hKey != lpOther->m_hKey) return false;
			return true;
		}

		return false;
	}

	_Check_return_ HRESULT registryPropertyBag::GetAllProps(ULONG FAR* lpcValues, LPSPropValue FAR* lppPropArray)
	{
		if (!lpcValues || !lppPropArray) return MAPI_E_INVALID_PARAMETER;
		*lpcValues = 0;
		*lppPropArray = nullptr;

		auto cchMaxValueNameLen = DWORD{}; // Param in RegQueryInfoKeyW is misnamed
		auto cValues = DWORD{};
		auto hRes = WC_W32(RegQueryInfoKeyW(
			m_hKey,
			nullptr, // lpClass
			nullptr, // lpcchClass
			nullptr, // lpReserved
			nullptr, // lpcSubKeys
			nullptr, // lpcbMaxSubKeyLen
			nullptr, // lpcbMaxClassLen
			&cValues, // lpcValues
			&cchMaxValueNameLen, // lpcbMaxValueNameLen
			nullptr, // lpcbMaxValueLen
			nullptr, // lpcbSecurityDescriptor
			nullptr)); // lpftLastWriteTime

		if (cValues && cchMaxValueNameLen)
		{
			*lpcValues = cValues;
			*lppPropArray = mapi::allocate<LPSPropValue>(cValues * sizeof(SPropValue));

			cchMaxValueNameLen++; // For null terminator
			auto szBuf = std::wstring(cchMaxValueNameLen, '\0');
			for (DWORD dwIndex = 0; dwIndex < cValues; dwIndex++)
			{
				auto cchValLen = cchMaxValueNameLen;
				szBuf.clear();
				hRes = WC_W32(RegEnumValueW(
					m_hKey,
					dwIndex,
					const_cast<wchar_t*>(szBuf.c_str()),
					&cchValLen,
					nullptr, // lpReserved
					nullptr, // lpType
					nullptr, // lpData
					nullptr)); // lpcbData
				if (hRes == S_OK)
				{
					auto prop = &(*lppPropArray)[dwIndex];
					prop->ulPropTag = PROP_TAG(PT_UNICODE, dwIndex);
					const auto szVal = strings::format(L"%ws", szBuf.c_str());
					prop->Value.lpszW = mapi::CopyStringW(szVal.c_str(), *lppPropArray);

					// TODO: Identify prop tags in the value name
					// Use type from tag to determine how to read data
					// Also use lpType
					// Convert data read to a MAPI prop Value
					// Figure out way to deal with S props
					// Figure out way to deal with named props
				}
			}
		}

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
