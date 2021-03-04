#include <core/stdafx.h>
#include <core/propertyBag/registryPropertyBag.h>
#include <core/mapi/mapiFunctions.h>
#include <core/utility/strings.h>
#include <core/utility/registry.h>
#include <core/utility/error.h>
#include <core/interpret/proptags.h>
#include <core/property/parseProperty.h>
#include <core/smartview/SmartView.h>
#include <core/propertyBag/registryProperty.h>

namespace propertybag
{
	registryPropertyBag::registryPropertyBag(HKEY hKey) : m_hKey(hKey) {}

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

	void registryPropertyBag::load()
	{
		m_props = {};
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
			auto szBuf = std::wstring(cchMaxValueNameLen, L'\0');
			cchMaxValueNameLen++; // For null terminator
			for (DWORD dwIndex = 0; dwIndex < cValues; dwIndex++)
			{
				auto dwType = DWORD{};
				auto cchValLen = cchMaxValueNameLen;
				szBuf.clear();
				hRes = WC_W32(RegEnumValueW(
					m_hKey,
					dwIndex,
					const_cast<wchar_t*>(szBuf.c_str()),
					&cchValLen,
					nullptr, // lpReserved
					&dwType, // lpType
					nullptr, // lpData
					nullptr)); // lpcbData
				if (hRes == S_OK)
				{
					const auto valName =
						std::wstring(szBuf.c_str()); // szBuf.size() is 0, so make a copy with a proper size

					m_props.push_back(std::make_shared<registryProperty>(m_hKey, valName, dwType));
				}
			}
		}

		m_loaded = true;
	}

	_Check_return_ std::vector<std::shared_ptr<model::mapiRowModel>> registryPropertyBag::GetAllModels()
	{
		ensureLoaded(true);
		auto models = std::vector<std::shared_ptr<model::mapiRowModel>>{};
		for (const auto& prop : m_props)
		{
			models.push_back(prop->toModel());
		}

		return models;
	}

	_Check_return_ std::shared_ptr<model::mapiRowModel> registryPropertyBag::GetOneModel(_In_ ULONG ulPropTag)
	{
		ensureLoaded();
		for (const auto& prop : m_props)
		{
			if (prop->ulPropTag() == ulPropTag)
			{
				return prop->toModel();
			}
		}

		const auto keyName = strings::format(L"%04x%04x", PROP_TYPE(ulPropTag), PROP_ID(ulPropTag));
		auto prop = std::make_shared<registryProperty>(m_hKey, keyName, REG_NONE);
		m_props.push_back(prop);
		return prop->toModel();
	}

	_Check_return_ LPSPropValue registryPropertyBag::GetOneProp(ULONG ulPropTag, const std::wstring& name)
	{
		ensureLoaded();
		if (!name.empty())
		{
			for (const auto& prop : m_props)
			{
				if (prop->toModel()->name() == name)
				{
					return prop->toSPropValue();
				}
			}
		}

		for (const auto& prop : m_props)
		{
			if (prop->toModel()->ulPropTag() == ulPropTag) return prop->toSPropValue();
		}

		return {};
	}

	_Check_return_ HRESULT registryPropertyBag::SetProps(ULONG cValues, LPSPropValue lpPropArray)
	{
		if (!cValues || !lpPropArray) return MAPI_E_INVALID_PARAMETER;

		return E_NOTIMPL;
	}

	_Check_return_ HRESULT
	registryPropertyBag::SetProp(_In_ LPSPropValue lpProp, _In_ ULONG ulPropTag, const std::wstring& name)
	{
		ensureLoaded();
		for (const auto& prop : m_props)
		{
			if (prop->toModel()->name() == name)
			{
				prop->set(lpProp);
				return S_OK;
			}
		}

		const auto keyName =
			name.empty() ? strings::format(L"%04x%04x", PROP_TYPE(ulPropTag), PROP_ID(ulPropTag)) : name;
		auto prop = std::make_shared<registryProperty>(
			m_hKey, keyName, PROP_TYPE(ulPropTag) == PT_STRING8 ? REG_SZ : REG_BINARY);
		prop->set(lpProp);
		m_props.push_back(prop);
		return S_OK;
	}

	_Check_return_ HRESULT registryPropertyBag::DeleteProp(_In_ ULONG /*ulPropTag*/, _In_ const std::wstring& name)
	{
		ensureLoaded();
		for (const auto& prop : m_props)
		{
			if (prop->toModel()->name() == name)
			{
				WC_W32_S(RegDeleteValueW(m_hKey, prop->name().c_str()));
				load(); // rather than try to rebuild our prop list, just reload
				return S_OK;
			}
		}

		return S_OK; // No need to error if the prop didn't exist.
	};
} // namespace propertybag