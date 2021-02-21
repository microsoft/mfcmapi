#pragma once
#include <core/propertyBag/propertyBag.h>

namespace propertybag
{
	class registryProperty
	{
	public:
		registryProperty(HKEY hKey, _In_ const std::wstring& name, DWORD dwType);
		registryProperty(const registryProperty&) = delete;
		registryProperty& operator=(const registryProperty&) = delete;

		_Check_return_ LPSPropValue toSPropValue()
		{
			ensureSPropValue();
			return &m_prop;
		}

		_Check_return_ std::shared_ptr<model::mapiRowModel> toModel()
		{
			ensureModel();
			return m_model;
		}

		void set(_In_opt_ LPSPropValue newValue);

	private:
		void ensureSPropValue();
		void ensureModel();

		HKEY m_hKey{};
		std::wstring m_name{};
		ULONG m_ulPropTag{};
		SPropValue m_prop{};
		std::shared_ptr<model::mapiRowModel> m_model;
		bool m_secure{};
		bool m_canParseMAPI{};
		DWORD m_dwType{};
		DWORD m_dwVal{};
		std::wstring m_szVal{};
		std::vector<BYTE> m_binVal{};
		std::vector<BYTE> m_unicodeVal{}; // in case we need to modify the bin to aid parsing
		std::string m_ansiVal{}; // For PT_STRING* in REG_SZ
	};
} // namespace propertybag