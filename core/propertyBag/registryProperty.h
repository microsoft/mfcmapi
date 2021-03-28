#pragma once
#include <core/propertyBag/propertyBag.h>

namespace propertybag
{
	class registryProperty
	{
	public:
		registryProperty(_In_ const HKEY hKey, _In_ const std::wstring& name, _In_ DWORD dwType, _In_ ULONG ulPropTag);
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

		std::wstring name() { return m_name; }
		ULONG ulPropTag() noexcept { return m_ulPropTag; }
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
		std::vector<BYTE> m_binVal{}; // raw data read from registry
		std::vector<BYTE> m_unicodeVal{}; // in case we need to modify the bin to aid parsing
		std::string m_ansiVal{}; // For PT_STRING* in REG_SZ

		// Temp storage for m_sOutputValue
		std::vector<BYTE> m_bin; // backing data for base MV structs
		std::vector<std::string> m_mvA; // Backing data for MV ANSI array
		std::vector<std::wstring> m_mvW; // Backing data for MV Unicode array
		std::vector<std::vector<BYTE>> m_mvBin; // Backing data for MV binary array
	};
} // namespace propertybag