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

		_Check_return_ std::shared_ptr<model::mapiRowModel> toModel();

	private:
		std::wstring m_name{};
		ULONG m_ulPropTag{};
		SPropValue m_prop{};
		bool m_secure{};
		bool m_canParseMAPI{};
		DWORD m_dwType{};
		DWORD m_dwVal{};
		std::wstring m_szVal{};
		std::vector<BYTE> m_binVal{};
		std::vector<BYTE> m_unicodeVal{}; // in case we need to modify the bin to aid parsing
	};
} // namespace propertybag