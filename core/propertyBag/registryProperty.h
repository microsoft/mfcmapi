#pragma once
#include <core/propertyBag/propertyBag.h>

namespace propertybag
{
	class registryProperty
	{
	public:
		registryProperty(
			_In_ const std::wstring& name,
			DWORD dwType,
			DWORD dwVal,
			_In_ const std::wstring& szVal,
			_In_ const std::vector<BYTE>& binVal)
			: m_name(name), m_dwType(dwType), m_dwVal(dwVal), m_szVal(szVal), m_binVal(binVal)
		{
		}

		registryProperty(const registryProperty&) = delete;
		registryProperty& operator=(const registryProperty&) = delete;

		_Check_return_ std::shared_ptr<model::mapiRowModel> toModel();

	private:
		std::wstring m_name{};
		DWORD m_dwType{};
		DWORD m_dwVal{};
		std::wstring m_szVal{};
		std::vector<BYTE> m_binVal{};
	};
} // namespace propertybag