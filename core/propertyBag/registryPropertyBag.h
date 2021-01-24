#pragma once
#include <core/propertyBag/propertyBag.h>
#include <core/propertyBag/registryProperty.h>

namespace propertybag
{
	class registryPropertyBag : public IMAPIPropertyBag
	{
	public:
		registryPropertyBag(HKEY hKey);
		virtual ~registryPropertyBag();
		registryPropertyBag(const registryPropertyBag&) = delete;
		registryPropertyBag& operator=(const registryPropertyBag&) = delete;

		propBagType GetType() const override { return propBagType::Registry; }
		bool IsEqual(const std::shared_ptr<IMAPIPropertyBag> lpPropBag) const override;

		// Returns the underlying MAPI prop object, if one exists. Does NOT ref count it.
		_Check_return_ LPMAPIPROP GetMAPIProp() const override { return nullptr; }

		_Check_return_ HRESULT Commit() override { return E_NOTIMPL; }
		_Check_return_ LPSPropValue GetOneProp(ULONG ulPropTag, const std::wstring& name) override;
		void FreeBuffer(LPSPropValue lpsPropValue) override { MAPIFreeBuffer(lpsPropValue); }
		_Check_return_ HRESULT SetProps(ULONG cValues, LPSPropValue lpPropArray) override;
		_Check_return_ HRESULT SetProp(LPSPropValue lpProp) override;
		_Check_return_ HRESULT DeleteProp(ULONG /*ulPropTag*/) noexcept override { return E_NOTIMPL; };
		_Check_return_ std::vector<std::shared_ptr<model::mapiRowModel>> GetAllModels() override;
		_Check_return_ std::shared_ptr<model::mapiRowModel> GetOneModel(_In_ ULONG ulPropTag) override;

	protected:
		propBagFlags GetFlags() const override;
		void GetAllProps();

	private:
		HKEY m_hKey{};
		std::vector<std::shared_ptr<registryProperty>> m_props{};
	};
} // namespace propertybag