#pragma once
#include <core/propertyBag/propertyBag.h>
#include <core/propertyBag/registryProperty.h>

namespace propertybag
{
	class registryPropertyBag : public IMAPIPropertyBag
	{
	public:
		registryPropertyBag(HKEY hKey);
		virtual ~registryPropertyBag() {}
		registryPropertyBag(const registryPropertyBag&) = delete;
		registryPropertyBag& operator=(const registryPropertyBag&) = delete;

		propBagType GetType() const override { return propBagType::Registry; }
		bool IsEqual(const std::shared_ptr<IMAPIPropertyBag> lpPropBag) const override;

		// Returns the underlying MAPI prop object, if one exists. Does NOT ref count it.
		_Check_return_ LPMAPIPROP GetMAPIProp() const override { return nullptr; }

		_Check_return_ HRESULT Commit() override { return S_OK; }
		_Check_return_ LPSPropValue GetOneProp(ULONG ulPropTag, const std::wstring& name) override;
		void FreeBuffer(LPSPropValue) override {}
		_Check_return_ HRESULT SetProps(ULONG cValues, LPSPropValue lpPropArray) override;
		_Check_return_ HRESULT
		SetProp(_In_ LPSPropValue lpProp, _In_ ULONG ulPropTag, const std::wstring& name) override;
		_Check_return_ HRESULT DeleteProp(_In_ ULONG ulPropTag, _In_ const std::wstring& name) override;
		_Check_return_ std::vector<std::shared_ptr<model::mapiRowModel>> GetAllModels() override;
		_Check_return_ std::shared_ptr<model::mapiRowModel> GetOneModel(_In_ ULONG ulPropTag) override;

	protected:
		propBagFlags GetFlags() const override { return propBagFlags::None; }

	private:
		void load();
		void ensureLoaded(bool force = false)
		{
			if (!m_loaded || force) load();
		}
		bool m_loaded{};
		HKEY m_hKey{};
		std::vector<std::shared_ptr<registryProperty>> m_props{};
	};
} // namespace propertybag