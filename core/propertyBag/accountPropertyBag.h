#pragma once
#include <core/propertyBag/propertyBag.h>

namespace propertybag
{
	class accountPropertyBag : public IMAPIPropertyBag
	{
	public:
		accountPropertyBag(_In_ std::wstring lpwszProfile, _In_ LPOLKACCOUNT lpAccount);
		virtual ~accountPropertyBag();
		accountPropertyBag(_In_ const accountPropertyBag&) = delete;
		accountPropertyBag& operator=(_In_ const accountPropertyBag&) = delete;

		propBagType GetType() const override { return propBagType::Account; }
		bool IsEqual(_In_ const std::shared_ptr<IMAPIPropertyBag> lpPropBag) const override;

		// Returns the underlying MAPI prop object, if one exists. Does NOT ref count it.
		_Check_return_ LPMAPIPROP GetMAPIProp() const override { return nullptr; }

		_Check_return_ HRESULT Commit() override { return E_NOTIMPL; }
		_Check_return_ LPSPropValue GetOneProp(_In_ ULONG ulPropTag) override;
		void FreeBuffer(_In_ LPSPropValue lpsPropValue) override { MAPIFreeBuffer(lpsPropValue); }
		_Check_return_ HRESULT SetProps(_In_ ULONG cValues, _In_ LPSPropValue lpPropArray) override;
		_Check_return_ HRESULT SetProp(_In_ LPSPropValue lpProp) override;
		_Check_return_ HRESULT DeleteProp(_In_ ULONG /*ulPropTag*/) noexcept override { return E_NOTIMPL; };

		_Check_return_ std::vector<std::shared_ptr<model::mapiRowModel>> GetAllModels() override;
		_Check_return_ std::shared_ptr<model::mapiRowModel> GetOneModel(_In_ ULONG ulPropTag) override;

	protected:
		propBagFlags GetFlags() const override;

	private:
		std::wstring m_lpwszProfile;
		LPOLKACCOUNT m_lpAccount;

		_Check_return_ std::shared_ptr<model::mapiRowModel>
		convertVarToModel(_In_ const ACCT_VARIANT& var, _In_ ULONG ulPropTag);
		SPropValue convertVarToMAPI(_In_ ULONG ulPropTag, _In_ const ACCT_VARIANT& var, _In_opt_ const VOID* pParent);
		_Check_return_ HRESULT SetProp(_In_ LPSPropValue lpProp, _In_ bool saveChanges);
	};
} // namespace propertybag