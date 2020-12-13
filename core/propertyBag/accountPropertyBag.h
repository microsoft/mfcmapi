#pragma once
#include <core/propertyBag/propertyBag.h>

namespace propertybag
{
	class accountPropertyBag : public IMAPIPropertyBag
	{
	public:
		accountPropertyBag(std::wstring lpwszProfile, LPOLKACCOUNT lpAccount);
		virtual ~accountPropertyBag();
		accountPropertyBag(const accountPropertyBag&) = delete;
		accountPropertyBag& operator=(const accountPropertyBag&) = delete;

		propBagFlags GetFlags() const override;
		propBagType GetType() const override { return propBagType::Account; }
		bool IsEqual(const std::shared_ptr<IMAPIPropertyBag> lpPropBag) const override;

		// Returns the underlying MAPI prop object, if one exists. Does NOT ref count it.
		_Check_return_ LPMAPIPROP GetMAPIProp() const override { return nullptr; }

		_Check_return_ HRESULT Commit() override { return E_NOTIMPL; }
		_Check_return_ LPSPropValue GetOneProp(ULONG ulPropTag) override;
		void FreeBuffer(LPSPropValue lpsPropValue) override { MAPIFreeBuffer(lpsPropValue); }
		_Check_return_ HRESULT SetProps(ULONG cValues, LPSPropValue lpPropArray) override;
		_Check_return_ HRESULT SetProp(LPSPropValue lpProp) override;
		_Check_return_ HRESULT DeleteProp(ULONG /*ulPropTag*/) noexcept override { return E_NOTIMPL; };

		_Check_return_ std::vector<std::shared_ptr<model::mapiRowModel>> GetAllModels() override;
		_Check_return_ std::shared_ptr<model::mapiRowModel> GetOneModel(ULONG ulPropTag) override;

	private:
		std::wstring m_lpwszProfile;
		LPOLKACCOUNT m_lpAccount;

		_Check_return_ std::shared_ptr<model::mapiRowModel> convertVarToModel(const ACCT_VARIANT& var, ULONG ulPropTag);
		SPropValue convertVarToMAPI(ULONG ulPropTag, const ACCT_VARIANT& var, _In_opt_ const VOID* pParent);
		_Check_return_ HRESULT SetProp(LPSPropValue lpProp, bool saveChanges);
	};
} // namespace propertybag