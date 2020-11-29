#pragma once
#include <core/propertyBag/propertyBag.h>
#include <core/mapi/account/actMgmt.h>

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
		_Check_return_ HRESULT GetAllProps(ULONG FAR* lpcValues, LPSPropValue FAR* lppPropArray) override;
		_Check_return_ HRESULT GetProps(
			LPSPropTagArray lpPropTagArray,
			ULONG ulFlags,
			ULONG FAR* lpcValues,
			LPSPropValue FAR* lppPropArray) override;
		_Check_return_ HRESULT GetProp(ULONG ulPropTag, LPSPropValue FAR* lppProp) override;
		void FreeBuffer(LPSPropValue lpsPropValue) override { MAPIFreeBuffer(lpsPropValue); }
		// TODO: This is for paste, something we don't yet support for rows
		_Check_return_ HRESULT SetProps(ULONG, LPSPropValue) override { return E_NOTIMPL; }
		_Check_return_ HRESULT SetProp(LPSPropValue lpProp) override;
		//TODO: Not supported yet
		_Check_return_ HRESULT DeleteProp(ULONG) override { return E_NOTIMPL; };

	private:
		std::wstring m_lpwszProfile;
		LPOLKACCOUNT m_lpAccount;

		SPropValue convertVarToMAPI(ULONG ulPropTag, ACCT_VARIANT var, _In_opt_ const VOID* pParent);
	};
} // namespace propertybag