#pragma once
#include <core/propertyBag/propertyBag.h>
#include <core/sortlistdata/sortListData.h>

namespace propertybag
{
	class mapiPropPropertyBag : public IMAPIPropertyBag
	{
	public:
		mapiPropPropertyBag(LPMAPIPROP lpProp, sortlistdata::sortListData* lpListData);
		virtual ~mapiPropPropertyBag();

		ULONG GetFlags() const override;
		propBagType GetType() const { return pbMAPIProp; }
		bool IsEqual(const std::shared_ptr<IMAPIPropertyBag> lpPropBag) const override;

		// Returns the underlying MAPI prop object, if one exists. Does NOT ref count it.
		_Check_return_ LPMAPIPROP GetMAPIProp() const { return m_lpProp; }

		_Check_return_ HRESULT Commit() override;
		_Check_return_ HRESULT GetAllProps(ULONG FAR* lpcValues, LPSPropValue FAR* lppPropArray) override;
		_Check_return_ HRESULT GetProps(
			LPSPropTagArray lpPropTagArray,
			ULONG ulFlags,
			ULONG FAR* lpcValues,
			LPSPropValue FAR* lppPropArray) override;
		_Check_return_ HRESULT GetProp(ULONG ulPropTag, LPSPropValue FAR* lppProp) override;
		void FreeBuffer(LPSPropValue lpProp) override;
		_Check_return_ HRESULT SetProps(ULONG cValues, LPSPropValue lpPropArray) override;
		_Check_return_ HRESULT SetProp(LPSPropValue lpProp) override;
		_Check_return_ HRESULT DeleteProp(ULONG ulPropTag) override;

	private:
		sortlistdata::sortListData* m_lpListData{};
		LPMAPIPROP m_lpProp{};
		bool m_bGetPropsSucceeded{};
	};
} // namespace propertybag