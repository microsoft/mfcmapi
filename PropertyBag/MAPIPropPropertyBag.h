#pragma once
#include <PropertyBag/PropertyBag.h>
#include <UI/Controls/SortList/SortListData.h>

namespace propertybag
{
	class MAPIPropPropertyBag : public IMAPIPropertyBag
	{
	public:
		MAPIPropPropertyBag(LPMAPIPROP lpProp, controls::sortlistdata::SortListData* lpListData);
		virtual ~MAPIPropPropertyBag();

		ULONG GetFlags() const override;
		propBagType GetType() const { return pbMAPIProp; }
		bool IsEqual(const IMAPIPropertyBag* lpPropBag) const override;

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
		controls::sortlistdata::SortListData* m_lpListData;
		LPMAPIPROP m_lpProp;
		bool m_bGetPropsSucceeded;
	};
} // namespace propertybag