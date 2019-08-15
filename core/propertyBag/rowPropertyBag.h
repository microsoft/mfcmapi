#pragma once
#include <core/propertyBag/propertyBag.h>
#include <core/sortlistdata/sortListData.h>

namespace propertybag
{
	class RowPropertyBag : public IMAPIPropertyBag
	{
	public:
		RowPropertyBag(sortlistdata::sortListData* lpListData);
		virtual ~RowPropertyBag() = default;

		ULONG GetFlags() const override;
		propBagType GetType() const override { return pbRow; }
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
		// None of our GetProps allocate anything, so nothing to do here
		void RowPropertyBag::FreeBuffer(LPSPropValue) override { return; }
		// TODO: This is for paste, something we don't yet support for rows
		_Check_return_ HRESULT SetProps(ULONG, LPSPropValue) override { return E_NOTIMPL; }
		_Check_return_ HRESULT SetProp(LPSPropValue lpProp) override;
		//TODO: Not supported yet
		_Check_return_ HRESULT DeleteProp(ULONG) override { return E_NOTIMPL; };

	private:
		sortlistdata::sortListData* m_lpListData{};

		ULONG m_cValues{};
		LPSPropValue m_lpProps{};
		bool m_bRowModified{};
	};
} // namespace propertybag