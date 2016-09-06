#pragma once
// MAPIPropPropertyBag.h : header file

#include "PropertyBag.h"
#include "SortList/SortListData.h"

class MAPIPropPropertyBag : public IMAPIPropertyBag
{
public:
	MAPIPropPropertyBag(LPMAPIPROP lpProp, SortListData* lpListData);
	virtual ~MAPIPropPropertyBag();

	ULONG GetFlags() override;
	propBagType GetType() override;
	bool IsEqual(LPMAPIPROPERTYBAG lpPropBag) override;

	_Check_return_ LPMAPIPROP GetMAPIProp() override;

	_Check_return_ HRESULT Commit() override;
	_Check_return_ HRESULT GetAllProps(
		ULONG FAR* lpcValues,
		LPSPropValue FAR* lppPropArray) override;
	_Check_return_ HRESULT GetProps(
		LPSPropTagArray lpPropTagArray,
		ULONG ulFlags,
		ULONG FAR* lpcValues,
		LPSPropValue FAR* lppPropArray) override;
	_Check_return_ HRESULT GetProp(
		ULONG ulPropTag,
		LPSPropValue FAR* lppProp) override;
	void FreeBuffer(LPSPropValue lpProp) override;
	_Check_return_ HRESULT SetProps(
		ULONG cValues,
		LPSPropValue lpPropArray) override;
	_Check_return_ HRESULT SetProp(
		LPSPropValue lpProp) override;
	_Check_return_ HRESULT DeleteProp(
		ULONG ulPropTag) override;

private:
	SortListData* m_lpListData;
	LPMAPIPROP m_lpProp;
	bool m_bGetPropsSucceeded;
};