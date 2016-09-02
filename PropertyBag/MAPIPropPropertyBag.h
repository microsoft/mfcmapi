#pragma once
// MAPIPropPropertyBag.h : header file

#include "PropertyBag.h"
#include "SortList/SortListData.h"

class MAPIPropPropertyBag : public IMAPIPropertyBag
{
public:
	MAPIPropPropertyBag(LPMAPIPROP lpProp, SortListData* lpListData);
	virtual ~MAPIPropPropertyBag();

	virtual ULONG GetFlags();
	virtual propBagType GetType();
	virtual bool IsEqual(LPMAPIPROPERTYBAG lpPropBag);

	virtual _Check_return_ LPMAPIPROP GetMAPIProp();

	virtual _Check_return_ HRESULT Commit();
	virtual _Check_return_ HRESULT GetAllProps(
		ULONG FAR* lpcValues,
		LPSPropValue FAR* lppPropArray);
	virtual _Check_return_ HRESULT GetProps(
		LPSPropTagArray lpPropTagArray,
		ULONG ulFlags,
		ULONG FAR* lpcValues,
		LPSPropValue FAR* lppPropArray);
	virtual _Check_return_ HRESULT GetProp(
		ULONG ulPropTag,
		LPSPropValue FAR* lppProp);
	virtual void FreeBuffer(LPSPropValue lpProp);
	virtual _Check_return_ HRESULT SetProps(
		ULONG cValues,
		LPSPropValue lpPropArray);
	virtual _Check_return_ HRESULT SetProp(
		LPSPropValue lpProp);
	virtual _Check_return_ HRESULT DeleteProp(
		ULONG ulPropTag);

private:
	SortListData* m_lpListData;
	LPMAPIPROP m_lpProp;
	bool m_bGetPropsSucceeded;
};