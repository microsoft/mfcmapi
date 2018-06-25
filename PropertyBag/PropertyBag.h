#pragma once

namespace propertybag
{
	enum propBagFlags
	{
		pbNone = 0x0000, // None
		pbModified = 0x0001, // The property bag has been modified
		pbBackedByGetProps = 0x0002, // The property bag is rendering from a GetProps call
	};

	enum propBagType
	{
		pbMAPIProp,
		pbRow,
	};

	class IMAPIPropertyBag;
	typedef IMAPIPropertyBag FAR* LPMAPIPROPERTYBAG;

	// TODO - Annotate for sal
	class IMAPIPropertyBag
	{
	public:
		virtual ~IMAPIPropertyBag() = default;

		virtual ULONG GetFlags() = 0;
		virtual propBagType GetType() = 0;
		virtual bool IsEqual(LPMAPIPROPERTYBAG lpPropBag) = 0;

		virtual _Check_return_ LPMAPIPROP GetMAPIProp() = 0;

		// TODO: Should Commit take flags?
		virtual _Check_return_ HRESULT Commit() = 0;
		virtual _Check_return_ HRESULT GetAllProps(ULONG FAR* lpcValues, LPSPropValue FAR* lppPropArray) = 0;
		virtual _Check_return_ HRESULT GetProps(
			LPSPropTagArray lpPropTagArray,
			ULONG ulFlags,
			ULONG FAR* lpcValues,
			LPSPropValue FAR* lppPropArray) = 0;
		virtual _Check_return_ HRESULT GetProp(ULONG ulPropTag, LPSPropValue FAR* lppProp) = 0;
		virtual void FreeBuffer(LPSPropValue lpProp) = 0;
		virtual _Check_return_ HRESULT SetProps(ULONG cValues, LPSPropValue lpPropArray) = 0;
		virtual _Check_return_ HRESULT SetProp(LPSPropValue lpProp) = 0;
		virtual _Check_return_ HRESULT DeleteProp(ULONG ulPropTag) = 0;
	};
}