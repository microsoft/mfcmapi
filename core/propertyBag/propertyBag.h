#pragma once

namespace propertybag
{
	enum class propBagFlags
	{
		None = 0x0000, // None
		Modified = 0x0001, // The property bag has been modified
		BackedByGetProps = 0x0002, // The property bag is rendering from a GetProps call
	};
	DEFINE_ENUM_FLAG_OPERATORS(propBagFlags)

	enum class propBagType
	{
		MAPIProp,
		Row,
	};

	// TODO - Annotate for sal
	class IMAPIPropertyBag
	{
	public:
		virtual ~IMAPIPropertyBag() = default;

		virtual propBagFlags GetFlags() const = 0;
		virtual propBagType GetType() const = 0;
		virtual bool IsEqual(const std::shared_ptr<IMAPIPropertyBag> lpPropBag) const = 0;

		virtual _Check_return_ LPMAPIPROP GetMAPIProp() const = 0;

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
} // namespace propertybag