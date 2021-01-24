#pragma once
#include <core/model/mapiRowModel.h>

namespace propertybag
{
	enum class propBagFlags
	{
		None = 0x0000, // None
		Modified = 0x0001, // The property bag has been modified
		BackedByGetProps = 0x0002, // The property bag is rendering from a GetProps call
		AB = 0x0004, // The property bag represents an Address Book object
		CanDelete = 0x0008, // The property bag supports deletion
	};
	DEFINE_ENUM_FLAG_OPERATORS(propBagFlags)

	enum class propBagType
	{
		MAPIProp,
		Row,
		Account,
		Registry
	};

	// TODO - Annotate for sal
	class IMAPIPropertyBag
	{
	public:
		virtual ~IMAPIPropertyBag() = default;

		virtual propBagType GetType() const = 0;
		virtual bool IsEqual(_In_ const std::shared_ptr<IMAPIPropertyBag> lpPropBag) const = 0;

		virtual _Check_return_ LPMAPIPROP GetMAPIProp() const = 0;

		// TODO: Should Commit take flags?
		virtual _Check_return_ HRESULT Commit() = 0;
		// MAPI Style property getter
		// Data will be freed (if needed) via FreeBuffer
		virtual _Check_return_ LPSPropValue GetOneProp(ULONG ulPropTag) = 0;
		virtual _Check_return_ LPSPropValue GetOneProp(const std::wstring& /*name*/) { return nullptr; }
		virtual void FreeBuffer(LPSPropValue lpProp) = 0;
		virtual _Check_return_ HRESULT SetProps(_In_ ULONG cValues, _In_ LPSPropValue lpPropArray) = 0;
		virtual _Check_return_ HRESULT SetProp(_In_ LPSPropValue lpProp) = 0;
		virtual _Check_return_ HRESULT DeleteProp(_In_ ULONG ulPropTag) = 0;
		bool IsAB() { return (GetFlags() & propBagFlags::AB) == propBagFlags::AB; }
		bool IsModified() { return (GetFlags() & propBagFlags::Modified) == propBagFlags::Modified; }
		bool IsBackedByGetProps()
		{
			return (GetFlags() & propBagFlags::BackedByGetProps) == propBagFlags::BackedByGetProps;
		}
		bool CanDelete() { return (GetFlags() & propBagFlags::CanDelete) == propBagFlags::CanDelete; }

		// Model implementation
		virtual _Check_return_ std::vector<std::shared_ptr<model::mapiRowModel>> GetAllModels() = 0;
		virtual _Check_return_ std::shared_ptr<model::mapiRowModel> GetOneModel(_In_ ULONG ulPropTag) = 0;

	protected:
		virtual propBagFlags GetFlags() const = 0;
	};
} // namespace propertybag