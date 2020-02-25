#pragma once
// Named Property Cache

namespace cache2
{
	struct NamePropNames;

	_Check_return_ std::vector<std::shared_ptr<NamedPropCacheEntry>> GetNamesFromIDs(
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ LPSPropTagArray* lppPropTags,
		_In_opt_ LPGUID lpPropSetGuid,
		ULONG ulFlags);
	_Check_return_ std::vector<std::shared_ptr<NamedPropCacheEntry>> GetNamesFromIDs(
		_In_ LPMAPIPROP lpMAPIProp,
		_In_opt_ const _SBinary* lpMappingSignature,
		_In_ LPSPropTagArray* lppPropTags,
		_In_opt_ LPGUID lpPropSetGuid,
		ULONG ulFlags);
	_Check_return_ std::vector<ULONG> GetIDsFromNames(
		_In_ LPMAPIPROP lpMAPIProp,
		ULONG cPropNames,
		_In_opt_count_(cPropNames) LPMAPINAMEID* lppPropNames,
		ULONG ulFlags);

	NamePropNames NameIDToStrings(
		ULONG ulPropTag, // optional 'original' prop tag
		_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
		_In_opt_ const MAPINAMEID* lpNameID, // optional named property information to avoid GetNamesFromIDs call
		_In_opt_ const _SBinary*
			lpMappingSignature, // optional mapping signature for object to speed named prop lookups
		bool
			bIsAB); // true if we know we're dealing with an address book property (they can be > 8000 and not named props)

	std::vector<std::wstring> NameIDToPropNames(_In_ const MAPINAMEID* lpNameID);
} // namespace cache2