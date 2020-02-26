#pragma once
// Named Property Cache

namespace cache2
{
	struct NamePropNames;

	_Check_return_ std::shared_ptr<namedPropCacheEntry> GetNameFromID(
		_In_ LPMAPIPROP lpMAPIProp,
		_In_opt_ const std::vector<BYTE>& sig,
		_In_ ULONG ulPropTag,
		ULONG ulFlags);
	_Check_return_ std::shared_ptr<namedPropCacheEntry>
	GetNameFromID(_In_ LPMAPIPROP lpMAPIProp, _In_ ULONG ulPropTag, ULONG ulFlags);
	// No signature form: look up and use signature if possible
	_Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>>
	GetNamesFromIDs(_In_ LPMAPIPROP lpMAPIProp, _In_ LPSPropTagArray* lppPropTags, ULONG ulFlags);
	// Signature form: if signature not passed then do not use a signature
	_Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>> GetNamesFromIDs(
		_In_ LPMAPIPROP lpMAPIProp,
		_In_opt_ const std::vector<BYTE>& sig,
		_In_ LPSPropTagArray* lppPropTags,
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
		_In_opt_ const std::vector<BYTE>& sig, // optional mapping signature for object to speed named prop lookups
		bool bIsAB); // true for an address book property (they can be > 8000 and not named props)

	std::vector<std::wstring> NameIDToPropNames(_In_ const MAPINAMEID* lpNameID);
} // namespace cache2