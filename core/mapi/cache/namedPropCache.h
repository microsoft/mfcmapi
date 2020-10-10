#pragma once

namespace cache
{
	class namedPropCacheEntry;
	namespace directMapi
	{
		// Returns a vector of NamedPropCacheEntry for the input tags
		// Sourced directly from MAPI
		_Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>>
		GetNamesFromIDs(_In_ LPMAPIPROP lpMAPIProp, _In_ LPSPropTagArray* lppPropTags, ULONG ulFlags);

		// Returns a vector of NamedPropCacheEntry for the input tags
		// Sourced directly from MAPI
		_Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>>
		GetNamesFromIDs(_In_ LPMAPIPROP lpMAPIProp, _In_ const std::vector<ULONG> tags, ULONG ulFlags);

		// Returns a vector of tags for the input names
		// Sourced directly from MAPI
		_Check_return_ LPSPropTagArray
		GetIDsFromNames(_In_ LPMAPIPROP lpMAPIProp, std::vector<MAPINAMEID> nameIDs, ULONG ulFlags);
	} // namespace directMapi

	class namedPropCache
	{
	private:
		static std::list<std::shared_ptr<namedPropCacheEntry>>& getCache() noexcept;

		_Check_return_ static std::shared_ptr<namedPropCacheEntry>
		find(const std::function<bool(const std::shared_ptr<namedPropCacheEntry>&)>& compare);

	public:
		_Check_return_ static std::shared_ptr<namedPropCacheEntry>
		find(const std::shared_ptr<cache::namedPropCacheEntry>& entry, bool bMatchSig, bool bMatchID, bool bMatchName);
		_Check_return_ static std::shared_ptr<namedPropCacheEntry>
		find(_In_ const std::vector<BYTE>& _sig, _In_ const MAPINAMEID& _mapiNameId);
		_Check_return_ static std::shared_ptr<namedPropCacheEntry>
		find(_In_ const std::vector<BYTE>& _sig, ULONG _ulPropID);
		_Check_return_ static std::shared_ptr<namedPropCacheEntry>
		find(ULONG _ulPropID, _In_ const MAPINAMEID& _mapiNameId);

		// Add a mapping to the cache if it doesn't already exist
		// If given a signature, we include it in our search.
		// If not, we search without it
		static void add(const std::vector<std::shared_ptr<namedPropCacheEntry>>& entries, const std::vector<BYTE>& sig);

		// If signature is empty then do not use a signature
		_Check_return_ static std::vector<std::shared_ptr<namedPropCacheEntry>> GetNamesFromIDs(
			_In_ LPMAPIPROP lpMAPIProp,
			const std::vector<BYTE>& sig,
			_In_ LPSPropTagArray* lppPropTags);

		// If signature is empty then do not use a signature
		static _Check_return_ LPSPropTagArray GetIDsFromNames(
			_In_ LPMAPIPROP lpMAPIProp,
			_In_ const std::vector<BYTE>& sig,
			_In_ std::vector<MAPINAMEID> nameIDs,
			ULONG ulFlags);
	};
} // namespace cache