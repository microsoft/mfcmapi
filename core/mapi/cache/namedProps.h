#pragma once
// Named Property Cache
#include <core/addin/addin.h>
#include <core/addin/MFCMAPI.h>

namespace cache
{
	struct NamePropNames
	{
		std::wstring name;
		std::wstring guid;
		std::wstring dasl;
		std::wstring bestPidLid;
		std::wstring otherPidLid;
	};

	class namedPropCacheEntry
	{
	public:
		namedPropCacheEntry(const MAPINAMEID* lpPropName, ULONG _ulPropID, _In_ const std::vector<BYTE>& _sig = {});

		// Disables making copies of NamedPropCacheEntry
		namedPropCacheEntry(const namedPropCacheEntry&) = delete;
		namedPropCacheEntry& operator=(const namedPropCacheEntry&) = delete;

		static std::shared_ptr<namedPropCacheEntry>& empty() noexcept
		{
			static std::shared_ptr<namedPropCacheEntry> _empty = std::make_shared<namedPropCacheEntry>(nullptr, 0);
			return _empty;
		}

		_Check_return_ inline static std::shared_ptr<cache::namedPropCacheEntry>
		make(const MAPINAMEID* lpPropName, ULONG _ulPropID, _In_ const std::vector<BYTE>& _sig = {})
		{
			return std::make_shared<cache::namedPropCacheEntry>(lpPropName, _ulPropID, _sig);
		}

		static inline bool valid(std::shared_ptr<cache::namedPropCacheEntry> item) noexcept
		{
			return item && item->ulPropID && item->ulPropID != PT_ERROR &&
				   (item->mapiNameId.Kind.lID || item->mapiNameId.Kind.lpwstrName);
		}
		ULONG getPropID() const noexcept { return ulPropID; }
		const NamePropNames& getNamePropNames() const noexcept { return namePropNames; }
		void setNamePropNames(const NamePropNames& _namePropNames) noexcept
		{
			namePropNames = _namePropNames;
			bStringsCached = true;
		}
		bool hasCachedStrings() const noexcept { return bStringsCached; }
		const MAPINAMEID* getMapiNameId() const noexcept { return &mapiNameId; }
		void setSig(const std::vector<BYTE>& _sig) { sig = _sig; }

		_Check_return_ bool
		match(const std::shared_ptr<namedPropCacheEntry>& entry, bool bMatchID, bool bMatchName) const;

		// Compare given a signature, MAPINAMEID
		_Check_return_ bool match(_In_ const std::vector<BYTE>& _sig, _In_ const MAPINAMEID& _mapiNameId) const;

		// Compare given a signature and property ID (ulPropID)
		_Check_return_ bool match(_In_ const std::vector<BYTE>& _sig, ULONG _ulPropID) const;

		// Compare given a id, MAPINAMEID
		_Check_return_ bool match(ULONG _ulPropID, _In_ const MAPINAMEID& _mapiNameId) const noexcept;

		void output() const;

	private:
		ULONG ulPropID{}; // MAPI ID (ala PROP_ID) for a named property
		MAPINAMEID mapiNameId{}; // guid, kind, value
		GUID guid{};
		std::wstring name{};
		std::vector<BYTE> sig{}; // Value of PR_MAPPING_SIGNATURE
		NamePropNames namePropNames{};
		bool bStringsCached{}; // We have cached strings
	};

	constexpr ULONG __LOWERBOUND = 0x8000;
	constexpr ULONG __UPPERBOUND = 0xFFFF;

	_Check_return_ std::shared_ptr<namedPropCacheEntry>
	GetNameFromID(_In_ const LPMAPIPROP lpMAPIProp, _In_opt_ const SBinary* sig, _In_ ULONG ulPropTag, ULONG ulFlags);
	_Check_return_ std::shared_ptr<namedPropCacheEntry>
	GetNameFromID(_In_ const LPMAPIPROP lpMAPIProp, _In_ ULONG ulPropTag, ULONG ulFlags);

	// No signature form: look up and use signature if possible
	_Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>>
	GetNamesFromIDs(_In_opt_ const LPMAPIPROP lpMAPIProp, _In_opt_ LPSPropTagArray lpPropTags, ULONG ulFlags);
	// Signature form: if signature not passed then do not use a signature
	_Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>> GetNamesFromIDs(
		_In_opt_ const LPMAPIPROP lpMAPIProp,
		_In_opt_ const SBinary* sig,
		_In_opt_ LPSPropTagArray lpPropTags,
		ULONG ulFlags);

	_Check_return_ LPSPropTagArray
	GetIDsFromNames(_In_ const LPMAPIPROP lpMAPIProp, _In_ std::vector<MAPINAMEID> nameIDs, _In_ ULONG ulFlags);

	NamePropNames NameIDToStrings(
		ULONG ulPropTag, // optional 'original' prop tag
		_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
		_In_opt_ const MAPINAMEID* lpNameID, // optional named property information to avoid GetNamesFromIDs call
		_In_opt_ const SBinary* sig, // optional mapping signature for object to speed named prop lookups
		bool bIsAB); // true for an address book property (they can be > 8000 and not named props)

	std::vector<std::wstring> NameIDToPropNames(_In_opt_ const MAPINAMEID* lpNameID);

	ULONG FindHighestNamedProp(_In_ LPMAPIPROP lpMAPIProp);

	_Check_return_ std::shared_ptr<namedPropCacheEntry> find(
		const std::vector<std::shared_ptr<namedPropCacheEntry>>& list,
		const std::function<bool(const std::shared_ptr<namedPropCacheEntry>&)>& compare);

	void FindNameIDArrayMatches(_In_ LONG lTarget, _Out_ ULONG* lpulNumExacts, _Out_ ULONG* lpulFirstExact) noexcept;
	_Check_return_ LPNAMEID_ARRAY_ENTRY GetDispIDFromName(_In_z_ LPCWSTR lpszDispIDName);
} // namespace cache