#pragma once
// Named Property Cache Entry

namespace cache2
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
		namedPropCacheEntry::namedPropCacheEntry(
			const MAPINAMEID* lpPropName,
			ULONG _ulPropID,
			_In_ const std::vector<BYTE>& _sig = {});

		// Disables making copies of NamedPropCacheEntry
		namedPropCacheEntry(const namedPropCacheEntry&) = delete;
		namedPropCacheEntry& operator=(const namedPropCacheEntry&) = delete;

		bool valid() { return mapiNameId.Kind.lID || mapiNameId.Kind.lpwstrName; }
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
		match(const namedPropCacheEntry* entry, bool bMatchSig, bool bMatchID, bool bMatchName) const;

		// Compare given a signature, guid, kind, and value
		_Check_return_ bool match(
			_In_ const std::vector<BYTE>& _sig,
			_In_ const GUID* lpguid,
			ULONG ulKind,
			LONG lID,
			_In_z_ LPCWSTR lpwstrName) const;

		// Compare given a signature and property ID (ulPropID)
		_Check_return_ bool match(_In_ const std::vector<BYTE>& _sig, ULONG _ulPropID) const;

		// Compare given a tag, guid, kind, and value
		_Check_return_ bool
		match(ULONG _ulPropID, _In_ const GUID* lpguid, ULONG ulKind, LONG lID, _In_z_ LPCWSTR lpwstrName) const
			noexcept;

	private:
		ULONG ulPropID{}; // MAPI ID (ala PROP_ID) for a named property
		MAPINAMEID mapiNameId{}; // guid, kind, value
		GUID guid{};
		std::wstring name{};
		std::vector<BYTE> sig{}; // Value of PR_MAPPING_SIGNATURE
		NamePropNames namePropNames{};
		bool bStringsCached{}; // We have cached strings
	};
} // namespace cache2