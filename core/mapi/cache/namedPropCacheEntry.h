#pragma once
// Named Property Cache Entry

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

	class NamedPropCacheEntry
	{
	public:
		NamedPropCacheEntry(
			ULONG _cbSig,
			_In_opt_count_(_cbSig) LPBYTE lpSig,
			LPMAPINAMEID lpPropName,
			ULONG _ulPropID);

		// Disables making copies of NamedPropCacheEntry
		NamedPropCacheEntry(const NamedPropCacheEntry&) = delete;
		NamedPropCacheEntry& operator=(const NamedPropCacheEntry&) = delete;

		ULONG getPropID() const noexcept { return ulPropID; }
		const NamePropNames& getNamePropNames() const noexcept { return namePropNames; }
		void setNamePropNames(const NamePropNames& _namePropNames) noexcept
		{
			namePropNames = _namePropNames;
			bStringsCached = true;
		}
		bool hasCachedStrings() const noexcept { return bStringsCached; }
		const MAPINAMEID* getMapiNameId() const noexcept { return &mapiNameId; }

		// Compare given a signature, guid, kind, and value
		_Check_return_ bool match(
			ULONG cbSig,
			_In_count_(cbSig) const BYTE* lpSig,
			_In_ const GUID* lpguid,
			ULONG ulKind,
			LONG lID,
			_In_z_ LPCWSTR lpwstrName) const noexcept;

		// Compare given a signature and property ID (ulPropID)
		_Check_return_ bool match(ULONG cbSig, _In_count_(cbSig) const BYTE* lpSig, ULONG _ulPropID) const noexcept;

		// Compare given a tag, guid, kind, and value
		_Check_return_ bool
		match(ULONG _ulPropID, _In_ const GUID* lpguid, ULONG ulKind, LONG lID, _In_z_ LPCWSTR lpwstrName) const
			noexcept;

	private:
		ULONG ulPropID{}; // MAPI ID (ala PROP_ID) for a named property
		MAPINAMEID mapiNameId{}; // guid, kind, value
		GUID guid;
		std::wstring name;
		std::vector<BYTE> sig{}; // Value of PR_MAPPING_SIGNATURE
		NamePropNames namePropNames{};
		bool bStringsCached{}; // We have cached strings

		// Go through all the details of copying allocated data to a cache entry
		void CopyToCacheData(const MAPINAMEID& src);
	};
} // namespace cache