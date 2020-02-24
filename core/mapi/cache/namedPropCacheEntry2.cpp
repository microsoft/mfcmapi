#include <core/stdafx.h>
#include <core/mapi/cache/namedPropCacheEntry2.h>
#include <core/mapi/cache/namedPropCache2.h>
#include <core/utility/error.h>
#include <core/mapi/mapiMemory.h>

namespace cache2
{
	NamedPropCacheEntry::NamedPropCacheEntry(
		ULONG _cbSig,
		_In_opt_count_(_cbSig) LPBYTE lpSig,
		LPMAPINAMEID lpPropName,
		ULONG _ulPropID)
		: ulPropID(_ulPropID)
	{
		if (lpPropName)
		{
			CopyToCacheData(*lpPropName);
		}

		if (_cbSig && lpSig)
		{
			sig.assign(lpSig, lpSig + _cbSig);
		}
	}

	_Check_return_ bool NamedPropCacheEntry::match(
		const std::shared_ptr<NamedPropCacheEntry>& entry,
		bool bMatchSig,
		bool bMatchID,
		bool bMatchName) const
	{
		if (!bMatchSig && entry->sig != sig) return false;
		if (!bMatchID && entry->ulPropID != ulPropID) return false;

		if (bMatchName)
		{
			if (entry->mapiNameId.ulKind != mapiNameId.ulKind) return false;
			if (MNID_ID == mapiNameId.ulKind && mapiNameId.Kind.lID != entry->mapiNameId.Kind.lID) return false;
			if (MNID_STRING == mapiNameId.ulKind && entry->name != name) return false;
			if (!IsEqualGUID(entry->guid, guid)) return false;
		}

		return true;
	}

	// Compare given a signature, guid, kind, and value
	_Check_return_ bool NamedPropCacheEntry::match(
		ULONG cbSig,
		_In_count_(cbSig) const BYTE* lpSig,
		_In_ const GUID* lpguid,
		ULONG ulKind,
		LONG lID,
		_In_z_ LPCWSTR lpwstrName) const noexcept
	{
		if (cbSig != sig.size()) return false;
		if (cbSig && memcmp(lpSig, sig.data(), cbSig) != 0) return false;

		if (mapiNameId.ulKind != ulKind) return false;
		if (MNID_ID == ulKind && mapiNameId.Kind.lID != lID) return false;
		if (MNID_STRING == ulKind && 0 != lstrcmpW(mapiNameId.Kind.lpwstrName, lpwstrName)) return false;
		if (0 != memcmp(mapiNameId.lpguid, lpguid, sizeof(GUID))) return false;

		return true;
	}

	// Compare given a signature and property ID (ulPropID)
	_Check_return_ bool
	NamedPropCacheEntry::match(ULONG cbSig, _In_count_(cbSig) const BYTE* lpSig, ULONG _ulPropID) const noexcept
	{
		if (sig.size() != cbSig) return false;
		if (cbSig && memcmp(lpSig, sig.data(), cbSig) != 0) return false;

		if (ulPropID != _ulPropID) return false;

		return true;
	}

	// Compare given a tag, guid, kind, and value
	_Check_return_ bool NamedPropCacheEntry::match(
		ULONG _ulPropID,
		_In_ const GUID* lpguid,
		ULONG ulKind,
		LONG lID,
		_In_z_ LPCWSTR lpwstrName) const noexcept
	{
		if (ulPropID != _ulPropID) return false;

		if (mapiNameId.ulKind != ulKind) return false;
		if (MNID_ID == ulKind && mapiNameId.Kind.lID != lID) return false;
		if (MNID_STRING == ulKind && 0 != lstrcmpW(mapiNameId.Kind.lpwstrName, lpwstrName)) return false;
		if (0 != memcmp(mapiNameId.lpguid, lpguid, sizeof(GUID))) return false;

		return true;
	}

	// Go through all the details of copying allocated data to a cache entry
	void NamedPropCacheEntry::CopyToCacheData(const MAPINAMEID& src)
	{
		mapiNameId.lpguid = nullptr;
		mapiNameId.Kind.lID = 0;

		if (src.lpguid)
		{
			guid = *src.lpguid;
			mapiNameId.lpguid = &guid;
		}

		mapiNameId.ulKind = src.ulKind;
		if (MNID_ID == src.ulKind)
		{
			mapiNameId.Kind.lID = src.Kind.lID;
		}
		else if (MNID_STRING == src.ulKind)
		{
			if (src.Kind.lpwstrName)
			{
				const auto cbName = cbPropName(src.Kind.lpwstrName);

				name = std::wstring(src.Kind.lpwstrName, cbName / sizeof WCHAR);
				mapiNameId.Kind.lpwstrName = name.data();
			}
		}
	}
} // namespace cache2