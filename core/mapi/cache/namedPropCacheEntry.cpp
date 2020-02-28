#include <core/stdafx.h>
#include <core/mapi/cache/namedPropCacheEntry.h>
#include <core/mapi/cache/namedPropCache.h>
#include <core/utility/error.h>
#include <core/mapi/mapiMemory.h>

namespace cache
{
	static ULONG cbPropName(LPCWSTR lpwstrName) noexcept
	{
		// lpwstrName is LPWSTR which means it's ALWAYS unicode
		// But some folks get it wrong and stuff ANSI data in there
		// So we check the string length both ways to make our best guess
#pragma warning(push)
#pragma warning(disable : 26490) // warning C26490: Don't use reinterpret_cast (type.1).
		const auto cchShortLen = strnlen_s(reinterpret_cast<LPCSTR>(lpwstrName), RSIZE_MAX);
#pragma warning(pop)
		const auto cchWideLen = wcsnlen_s(lpwstrName, RSIZE_MAX);
		auto cbName = ULONG();

		if (cchShortLen < cchWideLen)
		{
			// this is the *proper* case
			cbName = (cchWideLen + 1) * sizeof WCHAR;
		}
		else
		{
			// This is the case where ANSI data was shoved into a unicode string.
			// Add a couple extra NULL in case we read this as unicode again.
			cbName = (cchShortLen + 3) * sizeof CHAR;
		}

		return cbName;
	}

	namedPropCacheEntry::namedPropCacheEntry(
		const MAPINAMEID* lpPropName,
		ULONG _ulPropID,
		_In_ const std::vector<BYTE>& _sig)
		: ulPropID(_ulPropID), sig(_sig)
	{
		if (lpPropName)
		{
			if (lpPropName->lpguid)
			{
				guid = *lpPropName->lpguid;
				mapiNameId.lpguid = &guid;
			}

			mapiNameId.ulKind = lpPropName->ulKind;
			if (lpPropName->ulKind == MNID_ID)
			{
				mapiNameId.Kind.lID = lpPropName->Kind.lID;
			}
			else if (lpPropName->ulKind == MNID_STRING)
			{
				if (lpPropName->Kind.lpwstrName)
				{
					const auto cbName = cbPropName(lpPropName->Kind.lpwstrName);
					name = std::wstring(lpPropName->Kind.lpwstrName, cbName / sizeof WCHAR);
					mapiNameId.Kind.lpwstrName = name.data();
				}
			}
		}
	}

	_Check_return_ bool
	namedPropCacheEntry::match(const namedPropCacheEntry* entry, bool bMatchSig, bool bMatchID, bool bMatchName) const
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
	// If signature is empty then do not use a signature
	_Check_return_ bool namedPropCacheEntry::match(
		_In_ const std::vector<BYTE>& _sig,
		_In_ const GUID* lpguid,
		ULONG ulKind,
		LONG lID,
		_In_z_ LPCWSTR lpwstrName) const
	{
		if (!_sig.empty() && sig != _sig) return false;

		if (mapiNameId.ulKind != ulKind) return false;
		if (MNID_ID == ulKind && mapiNameId.Kind.lID != lID) return false;
		if (MNID_STRING == ulKind && 0 != lstrcmpW(mapiNameId.Kind.lpwstrName, lpwstrName)) return false;
		if (0 != memcmp(mapiNameId.lpguid, lpguid, sizeof(GUID))) return false;

		return true;
	}

	// Compare given a signature and property ID (ulPropID)
	// If signature is empty then do not use a signature
	_Check_return_ bool namedPropCacheEntry::match(_In_ const std::vector<BYTE>& _sig, ULONG _ulPropID) const
	{
		if (!_sig.empty() && sig != _sig) return false;
		if (ulPropID != _ulPropID) return false;

		return true;
	}

	// Compare given a tag, guid, kind, and value
	_Check_return_ bool namedPropCacheEntry::match(
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
} // namespace cache