#include <core/stdafx.h>
#include <core/mapi/cache/namedProps.h>
#include <core/mapi/cache/namedPropCache.h>
#include <core/interpret/guid.h>
#include <core/mapi/mapiMemory.h>
#include <core/mapi/mapiFunctions.h>
#include <core/utility/registry.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/addin/mfcmapi.h>
#include <core/addin/addin.h>
#include <core/utility/error.h>

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
			cbName = ULONG((cchWideLen + 1) * sizeof WCHAR);
		}
		else
		{
			// This is the case where ANSI data was shoved into a unicode string.
			// Add a couple extra NULL in case we read this as unicode again.
			cbName = ULONG((cchShortLen + 3) * sizeof CHAR);
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

	_Check_return_ bool namedPropCacheEntry::match(
		const std::shared_ptr<namedPropCacheEntry>& entry,
		bool bMatchSig,
		bool bMatchID,
		bool bMatchName) const
	{
		if (!entry) return false;
		if (bMatchSig && entry->sig != sig) return false;
		if (bMatchID && entry->ulPropID != ulPropID) return false;

		if (bMatchName)
		{
			if (entry->mapiNameId.ulKind != mapiNameId.ulKind) return false;
			if (MNID_ID == mapiNameId.ulKind && mapiNameId.Kind.lID != entry->mapiNameId.Kind.lID) return false;
			if (MNID_STRING == mapiNameId.ulKind && entry->name != name) return false;
			if (!IsEqualGUID(entry->guid, guid)) return false;
		}

		return true;
	}

	// Compare given a signature, MAPINAMEID
	// If signature is empty then do not use a signature
	_Check_return_ bool
	namedPropCacheEntry::match(_In_ const std::vector<BYTE>& _sig, _In_ const MAPINAMEID& _mapiNameId) const
	{
		if (!_sig.empty() && sig != _sig) return false;

		if (mapiNameId.ulKind != _mapiNameId.ulKind) return false;
		if (mapiNameId.ulKind == MNID_ID && mapiNameId.Kind.lID != _mapiNameId.Kind.lID) return false;
		if (mapiNameId.ulKind == MNID_STRING && 0 != lstrcmpW(mapiNameId.Kind.lpwstrName, _mapiNameId.Kind.lpwstrName))
			return false;
		if (0 != memcmp(mapiNameId.lpguid, _mapiNameId.lpguid, sizeof(GUID))) return false;

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

	// Compare given a id, MAPINAMEID
	_Check_return_ bool namedPropCacheEntry::match(ULONG _ulPropID, _In_ const MAPINAMEID& _mapiNameId) const noexcept
	{
		if (ulPropID != _ulPropID) return false;

		if (mapiNameId.ulKind != _mapiNameId.ulKind) return false;
		if (mapiNameId.ulKind == MNID_ID && mapiNameId.Kind.lID != _mapiNameId.Kind.lID) return false;
		if (mapiNameId.ulKind == MNID_STRING && 0 != lstrcmpW(mapiNameId.Kind.lpwstrName, _mapiNameId.Kind.lpwstrName))
			return false;
		if (0 != memcmp(mapiNameId.lpguid, _mapiNameId.lpguid, sizeof(GUID))) return false;

		return true;
	}

	void namedPropCacheEntry::output() const
	{
		if (fIsSet(output::dbgLevel::NamedPropCache))
		{
			output::DebugPrint(output::dbgLevel::NamedPropCache, L"id=%04X\n", ulPropID);
			const auto nameidString = strings::MAPINAMEIDToString(mapiNameId);
			output::DebugPrint(output::dbgLevel::NamedPropCache, L"%ws\n", nameidString.c_str());
			if (!sig.empty())
			{
				const auto sigStr = strings::BinToHexString(sig, true);
				output::DebugPrint(output::dbgLevel::NamedPropCache, L"sig=%ws\n", sigStr.c_str());
			}
		}
	}

	_Check_return_ std::shared_ptr<namedPropCacheEntry>
	GetNameFromID(_In_ LPMAPIPROP lpMAPIProp, _In_ ULONG ulPropTag, ULONG ulFlags)
	{
		auto tag = SPropTagArray{1, ulPropTag};
		auto lptag = &tag;
		const auto names = GetNamesFromIDs(lpMAPIProp, &lptag, ulFlags);
		if (names.size() == 1) return names[0];
		return {};
	}

	_Check_return_ std::shared_ptr<namedPropCacheEntry>
	GetNameFromID(_In_ LPMAPIPROP lpMAPIProp, _In_opt_ const SBinary* sig, _In_ ULONG ulPropTag, ULONG ulFlags)
	{
		auto tag = SPropTagArray{1, ulPropTag};
		auto lptag = &tag;
		auto names = std::vector<std::shared_ptr<namedPropCacheEntry>>{};
		if (sig)
		{
			names = GetNamesFromIDs(lpMAPIProp, sig, &lptag, ulFlags);
		}
		else
		{
			names = GetNamesFromIDs(lpMAPIProp, &lptag, ulFlags);
		}

		if (names.size() == 1) return names[0];
		return {};
	}

	// No signature form: look up and use signature if possible
	_Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>>
	GetNamesFromIDs(_In_ LPMAPIPROP lpMAPIProp, _In_opt_ LPSPropTagArray* lppPropTags, ULONG ulFlags)
	{
		SBinary sig = {};
		LPSPropValue lpProp = nullptr;
		// This error is too chatty to log - ignore it.
		const auto hRes = HrGetOneProp(lpMAPIProp, PR_MAPPING_SIGNATURE, &lpProp);
		if (SUCCEEDED(hRes) && lpProp && PT_BINARY == PROP_TYPE(lpProp->ulPropTag))
		{
			sig = mapi::getBin(lpProp);
		}

		const auto names = GetNamesFromIDs(lpMAPIProp, &sig, lppPropTags, ulFlags);
		MAPIFreeBuffer(lpProp);
		return names;
	}

	// Signature form: if signature is empty then do not use a signature
	_Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>> GetNamesFromIDs(
		_In_ LPMAPIPROP lpMAPIProp,
		_In_opt_ const SBinary* sig,
		_In_opt_ LPSPropTagArray* lppPropTags,
		ULONG ulFlags)
	{
		if (!lpMAPIProp) return {};

		// Check if we're bypassing the cache:
		if (!registry::cacheNamedProps ||
			// None of my code uses these flags, but bypass the cache if we see them
			ulFlags)
		{
			return directMapi::GetNamesFromIDs(lpMAPIProp, lppPropTags, ulFlags);
		}

		auto sigv = std::vector<BYTE>{};
		if (sig && sig->lpb && sig->cb) sigv = {sig->lpb, sig->lpb + sig->cb};
		return namedPropCache::GetNamesFromIDs(lpMAPIProp, sigv, lppPropTags);
	}

	_Check_return_ LPSPropTagArray
	GetIDsFromNames(_In_ LPMAPIPROP lpMAPIProp, _In_ std::vector<MAPINAMEID> nameIDs, _In_ ULONG ulFlags)
	{
		if (!lpMAPIProp) return {};

		// Check if we're bypassing the cache:
		if (!registry::cacheNamedProps ||
			// If no names were passed, we have to bypass the cache
			// Should we cache results?
			nameIDs.empty())
		{
			return directMapi::GetIDsFromNames(lpMAPIProp, nameIDs, ulFlags);
		}

		// Get a signature if we can.
		LPSPropValue lpProp = nullptr;
		auto sig = std::vector<BYTE>{};
		WC_MAPI_S(HrGetOneProp(lpMAPIProp, PR_MAPPING_SIGNATURE, &lpProp));

		if (lpProp && PT_BINARY == PROP_TYPE(lpProp->ulPropTag))
		{
			const auto bin = mapi::getBin(lpProp);
			sig = {bin.lpb, bin.lpb + bin.cb};
		}

		MAPIFreeBuffer(lpProp);

		return namedPropCache::GetIDsFromNames(lpMAPIProp, sig, nameIDs, ulFlags);
	}

	// TagToString will prepend the http://schemas.microsoft.com/MAPI/ for us since it's a constant
	// We don't compute a DASL string for non-named props as FormatMessage in TagToString can handle those
	NamePropNames NameIDToStrings(_In_ const MAPINAMEID* lpNameID, ULONG ulPropTag)
	{
		// Can't generate strings without a MAPINAMEID structure
		if (!lpNameID) return {};

		auto lpNamedPropCacheEntry = std::shared_ptr<namedPropCacheEntry>{};

		// If we're using the cache, look up the answer there and return
		if (registry::cacheNamedProps)
		{
			lpNamedPropCacheEntry = namedPropCache::find(PROP_ID(ulPropTag), *lpNameID);
			if (lpNamedPropCacheEntry && lpNamedPropCacheEntry->hasCachedStrings())
			{
				return lpNamedPropCacheEntry->getNamePropNames();
			}

			// We shouldn't ever get here without a cached entry
			if (!lpNamedPropCacheEntry)
			{
				output::DebugPrint(
					output::dbgLevel::NamedProp,
					L"NameIDToStrings: Failed to find cache entry for ulPropTag = 0x%08X\n",
					ulPropTag);
				return {};
			}
		}

		output::DebugPrint(output::dbgLevel::NamedProp, L"Parsing named property\n");
		output::DebugPrint(output::dbgLevel::NamedProp, L"ulPropTag = 0x%08x\n", ulPropTag);
		NamePropNames namePropNames;
		namePropNames.guid = guid::GUIDToStringAndName(lpNameID->lpguid);
		output::DebugPrint(output::dbgLevel::NamedProp, L"lpNameID->lpguid = %ws\n", namePropNames.guid.c_str());

		auto szDASLGuid = guid::GUIDToString(lpNameID->lpguid);

		if (lpNameID->ulKind == MNID_ID)
		{
			output::DebugPrint(
				output::dbgLevel::NamedProp,
				L"lpNameID->Kind.lID = 0x%04X = %d\n",
				lpNameID->Kind.lID,
				lpNameID->Kind.lID);
			auto pidlids = NameIDToPropNames(lpNameID);

			if (!pidlids.empty())
			{
				namePropNames.bestPidLid = pidlids.front();
				pidlids.erase(pidlids.begin());
				namePropNames.otherPidLid = strings::join(pidlids, L", ");
				// Printing hex first gets a nice sort without spacing tricks
				namePropNames.name = strings::format(
					L"id: 0x%04X=%d = %ws", // STRING_OK
					lpNameID->Kind.lID,
					lpNameID->Kind.lID,
					namePropNames.bestPidLid.c_str());

				if (!namePropNames.otherPidLid.empty())
				{
					namePropNames.name += strings::format(L" (%ws)", namePropNames.otherPidLid.c_str());
				}
			}
			else
			{
				// Printing hex first gets a nice sort without spacing tricks
				namePropNames.name = strings::format(
					L"id: 0x%04X=%d", // STRING_OK
					lpNameID->Kind.lID,
					lpNameID->Kind.lID);
			}

			namePropNames.dasl = strings::format(
				L"id/%ws/%04X%04X", // STRING_OK
				szDASLGuid.c_str(),
				lpNameID->Kind.lID,
				PROP_TYPE(ulPropTag));
		}
		else if (lpNameID->ulKind == MNID_STRING)
		{
			// lpwstrName is LPWSTR which means it's ALWAYS unicode
			// But some folks get it wrong and stuff ANSI data in there
			// So we check the string length both ways to make our best guess
			const auto cchShortLen = strnlen_s(reinterpret_cast<LPCSTR>(lpNameID->Kind.lpwstrName), RSIZE_MAX);
			const auto cchWideLen = wcsnlen_s(lpNameID->Kind.lpwstrName, RSIZE_MAX);

			if ((cchShortLen == 0 && cchWideLen == 0) || cchShortLen < cchWideLen)
			{
				// this is the *proper* case
				output::DebugPrint(
					output::dbgLevel::NamedProp, L"lpNameID->Kind.lpwstrName = \"%ws\"\n", lpNameID->Kind.lpwstrName);
				namePropNames.name = lpNameID->Kind.lpwstrName;

				namePropNames.dasl = strings::format(
					L"string/%ws/%ws", // STRING_OK
					szDASLGuid.c_str(),
					lpNameID->Kind.lpwstrName);
			}
			else
			{
				// this is the case where ANSI data was shoved into a unicode string.
				output::DebugPrint(
					output::dbgLevel::NamedProp,
					L"Warning: ANSI data was found in a unicode field. This is a bug on the part of the creator of "
					L"this named property\n");
				output::DebugPrint(
					output::dbgLevel::NamedProp,
					L"lpNameID->Kind.lpwstrName = \"%hs\"\n",
					reinterpret_cast<LPCSTR>(lpNameID->Kind.lpwstrName));

				auto szComment = strings::loadstring(IDS_NAMEWASANSI);
				namePropNames.name =
					strings::format(L"%hs %ws", reinterpret_cast<LPSTR>(lpNameID->Kind.lpwstrName), szComment.c_str());

				namePropNames.dasl = strings::format(
					L"string/%ws/%hs", // STRING_OK
					szDASLGuid.c_str(),
					LPSTR(lpNameID->Kind.lpwstrName));
			}
		}

		// We've built our strings - if we're caching, put them in the cache
		if (lpNamedPropCacheEntry)
		{
			lpNamedPropCacheEntry->setNamePropNames(namePropNames);
		}

		return namePropNames;
	}

	NamePropNames NameIDToStrings(
		ULONG ulPropTag, // optional 'original' prop tag
		_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
		_In_opt_ const MAPINAMEID* lpNameID, // optional named property information to avoid GetNamesFromIDs call
		_In_opt_ const SBinary* sig, // optional mapping signature for object to speed named prop lookups
		bool bIsAB) // true for an address book property (they can be > 8000 and not named props)
	{
		// If we weren't passed named property information and we need it, look it up
		// We check bIsAB here - some address book providers return garbage which will crash us
		if (!lpNameID && lpMAPIProp && // if we have an object
			!bIsAB && registry::parseNamedProps && // and we're parsing named props
			(registry::getPropNamesOnAllProps ||
			 PROP_ID(ulPropTag) >= 0x8000)) // and it's either a named prop or we're doing all props
		{
			const auto name = GetNameFromID(lpMAPIProp, sig, ulPropTag, NULL);
			if (cache::namedPropCacheEntry::valid(name))
			{
				return NameIDToStrings(name->getMapiNameId(), ulPropTag);
			}
		}

		if (lpNameID) return NameIDToStrings(lpNameID, ulPropTag);
		return {};
	}

	// Returns string built from NameIDArray
	std::vector<std::wstring> NameIDToPropNames(_In_ const MAPINAMEID* lpNameID)
	{
		std::vector<std::wstring> results;
		if (!lpNameID) return {};
		if (lpNameID->ulKind != MNID_ID) return {};
		ULONG ulMatch = ulNoMatch;

		if (NameIDArray.empty()) return {};

		for (ULONG ulCur = 0; ulCur < NameIDArray.size(); ulCur++)
		{
			if (NameIDArray[ulCur].lValue == lpNameID->Kind.lID)
			{
				ulMatch = ulCur;
				break;
			}
		}

		if (ulNoMatch != ulMatch)
		{
			for (auto ulCur = ulMatch; ulCur < NameIDArray.size(); ulCur++)
			{
				if (NameIDArray[ulCur].lValue != lpNameID->Kind.lID) break;
				// We don't acknowledge array entries without guids
				if (!NameIDArray[ulCur].lpGuid) continue;
				// But if we weren't asked about a guid, we don't check one
				if (lpNameID->lpguid && !IsEqualGUID(*lpNameID->lpguid, *NameIDArray[ulCur].lpGuid)) continue;

				results.push_back(NameIDArray[ulCur].lpszName);
			}
		}

		return results;
	}

	ULONG FindHighestNamedProp(_In_ LPMAPIPROP lpMAPIProp)
	{
		output::DebugPrint(
			output::dbgLevel::NamedProp, L"FindHighestNamedProp: Searching for the highest named prop mapping\n");

		ULONG ulLower = __LOWERBOUND;
		ULONG ulUpper = __UPPERBOUND;
		ULONG ulHighestKnown = 0;
		auto ulCurrent = (ulUpper + ulLower) / 2;

		while (ulUpper - ulLower > 1)
		{
			const auto ulPropTag = PROP_TAG(NULL, ulCurrent);
			const auto name = cache::GetNameFromID(lpMAPIProp, ulPropTag, NULL);
			if (cache::namedPropCacheEntry::valid(name))
			{
				// Found a named property, reset lower bound

				// Avoid NameIDToStrings call if we're not debug printing
				if (fIsSet(output::dbgLevel::NamedProp))
				{
					output::DebugPrint(
						output::dbgLevel::NamedProp,
						L"FindHighestNamedProp: Found a named property at 0x%04X.\n",
						ulCurrent);
					const auto namePropNames =
						cache::NameIDToStrings(ulPropTag, nullptr, name->getMapiNameId(), nullptr, false);
					output::DebugPrint(
						output::dbgLevel::NamedProp,
						L"FindHighestNamedProp: Name = %ws, GUID = %ws\n",
						namePropNames.name.c_str(),
						namePropNames.guid.c_str());
				}

				ulHighestKnown = ulCurrent;
				ulLower = ulCurrent;
			}
			else
			{
				// Did not find a named property, reset upper bound
				ulUpper = ulCurrent;
			}

			ulCurrent = (ulUpper + ulLower) / 2;
		}

		return ulHighestKnown;
	}
} // namespace cache