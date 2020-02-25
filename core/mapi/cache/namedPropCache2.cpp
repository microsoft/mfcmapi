#include <core/stdafx.h>
#include <core/mapi/cache/namedPropCacheEntry2.h>
#include <core/mapi/cache/namedPropCache2.h>
#include <core/interpret/guid.h>
#include <core/mapi/mapiMemory.h>
#include <core/utility/registry.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/addin/mfcmapi.h>
#include <core/addin/addin.h>
#include <core/utility/error.h>

namespace cache2
{
	class namedPropCache
	{
	public:
		static std::list<std::shared_ptr<namedPropCacheEntry>>& getCache() noexcept
		{
			// We keep a list of named prop cache entries
			static std::list<std::shared_ptr<namedPropCacheEntry>> cache;
			return cache;
		}

		// Returns a vector of NamedPropCacheEntry for the input tags
		// Sourced directly from MAPI
		static _Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>> GetNamesFromIDs(
			_In_ LPMAPIPROP lpMAPIProp,
			_In_ LPSPropTagArray* lppPropTags,
			_In_opt_ LPGUID lpPropSetGuid,
			ULONG ulFlags)
		{
			if (!lpMAPIProp) return {};

			LPMAPINAMEID* lppPropNames = nullptr;
			auto ulPropNames = ULONG{};

			WC_H_GETPROPS_S(
				lpMAPIProp->GetNamesFromIDs(lppPropTags, lpPropSetGuid, ulFlags, &ulPropNames, &lppPropNames));

			auto ids = std::vector<std::shared_ptr<namedPropCacheEntry>>{};

			if (ulPropNames && lppPropNames)
			{
				for (ULONG i = 0; i < ulPropNames; i++)
				{
					auto ulPropTag = ULONG{};
					if (lppPropTags && *lppPropTags) ulPropTag = (*lppPropTags)->aulPropTag[i];
					ids.emplace_back(std::make_shared<namedPropCacheEntry>(lppPropNames[i], ulPropTag));
				}
			}

			MAPIFreeBuffer(lppPropNames);
			return ids;
		}

		// Returns a vector of NamedPropCacheEntry for the input tags
		// Sourced directly from MAPI
		static _Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>> GetNamesFromIDs(
			_In_ LPMAPIPROP lpMAPIProp,
			_In_ const std::vector<ULONG> tags,
			_In_opt_ LPGUID lpPropSetGuid,
			ULONG ulFlags)
		{
			if (!lpMAPIProp) return {};

			auto ids = std::vector<std::shared_ptr<namedPropCacheEntry>>{};
			auto ulPropTags = mapi::allocate<LPSPropTagArray>(CbNewSPropTagArray(tags.size()));
			if (ulPropTags)
			{
				ulPropTags->cValues = tags.size();
				ULONG i = 0;
				for (const auto& tag : tags)
				{
					ulPropTags->aulPropTag[i++] = tag;
				}

				ids = GetNamesFromIDs(lpMAPIProp, &ulPropTags, lpPropSetGuid, ulFlags);
			}

			MAPIFreeBuffer(ulPropTags);

			return ids;
		}

		// Returns a vector of tags for the input names
		// Sourced directly from MAPI
		static _Check_return_ std::vector<ULONG> GetIDsFromNames(
			_In_ LPMAPIPROP lpMAPIProp,
			ULONG cPropNames,
			_In_opt_count_(cPropNames) LPMAPINAMEID* lppPropNames,
			ULONG ulFlags)
		{
			if (!lpMAPIProp) return {};

			LPSPropTagArray lpTags = nullptr;

			WC_H_GETPROPS_S(lpMAPIProp->GetIDsFromNames(cPropNames, lppPropNames, ulFlags, &lpTags));

			auto tags = std::vector<ULONG>{};

			if (lpTags)
			{
				for (ULONG i = 0; i < cPropNames; i++)
				{
					tags.emplace_back(lpTags->aulPropTag[i]);
				}
			}

			MAPIFreeBuffer(lpTags);
			return tags;
		}

		// Returns a vector of tags for the input names
		// Sourced directly from MAPI
		static _Check_return_ std::vector<ULONG>
		GetIDsFromNames(_In_ LPMAPIPROP lpMAPIProp, _In_ const std::vector<MAPINAMEID*>& names, ULONG ulFlags)
		{
			if (!lpMAPIProp) return {};

			auto lppNames = mapi::allocate<LPMAPINAMEID*>(sizeof(LPMAPINAMEID) * names.size());
			if (lppNames)
			{
				ULONG i = 0;
				for (auto name : names)
				{
					lppNames[i] = name;
				}
			}

			auto ids = GetIDsFromNames(lpMAPIProp, names.size(), lppNames, ulFlags);

			MAPIFreeBuffer(lppNames);

			return ids;
		}

		_Check_return_ static std::shared_ptr<namedPropCacheEntry>
		FindCacheEntry(const std::function<bool(const std::shared_ptr<namedPropCacheEntry>&)>& compare) noexcept
		{
			const auto& cache = getCache();
			const auto entry = find_if(cache.begin(), cache.end(), [compare](const auto& namedPropCacheEntry) noexcept {
				return compare(namedPropCacheEntry);
			});

			return entry != cache.end() ? *entry : nullptr;
		}

		// Add a mapping to the cache if it doesn't already exist
		// If given a signature, we include it in our search.
		// If not, we search without it
		static void AddMapping(std::vector<std::shared_ptr<namedPropCacheEntry>>& entries, const std::vector<BYTE>& sig)
		{
			auto& cache = getCache();
			for (auto& entry : entries)
			{
				auto match = std::shared_ptr<namedPropCacheEntry>{};
				if (sig.empty())
				{
					match = FindCacheEntry(
						[&](const auto& _entry) noexcept { return entry->match(_entry, false, true, true); });
				}
				else
				{
					entry->setSig(sig);
					match = FindCacheEntry(
						[&](const auto& _entry) noexcept { return entry->match(_entry, true, true, true); });
				}

				if (!match)
				{
					cache.emplace_back(entry);
				}
				//if (fIsSet(output::dbgLevel::NamedPropCacheMisses) && entry->->ulKind == MNID_ID)
				//{
				//	auto names = NameIDToPropNames(lppPropNames[ulSource]);
				//	if (names.empty())
				//	{
				//		output::DebugPrint(
				//			output::dbgLevel::NamedPropCacheMisses,
				//			L"AddMapping: Caching unknown property 0x%08X %ws\n",
				//			lppPropNames[ulSource]->Kind.lID,
				//			guid::GUIDToStringAndName(lppPropNames[ulSource]->lpguid).c_str());
				//	}
				//}
			}
		}

		_Check_return_ static std::vector<std::shared_ptr<namedPropCacheEntry>> CacheGetNamesFromIDs(
			_In_ LPMAPIPROP lpMAPIProp,
			const std::vector<BYTE>& sig,
			_In_ LPSPropTagArray* lppPropTags)
		{
			if (!lpMAPIProp || !lppPropTags || !*lppPropTags || sig.empty()) return {};

			// We're going to walk the cache, looking for the values we need. As soon as we have all the values we need, we're done
			// If we reach the end of the cache and don't have everything, we set up to make a GetNamesFromIDs call.

			auto results = std::vector<std::shared_ptr<namedPropCacheEntry>>{};
			const auto lpPropTags = *lppPropTags;

			auto misses = std::vector<ULONG>{};

			// First pass, find any misses we might have
			for (ULONG ulTarget = 0; ulTarget < lpPropTags->cValues; ulTarget++)
			{
				auto ulPropId = PROP_ID(lpPropTags->aulPropTag[ulTarget]);
				// ...check the cache
				const auto lpEntry =
					FindCacheEntry([&](const auto& entry) noexcept { return entry->match(sig, ulPropId); });

				if (lpEntry)
				{
					results.emplace_back(lpEntry);
				}
			}

			// Go to MAPI with whatever's left. We set up for a single call to GetNamesFromIDs.
			if (!misses.empty())
			{
				auto missed = GetNamesFromIDs(lpMAPIProp, misses, nullptr, NULL);
				// Cache the results
				AddMapping(missed, sig);
			}

			// Second pass, do our lookup with a populated cache
			for (ULONG ulTarget = 0; ulTarget < lpPropTags->cValues; ulTarget++)
			{
				const auto ulPropId = PROP_ID(lpPropTags->aulPropTag[ulTarget]);
				// ...check the cache
				const auto lpEntry =
					FindCacheEntry([&](const auto& entry) noexcept { return entry->match(sig, ulPropId); });

				if (lpEntry)
				{
					results.emplace_back(lpEntry);
				}
			}

			return results;
		}

		static _Check_return_ std::vector<ULONG> CacheGetIDsFromNames(
			_In_ LPMAPIPROP lpMAPIProp,
			_In_ const std::vector<BYTE>& sig,
			ULONG cPropNames,
			_In_count_(cPropNames) LPMAPINAMEID* lppPropNames,
			ULONG ulFlags)
		{
			if (!lpMAPIProp || !cPropNames || !*lppPropNames) return {};

			// We're going to walk the cache, looking for the values we need. As soon as we have all the values we need, we're done
			// If we reach the end of the cache and don't have everything, we set up to make a GetIDsFromNames call.

			auto misses = std::vector<MAPINAMEID*>{};

			// First pass, find the tags we don't have cached
			for (ULONG ulTarget = 0; ulTarget < cPropNames; ulTarget++)
			{
				const auto lpEntry = FindCacheEntry([&](const auto& entry) noexcept {
					return entry->match(
						sig,
						lppPropNames[ulTarget]->lpguid,
						lppPropNames[ulTarget]->ulKind,
						lppPropNames[ulTarget]->Kind.lID,
						lppPropNames[ulTarget]->Kind.lpwstrName);
				});

				if (!lpEntry)
				{
					misses.emplace_back(lppPropNames[ulTarget]);
				}
			}

			// Go to MAPI with whatever's left.
			if (!misses.empty())
			{
				auto missed = GetIDsFromNames(lpMAPIProp, misses, ulFlags);
				if (misses.size() == missed.size())
				{
					// Cache the results
					auto toCache = std::vector<std::shared_ptr<namedPropCacheEntry>>{};
					for (ULONG i = 0; i < misses.size(); i++)
					{
						toCache.emplace_back(std::make_shared<namedPropCacheEntry>(sig, misses[i], missed[i]));
					}

					AddMapping(toCache, sig);
				}
			}

			auto ids = std::vector<ULONG>{};
			// Second pass, do our lookup with a populated cache
			for (ULONG ulTarget = 0; ulTarget < cPropNames; ulTarget++)
			{
				const auto lpEntry = FindCacheEntry([&](const auto& entry) noexcept {
					return entry->match(
						sig,
						lppPropNames[ulTarget]->lpguid,
						lppPropNames[ulTarget]->ulKind,
						lppPropNames[ulTarget]->Kind.lID,
						lppPropNames[ulTarget]->Kind.lpwstrName);
				});

				if (lpEntry)
				{
					ids.emplace_back(lpEntry->getPropID());
				}
			}

			return ids;
		}
	};

	// No signature form: look up and use signature if possible
	_Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>> GetNamesFromIDs(
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ LPSPropTagArray* lppPropTags,
		_In_opt_ LPGUID lpPropSetGuid,
		ULONG ulFlags)
	{
		auto sig = std::vector<BYTE>{};
		LPSPropValue lpProp = nullptr;
		// This error is too chatty to log - ignore it.
		const auto hRes = HrGetOneProp(lpMAPIProp, PR_MAPPING_SIGNATURE, &lpProp);
		if (SUCCEEDED(hRes) && lpProp && PT_BINARY == PROP_TYPE(lpProp->ulPropTag))
		{
			sig = {lpProp->Value.bin.lpb, lpProp->Value.bin.lpb + lpProp->Value.bin.cb};
		}

		MAPIFreeBuffer(lpProp);

		return GetNamesFromIDs(lpMAPIProp, sig, lppPropTags, lpPropSetGuid, ulFlags);
	}

	// Signature form: if signature not passed then do not use a signature
	_Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>> GetNamesFromIDs(
		_In_ LPMAPIPROP lpMAPIProp,
		_In_opt_ const std::vector<BYTE>& sig,
		_In_ LPSPropTagArray* lppPropTags,
		_In_opt_ LPGUID lpPropSetGuid,
		ULONG ulFlags)
	{
		if (!lpMAPIProp) return {};

		// Check if we're bypassing the cache:
		if (!registry::cacheNamedProps ||
			// Assume an array was passed - none of my calling code passes a NULL tag array
			!lppPropTags || !*lppPropTags ||
			// None of my code uses these flags, but bypass the cache if we see them
			ulFlags || lpPropSetGuid)
		{
			return GetNamesFromIDs(lpMAPIProp, lppPropTags, lpPropSetGuid, ulFlags);
		}

		return namedPropCache::CacheGetNamesFromIDs(lpMAPIProp, sig, lppPropTags);
	}

	_Check_return_ std::vector<ULONG> GetIDsFromNames(
		_In_ LPMAPIPROP lpMAPIProp,
		ULONG cPropNames,
		_In_opt_count_(cPropNames) LPMAPINAMEID* lppPropNames,
		ULONG ulFlags)
	{
		if (!lpMAPIProp) return {};

		auto propTags = std::vector<ULONG>{};
		// Check if we're bypassing the cache:
		if (!registry::cacheNamedProps ||
			// If no names were passed, we have to bypass the cache
			// Should we cache results?
			!cPropNames || !lppPropNames || !*lppPropNames)
		{
			return namedPropCache::GetIDsFromNames(lpMAPIProp, cPropNames, lppPropNames, ulFlags);
		}

		LPSPropValue lpProp = nullptr;

		WC_MAPI_S(HrGetOneProp(lpMAPIProp, PR_MAPPING_SIGNATURE, &lpProp));

		if (lpProp && PT_BINARY == PROP_TYPE(lpProp->ulPropTag))
		{
			auto sig = std::vector<BYTE>{lpProp->Value.bin.lpb, lpProp->Value.bin.lpb + lpProp->Value.bin.cb};
			propTags = namedPropCache::CacheGetIDsFromNames(lpMAPIProp, sig, cPropNames, lppPropNames, ulFlags);
		}
		else
		{
			propTags = namedPropCache::GetIDsFromNames(lpMAPIProp, cPropNames, lppPropNames, ulFlags);
			if (cPropNames == propTags.size())
			{
				// Cache the results
				auto ids = std::vector<std::shared_ptr<namedPropCacheEntry>>{};
				for (ULONG i = 0; i < propTags.size(); i++)
				{
					ids.emplace_back(std::make_shared<namedPropCacheEntry>(lppPropNames[i], propTags[i]));
				}

				namedPropCache::AddMapping(ids, {});
			}
		}

		MAPIFreeBuffer(lpProp);

		return propTags;
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
			lpNamedPropCacheEntry = namedPropCache::FindCacheEntry([&](const auto& entry) noexcept {
				return entry->match(
					PROP_ID(ulPropTag),
					lpNameID->lpguid,
					lpNameID->ulKind,
					lpNameID->Kind.lID,
					lpNameID->Kind.lpwstrName);
			});
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
				L"id/%s/%04X%04X", // STRING_OK
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

			if (cchShortLen < cchWideLen)
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
		_In_opt_ const std::vector<BYTE>& sig, // optional mapping signature for object to speed named prop lookups
		bool bIsAB) // true for an address book property (they can be > 8000 and not named props)
	{
		// If we weren't passed named property information and we need it, look it up
		// We check bIsAB here - some address book providers return garbage which will crash us
		if (!lpNameID && lpMAPIProp && // if we have an object
			!bIsAB && registry::parseNamedProps && // and we're parsing named props
			(registry::getPropNamesOnAllProps ||
			 PROP_ID(ulPropTag) >= 0x8000)) // and it's either a named prop or we're doing all props
		{
			SPropTagArray tag = {};
			auto lpTag = &tag;
			tag.cValues = 1;
			tag.aulPropTag[0] = ulPropTag;

			const auto names = GetNamesFromIDs(lpMAPIProp, sig, &lpTag, nullptr, NULL);
			if (names.size() == 1)
			{
				return NameIDToStrings(names[0]->getMapiNameId(), ulPropTag);
			}
		}

		return {};
	}

	// Returns string built from NameIDArray
	std::vector<std::wstring> NameIDToPropNames(_In_ const MAPINAMEID* lpNameID)
	{
		std::vector<std::wstring> results;
		if (!lpNameID) return {};
		if (lpNameID->ulKind != MNID_ID) return {};
		ULONG ulMatch = cache::ulNoMatch;

		if (NameIDArray.empty()) return {};

		for (ULONG ulCur = 0; ulCur < NameIDArray.size(); ulCur++)
		{
			if (NameIDArray[ulCur].lValue == lpNameID->Kind.lID)
			{
				ulMatch = ulCur;
				break;
			}
		}

		if (cache::ulNoMatch != ulMatch)
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
} // namespace cache2