#include <core/stdafx.h>
#include <core/mapi/cache/namedPropCacheEntry.h>
#include <core/mapi/cache/namedPropCache.h>
#include <core/interpret/guid.h>
#include <core/mapi/mapiMemory.h>
#include <core/utility/registry.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/addin/mfcmapi.h>
#include <core/addin/addin.h>
#include <core/utility/error.h>

namespace cache
{
	namespace directMapi
	{
		// Returns a vector of NamedPropCacheEntry for the input tags
		// Sourced directly from MAPI
		static _Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>>
		GetNamesFromIDs(_In_ LPMAPIPROP lpMAPIProp, _In_ LPSPropTagArray* lppPropTags, ULONG ulFlags)
		{
			if (!lpMAPIProp) return {};

			LPMAPINAMEID* lppPropNames = nullptr;
			auto ulPropNames = ULONG{};

			WC_H_GETPROPS_S(lpMAPIProp->GetNamesFromIDs(lppPropTags, nullptr, ulFlags, &ulPropNames, &lppPropNames));

			auto ids = std::vector<std::shared_ptr<namedPropCacheEntry>>{};

			if (ulPropNames && lppPropNames)
			{
				for (ULONG i = 0; i < ulPropNames; i++)
				{
					auto ulPropID = ULONG{};
					if (lppPropTags && *lppPropTags) ulPropID = PROP_ID((*lppPropTags)->aulPropTag[i]);
					ids.emplace_back(std::make_shared<namedPropCacheEntry>(lppPropNames[i], ulPropID));
					// TODO: Figure out what misses look like here
					// lppPropNames[i]* will be null...
				}
			}

			MAPIFreeBuffer(lppPropNames);
			return ids;
		}

		// Returns a vector of NamedPropCacheEntry for the input tags
		// Sourced directly from MAPI
		static _Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>>
		GetNamesFromIDs(_In_ LPMAPIPROP lpMAPIProp, _In_ const std::vector<ULONG> tags, ULONG ulFlags)
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

				ids = GetNamesFromIDs(lpMAPIProp, &ulPropTags, ulFlags);
			}

			MAPIFreeBuffer(ulPropTags);

			return ids;
		}

		// Returns a vector of tags for the input names
		// Sourced directly from MAPI
		static _Check_return_ LPSPropTagArray
		GetIDsFromNames(_In_ LPMAPIPROP lpMAPIProp, std::vector<MAPINAMEID> nameIDs, ULONG ulFlags)
		{
			if (!lpMAPIProp) return {};

			LPSPropTagArray lpTags = nullptr;

			if (nameIDs.empty())
			{
				WC_H_GETPROPS_S(lpMAPIProp->GetIDsFromNames(0, nullptr, ulFlags, &lpTags));
			}
			else
			{
				std::vector<const MAPINAMEID*> lpNameIDs = {};
				for (const auto& nameID : nameIDs)
				{
					lpNameIDs.emplace_back(&nameID);
				}

				auto names = const_cast<MAPINAMEID**>(lpNameIDs.data());
				WC_H_GETPROPS_S(lpMAPIProp->GetIDsFromNames(lpNameIDs.size(), names, ulFlags, &lpTags));
			}

			return lpTags;
		}
	} // namespace directMapi

	class namedPropCache
	{
	public:
		static std::list<std::shared_ptr<namedPropCacheEntry>>& getCache() noexcept
		{
			// We keep a list of named prop cache entries
			static std::list<std::shared_ptr<namedPropCacheEntry>> cache;
			return cache;
		}

		_Check_return_ static std::shared_ptr<namedPropCacheEntry>
		find(const std::function<bool(const std::shared_ptr<namedPropCacheEntry>&)>& compare)
		{
			const auto& cache = getCache();
			const auto entry =
				find_if(cache.begin(), cache.end(), [compare](const auto& _entry) { return compare(_entry); });

			return entry != cache.end() ? *entry : nullptr;
		}

		// Add a mapping to the cache if it doesn't already exist
		// If given a signature, we include it in our search.
		// If not, we search without it
		static void add(std::vector<std::shared_ptr<namedPropCacheEntry>>& entries, const std::vector<BYTE>& sig)
		{
			auto& cache = getCache();
			for (auto& entry : entries)
			{
				auto match = std::shared_ptr<namedPropCacheEntry>{};
				if (sig.empty())
				{
					match = find([&](const auto& _entry) { return entry->match(_entry.get(), false, true, true); });
				}
				else
				{
					entry->setSig(sig);
					match = find([&](const auto& _entry) { return entry->match(_entry.get(), true, true, true); });
				}

				if (!match)
				{
					if (fIsSet(output::dbgLevel::NamedPropCacheMisses))
					{
						const auto mni = entry->getMapiNameId();
						auto names = NameIDToPropNames(mni);
						if (names.empty())
						{
							output::DebugPrint(
								output::dbgLevel::NamedPropCacheMisses,
								L"add: Caching unknown property 0x%08X %ws\n",
								mni->Kind.lID,
								guid::GUIDToStringAndName(mni->lpguid).c_str());
						}
						else
						{
							output::DebugPrint(
								output::dbgLevel::NamedPropCacheMisses,
								L"add: Caching property 0x%08X %ws = %ws\n",
								mni->Kind.lID,
								guid::GUIDToStringAndName(mni->lpguid).c_str(),
								names[0].c_str());
						}
					}

					cache.emplace_back(entry);
				}
			}
		}

		// If signature is empty then do not use a signature
		_Check_return_ static std::vector<std::shared_ptr<namedPropCacheEntry>>
		GetNamesFromIDs(_In_ LPMAPIPROP lpMAPIProp, const std::vector<BYTE>& sig, _In_ LPSPropTagArray* lppPropTags)
		{
			if (!lpMAPIProp || !lppPropTags || !*lppPropTags) return {};

			// We're going to walk the cache, looking for the values we need. As soon as we have all the values we need, we're done
			// If we reach the end of the cache and don't have everything, we set up to make a GetNamesFromIDs call.

			auto results = std::vector<std::shared_ptr<namedPropCacheEntry>>{};
			const auto lpPropTags = *lppPropTags;

			auto misses = std::vector<ULONG>{};

			// First pass, find any misses we might have
			for (ULONG ulTarget = 0; ulTarget < lpPropTags->cValues; ulTarget++)
			{
				const auto ulPropTag = lpPropTags->aulPropTag[ulTarget];
				const auto ulPropId = PROP_ID(ulPropTag);
				// ...check the cache
				const auto lpEntry = find([&](const auto& entry) noexcept { return entry->match(sig, ulPropId); });

				if (!lpEntry)
				{
					misses.emplace_back(ulPropTag);
				}
			}

			// Go to MAPI with whatever's left. We set up for a single call to GetNamesFromIDs.
			if (!misses.empty())
			{
				auto missed = directMapi::GetNamesFromIDs(lpMAPIProp, misses, NULL);
				// Cache the results
				add(missed, sig);
			}

			// Second pass, do our lookup with a populated cache
			for (ULONG ulTarget = 0; ulTarget < lpPropTags->cValues; ulTarget++)
			{
				const auto ulPropId = PROP_ID(lpPropTags->aulPropTag[ulTarget]);
				// ...check the cache
				const auto lpEntry = find([&](const auto& entry) noexcept { return entry->match(sig, ulPropId); });

				if (lpEntry)
				{
					results.emplace_back(lpEntry);
				}
				else
				{
					results.emplace_back(std::make_shared<namedPropCacheEntry>(nullptr, ulPropId, sig));
				}
			}

			return results;
		}

		// If signature is empty then do not use a signature
		static _Check_return_ LPSPropTagArray GetIDsFromNames(
			_In_ LPMAPIPROP lpMAPIProp,
			_In_ const std::vector<BYTE>& sig,
			_In_ std::vector<MAPINAMEID> nameIDs,
			ULONG ulFlags)
		{
			if (!lpMAPIProp || !nameIDs.size()) return {};

			// We're going to walk the cache, looking for the values we need. As soon as we have all the values we need, we're done
			// If we reach the end of the cache and don't have everything, we set up to make a GetIDsFromNames call.

			auto misses = std::vector<MAPINAMEID>{};

			// First pass, find the tags we don't have cached
			for (const auto& nameID : nameIDs)
			{
				const auto lpEntry = find([&](const auto& entry) noexcept {
					return entry->match(sig, nameID.lpguid, nameID.ulKind, nameID.Kind.lID, nameID.Kind.lpwstrName);
				});

				if (!lpEntry)
				{
					misses.emplace_back(nameID);
				}
			}

			// Go to MAPI with whatever's left.
			if (!misses.empty())
			{
				auto missed = directMapi::GetIDsFromNames(lpMAPIProp, misses, ulFlags);
				if (missed && missed->cValues == misses.size())
				{
					// Cache the results
					auto toCache = std::vector<std::shared_ptr<namedPropCacheEntry>>{};
					for (ULONG i = 0; i < misses.size(); i++)
					{
						toCache.emplace_back(
							std::make_shared<namedPropCacheEntry>(&misses[i], missed->aulPropTag[i], sig));
					}

					add(toCache, sig);
				}

				MAPIFreeBuffer(missed);
			}

			// Second pass, do our lookup with a populated cache
			auto results = mapi::allocate<LPSPropTagArray>(CbNewSPropTagArray(nameIDs.size()));
			results->cValues = nameIDs.size();
			ULONG i = 0;
			for (const auto nameID : nameIDs)
			{
				const auto lpEntry = find([&](const auto& entry) noexcept {
					return entry->match(sig, nameID.lpguid, nameID.ulKind, nameID.Kind.lID, nameID.Kind.lpwstrName);
				});

				results->aulPropTag[i++] = lpEntry ? lpEntry->getPropID() : 0;
			}

			return results;
		}
	};

	_Check_return_ std::shared_ptr<namedPropCacheEntry>
	GetNameFromID(_In_ LPMAPIPROP lpMAPIProp, _In_ ULONG ulPropTag, ULONG ulFlags)
	{
		auto tag = SPropTagArray{1, ulPropTag};
		auto lptag = &tag;
		const auto names = GetNamesFromIDs(lpMAPIProp, &lptag, ulFlags);
		if (names.size() == 1) return names[0];
		return {};
	}

	_Check_return_ std::shared_ptr<namedPropCacheEntry> GetNameFromID(
		_In_ LPMAPIPROP lpMAPIProp,
		_In_opt_ const std::vector<BYTE>& sig,
		_In_ ULONG ulPropTag,
		ULONG ulFlags)
	{
		auto tag = SPropTagArray{1, ulPropTag};
		auto lptag = &tag;
		const auto names = GetNamesFromIDs(lpMAPIProp, sig, &lptag, ulFlags);
		if (names.size() == 1) return names[0];
		return {};
	}

	// No signature form: look up and use signature if possible
	_Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>>
	GetNamesFromIDs(_In_ LPMAPIPROP lpMAPIProp, _In_ LPSPropTagArray* lppPropTags, ULONG ulFlags)
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

		return GetNamesFromIDs(lpMAPIProp, sig, lppPropTags, ulFlags);
	}

	// Signature form: if signature is empty then do not use a signature
	_Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>> GetNamesFromIDs(
		_In_ LPMAPIPROP lpMAPIProp,
		_In_opt_ const std::vector<BYTE>& sig,
		_In_ LPSPropTagArray* lppPropTags,
		ULONG ulFlags)
	{
		if (!lpMAPIProp) return {};

		// Check if we're bypassing the cache:
		if (!registry::cacheNamedProps ||
			// Assume an array was passed - none of my calling code passes a NULL tag array
			!lppPropTags || !*lppPropTags ||
			// None of my code uses these flags, but bypass the cache if we see them
			ulFlags)
		{
			return directMapi::GetNamesFromIDs(lpMAPIProp, lppPropTags, ulFlags);
		}

		return namedPropCache::GetNamesFromIDs(lpMAPIProp, sig, lppPropTags);
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
			sig = std::vector<BYTE>{lpProp->Value.bin.lpb, lpProp->Value.bin.lpb + lpProp->Value.bin.cb};
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
			lpNamedPropCacheEntry = namedPropCache::find([&](const auto& entry) noexcept {
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
			const auto name = GetNameFromID(lpMAPIProp, sig, ulPropTag, NULL);
			if (name->valid())
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
} // namespace cache