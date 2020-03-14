#include <core/stdafx.h>
#include <core/mapi/cache/namedPropCache.h>
#include <core/mapi/cache/namedProps.h>
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
	namespace directMapi
	{
		// Returns a vector of NamedPropCacheEntry for the input tags
		// Sourced directly from MAPI
		_Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>>
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
					if (lppPropTags && *lppPropTags) ulPropID = PROP_ID(mapi::getTag(*lppPropTags, i));
					ids.emplace_back(std::make_shared<namedPropCacheEntry>(lppPropNames[i], ulPropID));
				}
			}

			MAPIFreeBuffer(lppPropNames);
			return ids;
		}

		// Returns a vector of NamedPropCacheEntry for the input tags
		// Sourced directly from MAPI
		_Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>>
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
					mapi::setTag(ulPropTags, i++) = tag;
				}

				ids = GetNamesFromIDs(lpMAPIProp, &ulPropTags, ulFlags);
			}

			MAPIFreeBuffer(ulPropTags);

			return ids;
		}

		// Returns a vector of tags for the input names
		// Sourced directly from MAPI
		_Check_return_ LPSPropTagArray
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

	std::list<std::shared_ptr<namedPropCacheEntry>>& namedPropCache::getCache() noexcept
	{
		// We keep a list of named prop cache entries
		static std::list<std::shared_ptr<namedPropCacheEntry>> cache;
		return cache;
	}

	_Check_return_ std::shared_ptr<namedPropCacheEntry>
	namedPropCache::find(const std::function<bool(const std::shared_ptr<namedPropCacheEntry>&)>& compare)
	{
		const auto& cache = getCache();
		const auto entry =
			find_if(cache.begin(), cache.end(), [compare](const auto& _entry) { return compare(_entry); });

		return entry != cache.end() ? *entry : nullptr;
	}

	_Check_return_ std::shared_ptr<namedPropCacheEntry> namedPropCache::find(
		const std::shared_ptr<cache::namedPropCacheEntry>& entry,
		bool bMatchSig,
		bool bMatchID,
		bool bMatchName)
	{
		return find([&](const auto& _entry) { return _entry->match(*entry.get(), bMatchSig, bMatchID, bMatchName); });
	}

	_Check_return_ std::shared_ptr<namedPropCacheEntry>
	namedPropCache::find(_In_ const std::vector<BYTE>& _sig, _In_ const MAPINAMEID& _mapiNameId)
	{
		return find([&](const auto& _entry) { return _entry->match(_sig, _mapiNameId); });
	}

	_Check_return_ std::shared_ptr<namedPropCacheEntry>
	namedPropCache::find(_In_ const std::vector<BYTE>& _sig, ULONG _ulPropID)
	{
		return find([&](const auto& _entry) { return _entry->match(_sig, _ulPropID); });
	}

	_Check_return_ std::shared_ptr<namedPropCacheEntry>
	namedPropCache::find(ULONG _ulPropID, _In_ const MAPINAMEID& _mapiNameId)
	{
		return find([&](const auto& _entry) { return _entry->match(_ulPropID, _mapiNameId); });
	}

	// Add a mapping to the cache if it doesn't already exist
	// If given a signature, we include it in our search.
	// If not, we search without it
	void namedPropCache::add(std::vector<std::shared_ptr<namedPropCacheEntry>>& entries, const std::vector<BYTE>& sig)
	{
		auto& cache = getCache();
		for (auto& entry : entries)
		{
			auto match = std::shared_ptr<namedPropCacheEntry>{};
			if (sig.empty())
			{
				match = find(entry, false, true, true);
			}
			else
			{
				entry->setSig(sig);
				match = find(entry, true, true, true);
			}

			if (!match)
			{
				if (fIsSet(output::dbgLevel::NamedPropCacheMisses))
				{
					const auto mni = entry->getMapiNameId();
					const auto names = NameIDToPropNames(mni);
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
	_Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>> namedPropCache::GetNamesFromIDs(
		_In_ LPMAPIPROP lpMAPIProp,
		const std::vector<BYTE>& sig,
		_In_ LPSPropTagArray* lppPropTags)
	{
		if (!lpMAPIProp || !lppPropTags || !*lppPropTags) return {};

		// We're going to walk the cache, looking for the values we need. As soon as we have all the values we need, we're done
		// If we reach the end of the cache and don't have everything, we set up to make a GetNamesFromIDs call.

		auto results = std::vector<std::shared_ptr<namedPropCacheEntry>>{};
		const SPropTagArray* lpPropTags = *lppPropTags;

		auto misses = std::vector<ULONG>{};

		// First pass, find any misses we might have
		for (ULONG ulTarget = 0; ulTarget < lpPropTags->cValues; ulTarget++)
		{
			const auto ulPropTag = mapi::getTag(lpPropTags, ulTarget);
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
			const auto ulPropId = PROP_ID(mapi::getTag(lpPropTags, ulTarget));
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
	_Check_return_ LPSPropTagArray namedPropCache::GetIDsFromNames(
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
			const auto lpEntry = find([&](const auto& entry) noexcept { return entry->match(sig, nameID); });

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
						std::make_shared<namedPropCacheEntry>(&misses[i], mapi::getTag(missed, i), sig));
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
			const auto lpEntry = find([&](const auto& entry) noexcept { return entry->match(sig, nameID); });

			mapi::setTag(results, i++) = lpEntry ? lpEntry->getPropID() : 0;
		}

		return results;
	}
} // namespace cache