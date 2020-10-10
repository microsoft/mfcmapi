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
		_Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>>
		GetRange(_In_ LPMAPIPROP lpMAPIProp, ULONG start, ULONG end)
		{
			if (start > end) return {};
			// Allocate our tag array
			ULONG count = end - start + 1;
			auto lpTag = mapi::allocate<LPSPropTagArray>(CbNewSPropTagArray(count));
			if (lpTag)
			{
				// Populate the array
				lpTag->cValues = count;
				for (ULONG tag = start, i = 0; tag <= end; tag++, i++)
				{
					mapi::setTag(lpTag, i) = PROP_TAG(PT_NULL, tag);
				}

				auto names = GetNamesFromIDs(lpMAPIProp, &lpTag, NULL);
				MAPIFreeBuffer(lpTag);
				return names;
			}

			return {};
		}

		_Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>> GetAllNamesFromIDs(_In_ LPMAPIPROP lpMAPIProp)
		{
			LPSPropTagArray pProps = nullptr;
			auto names = std::vector<std::shared_ptr<namedPropCacheEntry>>{};

			// We didn't get any names - try manual
			const auto ulLowerBound = __LOWERBOUND;
			const auto ulUpperBound = FindHighestNamedProp(lpMAPIProp);
			const auto batchSize = 200;

			output::DebugPrint(
				output::dbgLevel::NamedProp,
				L"GetAllNamesFromIDs: Walking through all IDs from 0x%X to 0x%X, looking for mappings to names\n",
				ulLowerBound,
				ulUpperBound);
			for (auto iTag = ulLowerBound; iTag <= ulUpperBound && iTag < __UPPERBOUND; iTag += batchSize - 1)
			{
				const auto range = GetRange(lpMAPIProp, iTag, min(iTag + batchSize - 1, ulUpperBound));
				for (const auto& name : range)
				{
					if (name->getMapiNameId()->lpguid != nullptr)
						if (name->getMapiNameId()->lpguid != nullptr)
						{
							names.push_back(name);
						}
				}
			}

			return names;
		}

		// Returns a vector of NamedPropCacheEntry for the input tags
		// Sourced directly from MAPI
		_Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>>
		GetNamesFromIDs(_In_ LPMAPIPROP lpMAPIProp, _In_opt_ LPSPropTagArray* lppPropTags, ULONG ulFlags)
		{
			if (!lpMAPIProp) return {};

			LPMAPINAMEID* lppPropNames = nullptr;
			auto ulPropNames = ULONG{};

			// Try a direct call first
			const auto hRes =
				WC_H_GETPROPS(lpMAPIProp->GetNamesFromIDs(lppPropTags, nullptr, ulFlags, &ulPropNames, &lppPropNames));

			// If we failed and we were doing an all props lookup, try it manually instead
			if (hRes == MAPI_E_CALL_FAILED && (!lppPropTags || !*lppPropTags))
			{
				return GetAllNamesFromIDs(lpMAPIProp);
			}

			auto ids = std::vector<std::shared_ptr<namedPropCacheEntry>>{};

			if (ulPropNames && lppPropNames)
			{
				for (ULONG i = 0; i < ulPropNames; i++)
				{
					auto ulPropID = ULONG{};
					if (lppPropTags && *lppPropTags) ulPropID = PROP_ID(mapi::getTag(*lppPropTags, i));
					ids.emplace_back(namedPropCacheEntry::make(lppPropNames[i], ulPropID));
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

			auto countTags = ULONG(tags.size());
			auto ids = std::vector<std::shared_ptr<namedPropCacheEntry>>{};
			auto ulPropTags = mapi::allocate<LPSPropTagArray>(CbNewSPropTagArray(countTags));
			if (ulPropTags)
			{
				ulPropTags->cValues = countTags;
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
				WC_H_GETPROPS_S(lpMAPIProp->GetIDsFromNames(ULONG(lpNameIDs.size()), names, ulFlags, &lpTags));
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

		if (entry != cache.end())
		{
			output::DebugPrint(output::dbgLevel::NamedPropCache, L"find: found match\n");
			return *entry;
		}
		else
		{
			output::DebugPrint(output::dbgLevel::NamedPropCache, L"find: no match\n");
			return namedPropCacheEntry::empty();
		}
	}

	_Check_return_ std::shared_ptr<namedPropCacheEntry> namedPropCache::find(
		const std::shared_ptr<cache::namedPropCacheEntry>& entry,
		bool bMatchSig,
		bool bMatchID,
		bool bMatchName)
	{
		if (fIsSet(output::dbgLevel::NamedPropCache))
		{
			output::DebugPrint(
				output::dbgLevel::NamedPropCache,
				L"find: bMatchSig=%d, bMatchID=%d, bMatchName=%d\n",
				bMatchSig,
				bMatchID,
				bMatchName);
			entry->output();
		}

		return find([&](const auto& _entry) { return _entry->match(entry, bMatchSig, bMatchID, bMatchName); });
	}

	_Check_return_ std::shared_ptr<namedPropCacheEntry>
	namedPropCache::find(_In_ const std::vector<BYTE>& _sig, _In_ const MAPINAMEID& _mapiNameId)
	{
		if (fIsSet(output::dbgLevel::NamedPropCache))
		{
			const auto nameidString = strings::MAPINAMEIDToString(_mapiNameId);
			if (_sig.empty())
			{
				output::DebugPrint(output::dbgLevel::NamedPropCache, L"find: _mapiNameId: %ws\n", nameidString.c_str());
			}
			else
			{
				const auto sigStr = strings::BinToHexString(_sig, true);
				output::DebugPrint(
					output::dbgLevel::NamedPropCache,
					L"find: _sig=%ws, _mapiNameId: %ws\n",
					sigStr.c_str(),
					nameidString.c_str());
			}
		}

		return find([&](const auto& _entry) { return _entry->match(_sig, _mapiNameId); });
	}

	_Check_return_ std::shared_ptr<namedPropCacheEntry>
	namedPropCache::find(_In_ const std::vector<BYTE>& _sig, ULONG _ulPropID)
	{
		if (fIsSet(output::dbgLevel::NamedPropCache))
		{
			if (_sig.empty())
			{
				output::DebugPrint(output::dbgLevel::NamedPropCache, L"find: _ulPropID=%04X\n", _ulPropID);
			}
			else
			{
				const auto sigStr = strings::BinToHexString(_sig, true);
				output::DebugPrint(
					output::dbgLevel::NamedPropCache, L"find: _sig=%ws, _ulPropID=%04X\n", sigStr.c_str(), _ulPropID);
			}
		}

		return find([&](const auto& _entry) { return _entry->match(_sig, _ulPropID); });
	}

	_Check_return_ std::shared_ptr<namedPropCacheEntry>
	namedPropCache::find(ULONG _ulPropID, _In_ const MAPINAMEID& _mapiNameId)
	{
		if (fIsSet(output::dbgLevel::NamedPropCache))
		{
			const auto nameidString = strings::MAPINAMEIDToString(_mapiNameId);
			output::DebugPrint(
				output::dbgLevel::NamedPropCache,
				L"find: _ulPropID=%04X, _mapiNameId: %ws\n",
				_ulPropID,
				nameidString.c_str());
		}

		return find([&](const auto& _entry) { return _entry->match(_ulPropID, _mapiNameId); });
	}

	// Add a mapping to the cache if it doesn't already exist
	// If given a signature, we include it in our search.
	// If not, we search without it
	void
	namedPropCache::add(const std::vector<std::shared_ptr<namedPropCacheEntry>>& entries, const std::vector<BYTE>& sig)
	{
		auto& cache = getCache();
		for (auto& entry : entries)
		{
			output::DebugPrint(output::dbgLevel::NamedPropCache, L"add:\n");
			entry->output();
			if (entry->getPropID() == PT_ERROR)
			{
				output::DebugPrint(output::dbgLevel::NamedPropCache, L"add: Skipping error property\n");
				continue;
			}

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

			if (!namedPropCacheEntry::valid(match))
			{
				if (fIsSet(output::dbgLevel::NamedPropCache))
				{
					const auto names = NameIDToPropNames(entry->getMapiNameId());
					if (names.empty())
					{
						output::DebugPrint(output::dbgLevel::NamedPropCache, L"add: Caching unknown property\n");
						entry->output();
					}
					else
					{
						output::DebugPrint(
							output::dbgLevel::NamedPropCache, L"add: Caching property %ws\n", names[0].c_str());
						entry->output();
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
		_In_opt_ LPSPropTagArray* lppPropTags)
	{
		if (!lpMAPIProp) return {};

		// If this is a get all names call, we have to go direct to MAPI since we cannot trust the cache is full.
		if (!lppPropTags || !*lppPropTags)
		{
			output::DebugPrint(output::dbgLevel::NamedPropCache, L"GetNamesFromIDs: making direct all for all props\n");
			LPSPropTagArray pProps = nullptr;
			auto names = directMapi::GetNamesFromIDs(lpMAPIProp, &pProps, NULL);
			// Cache the results
			add(names, sig);
			return names;
		}

		// We're going to walk the cache, looking for the values we need. As soon as we have all the values we need, we're done
		// If we reach the end of the cache and don't have everything, we set up to make a GetNamesFromIDs call.

		auto results = std::vector<std::shared_ptr<namedPropCacheEntry>>{};
		const SPropTagArray* lpPropTags = *lppPropTags;

		auto misses = std::vector<ULONG>{};

		// First pass, find any misses we might have
		output::DebugPrint(output::dbgLevel::NamedPropCache, L"GetNamesFromIDs: Looking for misses\n");
		for (ULONG ulTarget = 0; ulTarget < lpPropTags->cValues; ulTarget++)
		{
			const auto ulPropTag = mapi::getTag(lpPropTags, ulTarget);
			const auto ulPropId = PROP_ID(ulPropTag);
			// ...check the cache
			const auto lpEntry = find(sig, ulPropId);

			if (!namedPropCacheEntry::valid(lpEntry))
			{
				misses.emplace_back(ulPropTag);
			}
		}

		// Go to MAPI with whatever's left. We set up for a single call to GetNamesFromIDs.
		if (!misses.empty())
		{
			output::DebugPrint(
				output::dbgLevel::NamedPropCache, L"GetNamesFromIDs: Add %d misses to cache\n", misses.size());
			auto missed = directMapi::GetNamesFromIDs(lpMAPIProp, misses, NULL);
			// Cache the results
			add(missed, sig);
		}

		// Second pass, do our lookup with a populated cache
		output::DebugPrint(output::dbgLevel::NamedPropCache, L"GetNamesFromIDs: Lookup again from cache\n");
		for (ULONG ulTarget = 0; ulTarget < lpPropTags->cValues; ulTarget++)
		{
			const auto ulPropId = PROP_ID(mapi::getTag(lpPropTags, ulTarget));
			// ...check the cache
			const auto lpEntry = find(sig, ulPropId);

			if (namedPropCacheEntry::valid(lpEntry))
			{
				results.emplace_back(lpEntry);
			}
			else
			{
				results.emplace_back(namedPropCacheEntry::make(nullptr, ulPropId, sig));
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
		output::DebugPrint(output::dbgLevel::NamedPropCache, L"GetIDsFromNames: Looking for misses\n");
		for (const auto& nameID : nameIDs)
		{
			const auto lpEntry = find(sig, nameID);

			if (!namedPropCacheEntry::valid(lpEntry))
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
					toCache.emplace_back(namedPropCacheEntry::make(&misses[i], mapi::getTag(missed, i), sig));
				}

				output::DebugPrint(
					output::dbgLevel::NamedPropCache, L"GetIDsFromNames: Add %d misses to cache\n", misses.size());
				add(toCache, sig);
			}

			MAPIFreeBuffer(missed);
		}

		// Second pass, do our lookup with a populated cache
		auto countIDs = ULONG(nameIDs.size());
		output::DebugPrint(output::dbgLevel::NamedPropCache, L"GetIDsFromNames: Lookup again from cache\n");
		auto results = mapi::allocate<LPSPropTagArray>(CbNewSPropTagArray(countIDs));
		if (results)
		{
			results->cValues = countIDs;
			ULONG i = 0;
			for (const auto& nameID : nameIDs)
			{
				const auto lpEntry = find(sig, nameID);

				mapi::setTag(results, i++) = namedPropCacheEntry::valid(lpEntry) ? lpEntry->getPropID() : 0;
			}
		}

		return results;
	}
} // namespace cache