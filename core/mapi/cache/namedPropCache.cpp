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
		_Check_return_ HRESULT GetRange(
			_In_ LPMAPIPROP lpMAPIProp,
			ULONG start,
			ULONG end,
			std::vector<std::shared_ptr<namedPropCacheEntry>> &names)
		{
			if (start > end) return {};
			// Allocate our tag array
			const auto count = end - start + 1;
			auto lpTag = mapi::allocate<LPSPropTagArray>(CbNewSPropTagArray(count));
			if (lpTag)
			{
				// Populate the array
				lpTag->cValues = count;
				for (ULONG tag = start, i = 0; tag <= end; tag++, i++)
				{
					mapi::setTag(lpTag, i) = PROP_TAG(PT_NULL, tag);
				}

				const HRESULT hRes = WC_H(GetNamesFromIDs(lpMAPIProp, &lpTag, NULL, names));
				MAPIFreeBuffer(lpTag);
				return hRes;
			}

			return MAPI_E_NOT_ENOUGH_MEMORY;
		}

		_Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>> GetAllNamesFromIDs(_In_ LPMAPIFOLDER lpMAPIFolder)
		{
			auto names = std::vector<std::shared_ptr<namedPropCacheEntry>>{};

			// We didn't get any names - try manual
			constexpr auto ulLowerBound = __LOWERBOUND;
			const auto ulUpperBound = FindHighestNamedProp(lpMAPIFolder);
			ULONG batchSize = registry::namedPropBatchSize;

			output::DebugPrint(
				output::dbgLevel::NamedProp,
				L"GetAllNamesFromIDs: Walking through all IDs from 0x%X to 0x%X, looking for mappings to names\n",
				ulLowerBound,
				ulUpperBound);

			auto iTag = ulLowerBound;
			while (iTag <= ulUpperBound && iTag < __UPPERBOUND)
			{
				auto end = min(iTag + batchSize - 1, ulUpperBound);
				std::vector<std::shared_ptr<namedPropCacheEntry>> range;
				range.reserve(batchSize);

				// Trying to get a range of these props can fail with MAPI_E_CALL_FAILED if it's too big for the buffer
				// In this scenario, reopen the object, lower the batch size, and try again
				HRESULT hRes = WC_H(GetRange(lpMAPIFolder, iTag, end, range));
				if (hRes == MAPI_E_CALL_FAILED)
				{
					LPMAPIFOLDER newFolder = nullptr;
					ULONG objt = 0;
					hRes = WC_H(lpMAPIFolder->OpenEntry(
						0, nullptr, &IID_IMAPIFolder, 0, &objt, reinterpret_cast<LPUNKNOWN*>(&newFolder)));

					// If we can't re-open the root folder, something is bad - just break out
					if (FAILED(hRes))
					{
						break;
					}

					// Swap the folder, lower the batch size, and try again
					lpMAPIFolder->Release();
					lpMAPIFolder = newFolder;
					batchSize /= 2;
					continue;
				}
				else if (FAILED(hRes))
				{
					// Other errors are unexpected
					break;
				}

				for (const auto& name : range)
				{
					if (name->getMapiNameId()->lpguid != nullptr)
					{
						names.push_back(name);
					}
				}

				// Go to the next batch
				iTag += batchSize - 1;
			}

			lpMAPIFolder->Release();
			return names;
		}

		_Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>>
		GetAllNamesFromIDsFromContainer(LPMAPICONTAINER lpMAPIContainer)
		{
			LPMAPIFOLDER lpRootFolder = nullptr;
			if (FAILED(GetRootFolder(lpMAPIContainer, &lpRootFolder)))
			{
				return {};
			}

			return GetAllNamesFromIDs(lpRootFolder);
		}

		_Check_return_ std::vector<std::shared_ptr<namedPropCacheEntry>> GetAllNamesFromIDsFromMdb(LPMDB lpMdb)
		{
			LPMAPIFOLDER lpRootFolder = nullptr;
			if (FAILED(GetRootFolder(lpMdb, &lpRootFolder)))
			{
				return {};
			}

			return GetAllNamesFromIDs(lpRootFolder);
		}

		// Returns a vector of NamedPropCacheEntry for the input tags
		// Sourced directly from MAPI
		_Check_return_ HRESULT GetNamesFromIDs(
			_In_ LPMAPIPROP lpMAPIProp,
			_In_opt_ LPSPropTagArray* lppPropTags,
			ULONG ulFlags,
			std::vector<std::shared_ptr<namedPropCacheEntry>> &names)
		{
			if (!lpMAPIProp)
			{
				return MAPI_E_INVALID_PARAMETER;
			}

			LPMAPINAMEID* lppPropNames = nullptr;
			auto ulPropNames = ULONG{};

			// Try a direct call first
			const auto hRes =
				WC_H_GETPROPS(lpMAPIProp->GetNamesFromIDs(lppPropTags, nullptr, ulFlags, &ulPropNames, &lppPropNames));

			// If we failed and we were doing an all props lookup, try it manually instead
			if (hRes == MAPI_E_CALL_FAILED && (!*lppPropTags))
			{
				LPMAPICONTAINER lpContainer = nullptr;
				HRESULT hRes2 = WC_H(lpMAPIProp->QueryInterface(IID_IMAPIContainer, reinterpret_cast<LPVOID*>(&lpContainer)));
				if (SUCCEEDED(hRes2))
				{
					names = GetAllNamesFromIDsFromContainer(lpContainer);
					// REVIEW: Doesn't QueryInterface effectively AddRef?
					lpContainer->Release();
					return S_OK;
				}

				LPMDB lpMdb = nullptr;
				hRes2 = WC_H(lpMAPIProp->QueryInterface(IID_IMsgStore, reinterpret_cast<LPVOID*>(&lpMdb)));
				if (SUCCEEDED(hRes2))
				{
					names = GetAllNamesFromIDsFromMdb(lpMdb);
					// REVIEW: Doesn't QueryInterface effectively AddRef?
					lpMdb->Release();
					return S_OK;
				}

				// Can't do the get all props special route because object wasn't IID_IMAPIContainer
				return hRes;
			}

			if (ulPropNames && lppPropNames)
			{
				for (ULONG i = 0; i < ulPropNames; i++)
				{
					auto ulPropID = ULONG{};
					if (*lppPropTags) ulPropID = PROP_ID(mapi::getTag(*lppPropTags, i));
					names.emplace_back(namedPropCacheEntry::make(lppPropNames[i], ulPropID));
				}
			}

			MAPIFreeBuffer(lppPropNames);
			return hRes;
		}

		// Returns a vector of NamedPropCacheEntry for the input tags
		// Sourced directly from MAPI
		_Check_return_ HRESULT GetNamesFromIDs(
			_In_ LPMAPIPROP lpMAPIProp,
			_In_ const std::vector<ULONG> tags,
			ULONG ulFlags,
			std::vector<std::shared_ptr<namedPropCacheEntry>> &names)
		{
			if (!lpMAPIProp)
			{
				return MAPI_E_INVALID_PARAMETER;
			}

			auto countTags = ULONG(tags.size());
			auto ulPropTags = mapi::allocate<LPSPropTagArray>(CbNewSPropTagArray(countTags));
			if (ulPropTags)
			{
				ulPropTags->cValues = countTags;
				ULONG i = 0;
				for (const auto& tag : tags)
				{
					mapi::setTag(ulPropTags, i++) = tag;
				}

				return WC_H(GetNamesFromIDs(lpMAPIProp, &ulPropTags, ulFlags, names));
			}

			MAPIFreeBuffer(ulPropTags);

			return MAPI_E_NOT_ENOUGH_MEMORY;
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
		_In_ const std::vector<BYTE>& sig,
		_In_opt_ LPSPropTagArray* lppPropTags)
	{
		if (!lpMAPIProp) return {};

		// If this is a get all names call, we have to go direct to MAPI since we cannot trust the cache is full.
		if (!*lppPropTags)
		{
			output::DebugPrint(output::dbgLevel::NamedPropCache, L"GetNamesFromIDs: making direct all for all props\n");
			LPSPropTagArray pProps = nullptr;

			std::vector<std::shared_ptr<namedPropCacheEntry>> names;
			WC_H_S(directMapi::GetNamesFromIDs(lpMAPIProp, &pProps, NULL, names));

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
			std::vector<std::shared_ptr<namedPropCacheEntry>> missed;
			missed.reserve(misses.size());
			WC_H_S(directMapi::GetNamesFromIDs(lpMAPIProp, misses, NULL, missed));

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