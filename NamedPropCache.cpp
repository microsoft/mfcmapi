#include "stdafx.h"
#include "NamedPropCache.h"

// We keep a list of named prop cache entries
std::list<NamedPropCacheEntry> g_lpNamedPropCache;

void UninitializeNamedPropCache()
{
	g_lpNamedPropCache.clear();
}

// Go through all the details of copying allocated data to or from a cache entry
void CopyCacheData(
	const MAPINAMEID &src,
	MAPINAMEID& dst,
	_In_opt_ LPVOID lpMAPIParent) // If passed, allocate using MAPI with this as a parent
{
	dst.lpguid = nullptr;
	dst.Kind.lID = 0;

	if (src.lpguid)
	{
		if (lpMAPIParent)
		{
			MAPIAllocateMore(sizeof(GUID), lpMAPIParent, reinterpret_cast<LPVOID*>(&dst.lpguid));
		}
		else dst.lpguid = new GUID;

		if (dst.lpguid)
		{
			memcpy(dst.lpguid, src.lpguid, sizeof GUID);
		}
	}

	dst.ulKind = src.ulKind;
	if (MNID_ID == src.ulKind)
	{
		dst.Kind.lID = src.Kind.lID;
	}
	else if (MNID_STRING == src.ulKind)
	{
		if (src.Kind.lpwstrName)
		{
			// lpSrcName is LPWSTR which means it's ALWAYS unicode
			// But some folks get it wrong and stuff ANSI data in there
			// So we check the string length both ways to make our best guess
			size_t cchShortLen = NULL;
			size_t cchWideLen = NULL;
			auto hRes = S_OK;
			WC_H(StringCchLengthA(reinterpret_cast<LPSTR>(src.Kind.lpwstrName), STRSAFE_MAX_CCH, &cchShortLen));
			WC_H(StringCchLengthW(src.Kind.lpwstrName, STRSAFE_MAX_CCH, &cchWideLen));
			size_t cbName = NULL;

			if (cchShortLen < cchWideLen)
			{
				// this is the *proper* case
				cbName = (cchWideLen + 1) * sizeof WCHAR;
			}
			else
			{
				// this is the case where ANSI data was shoved into a unicode string.
				// add a couple extra NULL in case we read this as unicode again.
				cbName = (cchShortLen + 3) * sizeof CHAR;
			}

			if (lpMAPIParent)
			{
				MAPIAllocateMore(static_cast<ULONG>(cbName), lpMAPIParent, reinterpret_cast<LPVOID*>(&dst.Kind.lpwstrName));
			}
			else dst.Kind.lpwstrName = reinterpret_cast<LPWSTR>(new BYTE[cbName]);

			if (dst.Kind.lpwstrName)
			{
				memcpy(dst.Kind.lpwstrName, src.Kind.lpwstrName, cbName);
			}
		}
	}
}

NamedPropCacheEntry::NamedPropCacheEntry(
	ULONG cbSig,
	_In_opt_count_(cbSig) LPBYTE lpSig,
	LPMAPINAMEID lpPropName,
	ULONG ulPropID)
{
	this->ulPropID = ulPropID;
	this->cbSig = cbSig;
	this->lpmniName = nullptr;
	this->lpSig = nullptr;
	this->bStringsCached = false;

	if (lpPropName)
	{
		lpmniName = new MAPINAMEID;

		if (lpmniName)
		{
			CopyCacheData(*lpPropName, *lpmniName, nullptr);
		}
	}

	if (cbSig && lpSig)
	{
		this->lpSig = new BYTE[cbSig];
		if (this->lpSig)
		{
			memcpy(this->lpSig, lpSig, cbSig);
		}
	}
}

NamedPropCacheEntry::~NamedPropCacheEntry()
{
	if (lpmniName)
	{
		delete lpmniName->lpguid;
		if (MNID_STRING == lpmniName->ulKind)
		{
			delete[] lpmniName->Kind.lpwstrName;
		}

		delete lpmniName;
	}

	delete[] lpSig;
}

// Given a signature and property ID (ulPropID), finds the named prop mapping in the cache
_Check_return_ LPNAMEDPROPCACHEENTRY FindCacheEntry(ULONG cbSig, _In_count_(cbSig) LPBYTE lpSig, ULONG ulPropID)
{
	auto entry = find_if(begin(g_lpNamedPropCache), end(g_lpNamedPropCache), [&](NamedPropCacheEntry &namedPropCacheEntry)
	{
		if (namedPropCacheEntry.ulPropID != ulPropID) return false;
		if (namedPropCacheEntry.cbSig != cbSig) return false;
		if (cbSig && memcmp(lpSig, namedPropCacheEntry.lpSig, cbSig)) return false;

		return true;
	});

	return entry != end(g_lpNamedPropCache) ? &(*entry) : nullptr;
}

// Given a signature, guid, kind, and value, finds the named prop mapping in the cache
_Check_return_ LPNAMEDPROPCACHEENTRY FindCacheEntry(ULONG cbSig, _In_count_(cbSig) LPBYTE lpSig, _In_ LPGUID lpguid, ULONG ulKind, LONG lID, _In_z_ LPWSTR lpwstrName)
{
	auto entry = find_if(begin(g_lpNamedPropCache), end(g_lpNamedPropCache), [&](NamedPropCacheEntry &namedPropCacheEntry)
	{
		if (!namedPropCacheEntry.lpmniName) return false;
		if (namedPropCacheEntry.lpmniName->ulKind != ulKind) return false;
		if (MNID_ID == ulKind && namedPropCacheEntry.lpmniName->Kind.lID != lID) return false;
		if (MNID_STRING == ulKind && 0 != lstrcmpW(namedPropCacheEntry.lpmniName->Kind.lpwstrName, lpwstrName)) return false;;
		if (0 != memcmp(namedPropCacheEntry.lpmniName->lpguid, lpguid, sizeof(GUID))) return false;
		if (cbSig != namedPropCacheEntry.cbSig) return false;
		if (cbSig && memcmp(lpSig, namedPropCacheEntry.lpSig, cbSig)) return false;

		return true;
	});

	return entry != end(g_lpNamedPropCache) ? &(*entry) : nullptr;
}

// Given a tag, guid, kind, and value, finds the named prop mapping in the cache
_Check_return_ LPNAMEDPROPCACHEENTRY FindCacheEntry(ULONG ulPropID, _In_ LPGUID lpguid, ULONG ulKind, LONG lID, _In_z_ LPWSTR lpwstrName)
{
	auto entry = find_if(begin(g_lpNamedPropCache), end(g_lpNamedPropCache), [&](NamedPropCacheEntry &namedPropCacheEntry)
	{
		if (namedPropCacheEntry.ulPropID != ulPropID) return false;
		if (!namedPropCacheEntry.lpmniName) return false;
		if (namedPropCacheEntry.lpmniName->ulKind != ulKind) return false;
		if (MNID_ID == ulKind && namedPropCacheEntry.lpmniName->Kind.lID != lID) return false;
		if (MNID_STRING == ulKind && 0 != lstrcmpW(namedPropCacheEntry.lpmniName->Kind.lpwstrName, lpwstrName)) return false;
		if (0 != memcmp(namedPropCacheEntry.lpmniName->lpguid, lpguid, sizeof(GUID))) return false;

		return true;
	});

	return entry != end(g_lpNamedPropCache) ? &(*entry) : nullptr;
}

void AddMapping(ULONG cbSig, // Count bytes of signature
	_In_opt_count_(cbSig) LPBYTE lpSig, // Signature
	ULONG ulNumProps, // Number of mapped names
	_In_count_(ulNumProps) LPMAPINAMEID* lppPropNames, // Output from GetNamesFromIDs, input for GetIDsFromNames
	_In_ LPSPropTagArray lpTag) // Input for GetNamesFromIDs, output from GetIDsFromNames
{
	if (!ulNumProps || !lppPropNames || !lpTag) return;
	if (ulNumProps != lpTag->cValues) return; // Wouldn't know what to do with this

	for (ULONG ulSource = 0; ulSource < ulNumProps; ulSource++)
	{
		if (lppPropNames[ulSource])
		{
			g_lpNamedPropCache.emplace_back(cbSig, lpSig, lppPropNames[ulSource], PROP_ID(lpTag->aulPropTag[ulSource]));
		}
	}
}

// Add to the cache entries that don't have a mapping signature
// For each one, we have to check that the item isn't already in the cache
// Since this function should rarely be hit, we'll do it the slow but easy way...
// One entry at a time
void AddMappingWithoutSignature(ULONG ulNumProps, // Number of mapped names
	_In_count_(ulNumProps) LPMAPINAMEID* lppPropNames, // Output from GetNamesFromIDs, input for GetIDsFromNames
	_In_ LPSPropTagArray lpTag) // Input for GetNamesFromIDs, output from GetIDsFromNames
{
	if (!ulNumProps || !lppPropNames || !lpTag) return;
	if (ulNumProps != lpTag->cValues) return; // Wouldn't know what to do with this

	SPropTagArray sptaTag = { 0 };
	sptaTag.cValues = 1;
	for (ULONG ulProp = 0; ulProp < ulNumProps; ulProp++)
	{
		if (lppPropNames[ulProp])
		{
			auto lpNamedPropCacheEntry = FindCacheEntry(
				PROP_ID(lpTag->aulPropTag[ulProp]),
				lppPropNames[ulProp]->lpguid,
				lppPropNames[ulProp]->ulKind,
				lppPropNames[ulProp]->Kind.lID,
				lppPropNames[ulProp]->Kind.lpwstrName);
			if (!lpNamedPropCacheEntry)
			{
				sptaTag.aulPropTag[0] = lpTag->aulPropTag[ulProp];
				AddMapping(0, nullptr, 1, &lppPropNames[ulProp], &sptaTag);
			}
		}
	}
}

_Check_return_ HRESULT CacheGetNamesFromIDs(_In_ LPMAPIPROP lpMAPIProp,
	ULONG cbSig,
	_In_count_(cbSig) LPBYTE lpSig,
	_In_ LPSPropTagArray* lppPropTags,
	_Out_ ULONG* lpcPropNames,
	_Out_ _Deref_post_cap_(*lpcPropNames) LPMAPINAMEID** lpppPropNames)
{
	if (!lpMAPIProp || !lppPropTags || !*lppPropTags || !cbSig || !lpSig) return MAPI_E_INVALID_PARAMETER;

	auto hRes = S_OK;

	// We're going to walk the cache, looking for the values we need. As soon as we have all the values we need, we're done
	// If we reach the end of the cache and don't have everything, we set up to make a GetNamesFromIDs call.

	auto lpPropTags = *lppPropTags;
	// First, allocate our results using MAPI
	LPMAPINAMEID* lppNameIDs = nullptr;
	EC_H(MAPIAllocateBuffer(sizeof(MAPINAMEID*)* lpPropTags->cValues, reinterpret_cast<LPVOID*>(&lppNameIDs)));

	if (lppNameIDs)
	{
		memset(lppNameIDs, 0, sizeof(MAPINAMEID*)* lpPropTags->cValues);

		// Assume we'll miss on everything
		auto ulMisses = lpPropTags->cValues;

		// For each tag we wish to look up...
		for (ULONG ulTarget = 0; ulTarget < lpPropTags->cValues; ulTarget++)
		{
			// ...check the cache
			auto lpEntry = FindCacheEntry(cbSig, lpSig, PROP_ID(lpPropTags->aulPropTag[ulTarget]));

			if (lpEntry)
			{
				// We have a hit - copy the data over
				lppNameIDs[ulTarget] = lpEntry->lpmniName;

				// Got a hit, decrement the miss counter
				ulMisses--;
			}
		}

		// Go to MAPI with whatever's left. We set up for a single call to GetNamesFromIDs.
		if (0 != ulMisses)
		{
			LPSPropTagArray lpUncachedTags = nullptr;
			EC_H(MAPIAllocateBuffer(CbNewSPropTagArray(ulMisses), reinterpret_cast<LPVOID*>(&lpUncachedTags)));
			if (lpUncachedTags)
			{
				memset(lpUncachedTags, 0, CbNewSPropTagArray(ulMisses));
				lpUncachedTags->cValues = ulMisses;
				ULONG ulUncachedTag = NULL;
				for (ULONG ulTarget = 0; ulTarget < lpPropTags->cValues; ulTarget++)
				{
					// We're looking for any result which doesn't have a mapping
					if (!lppNameIDs[ulTarget])
					{
						lpUncachedTags->aulPropTag[ulUncachedTag] = lpPropTags->aulPropTag[ulTarget];
						ulUncachedTag++;
					}
				}

				ULONG ulUncachedPropNames = 0;
				LPMAPINAMEID* lppUncachedPropNames = nullptr;

				WC_H_GETPROPS(lpMAPIProp->GetNamesFromIDs(
					&lpUncachedTags,
					nullptr,
					NULL,
					&ulUncachedPropNames,
					&lppUncachedPropNames));
				if (SUCCEEDED(hRes) && ulUncachedPropNames == ulMisses && lppUncachedPropNames)
				{
					// Cache the results
					AddMapping(cbSig, lpSig, ulUncachedPropNames, lppUncachedPropNames, lpUncachedTags);

					// Copy our results over
					// Loop over the target array, looking for empty slots
					ulUncachedTag = 0;
					for (ULONG ulTarget = 0; ulTarget < lpPropTags->cValues; ulTarget++)
					{
						// Found an empty slot
						if (!lppNameIDs[ulTarget])
						{
							// copy the next result into it
							if (lppUncachedPropNames[ulUncachedTag])
							{
								LPMAPINAMEID lpNameID = nullptr;

								EC_H(MAPIAllocateMore(sizeof(MAPINAMEID), lppNameIDs, reinterpret_cast<LPVOID*>(&lpNameID)));
								if (lpNameID)
								{
									CopyCacheData(*lppUncachedPropNames[ulUncachedTag], *lpNameID, lppNameIDs);
									lppNameIDs[ulTarget] = lpNameID;

									// Got a hit, decrement the miss counter
									ulMisses--;
								}
							}
							// Whether we copied or not, move on to the next one
							ulUncachedTag++;
						}
					}
				}

				MAPIFreeBuffer(lppUncachedPropNames);
			}

			MAPIFreeBuffer(lpUncachedTags);

			if (ulMisses != 0) hRes = MAPI_W_ERRORS_RETURNED;
		}

		*lpppPropNames = lppNameIDs;
		if (lpcPropNames) *lpcPropNames = lpPropTags->cValues;
	}

	return hRes;
}

_Check_return_ HRESULT GetNamesFromIDs(_In_ LPMAPIPROP lpMAPIProp,
	_In_ LPSPropTagArray* lppPropTags,
	_In_opt_ LPGUID lpPropSetGuid,
	ULONG ulFlags,
	_Out_ ULONG* lpcPropNames,
	_Out_ _Deref_post_cap_(*lpcPropNames) LPMAPINAMEID** lpppPropNames)
{
	return GetNamesFromIDs(
		lpMAPIProp,
		nullptr,
		lppPropTags,
		lpPropSetGuid,
		ulFlags,
		lpcPropNames,
		lpppPropNames);
}

_Check_return_ HRESULT GetNamesFromIDs(_In_ LPMAPIPROP lpMAPIProp,
	_In_opt_ LPSBinary lpMappingSignature,
	_In_ LPSPropTagArray* lppPropTags,
	_In_opt_ LPGUID lpPropSetGuid,
	ULONG ulFlags,
	_Out_ ULONG* lpcPropNames,
	_Out_ _Deref_post_cap_(*lpcPropNames) LPMAPINAMEID** lpppPropNames)
{
	if (!lpMAPIProp) return MAPI_E_INVALID_PARAMETER;

	// Check if we're bypassing the cache:
	if (!fCacheNamedProps() ||
		// Assume an array was passed - none of my calling code passes a NULL tag array
		!lppPropTags || !*lppPropTags ||
		// None of my code uses these flags, but bypass the cache if we see them
		ulFlags || lpPropSetGuid)
	{
		return lpMAPIProp->GetNamesFromIDs(lppPropTags, lpPropSetGuid, ulFlags, lpcPropNames, lpppPropNames);
	}

	auto hRes = S_OK;
	LPSPropValue lpProp = nullptr;

	if (!lpMappingSignature)
	{
		// This error is too chatty to log - ignore it.
		hRes = HrGetOneProp(lpMAPIProp, PR_MAPPING_SIGNATURE, &lpProp);

		if (SUCCEEDED(hRes) && lpProp && PT_BINARY == PROP_TYPE(lpProp->ulPropTag))
		{
			lpMappingSignature = &lpProp->Value.bin;
		}
	}

	if (lpMappingSignature)
	{
		WC_H_GETPROPS(CacheGetNamesFromIDs(lpMAPIProp,
			lpMappingSignature->cb,
			lpMappingSignature->lpb,
			lppPropTags,
			lpcPropNames,
			lpppPropNames));
	}
	else
	{
		hRes = S_OK;
		WC_H_GETPROPS(lpMAPIProp->GetNamesFromIDs(lppPropTags, lpPropSetGuid, ulFlags, lpcPropNames, lpppPropNames));
		// Cache the results
		if (SUCCEEDED(hRes))
		{
			AddMappingWithoutSignature(*lpcPropNames, *lpppPropNames, *lppPropTags);
		}
	}

	MAPIFreeBuffer(lpProp);
	return hRes;
}

_Check_return_ HRESULT CacheGetIDsFromNames(_In_ LPMAPIPROP lpMAPIProp,
	ULONG cbSig,
	_In_count_(cbSig) LPBYTE lpSig,
	ULONG cPropNames,
	_In_count_(cPropNames) LPMAPINAMEID* lppPropNames,
	ULONG ulFlags,
	_Out_opt_cap_(cPropNames) LPSPropTagArray* lppPropTags)
{
	if (!lpMAPIProp || !cPropNames || !*lppPropNames || !lppPropTags) return MAPI_E_INVALID_PARAMETER;

	auto hRes = S_OK;

	// We're going to walk the cache, looking for the values we need. As soon as we have all the values we need, we're done
	// If we reach the end of the cache and don't have everything, we set up to make a GetIDsFromNames call.

	// First, allocate our results using MAPI
	LPSPropTagArray lpPropTags = nullptr;
	EC_H(MAPIAllocateBuffer(CbNewSPropTagArray(cPropNames), reinterpret_cast<LPVOID*>(&lpPropTags)));

	if (lpPropTags)
	{
		memset(lpPropTags, 0, CbNewSPropTagArray(cPropNames));
		lpPropTags->cValues = cPropNames;

		// Assume we'll miss on everything
		auto ulMisses = cPropNames;

		// For each tag we wish to look up...
		for (ULONG ulTarget = 0; ulTarget < cPropNames; ulTarget++)
		{
			// ...check the cache
			auto lpEntry = FindCacheEntry(
				cbSig,
				lpSig,
				lppPropNames[ulTarget]->lpguid,
				lppPropNames[ulTarget]->ulKind,
				lppPropNames[ulTarget]->Kind.lID,
				lppPropNames[ulTarget]->Kind.lpwstrName);

			if (lpEntry)
			{
				// We have a hit - copy the data over
				lpPropTags->aulPropTag[ulTarget] = PROP_TAG(PT_UNSPECIFIED, lpEntry->ulPropID);

				// Got a hit, decrement the miss counter
				ulMisses--;
			}
		}

		// Go to MAPI with whatever's left. We set up for a single call to GetIDsFromNames.
		if (0 != ulMisses)
		{
			LPMAPINAMEID* lppUncachedPropNames = nullptr;
			EC_H(MAPIAllocateBuffer(sizeof(LPMAPINAMEID)* ulMisses, reinterpret_cast<LPVOID*>(&lppUncachedPropNames)));
			if (lppUncachedPropNames)
			{
				memset(lppUncachedPropNames, 0, sizeof(LPMAPINAMEID)* ulMisses);
				ULONG ulUncachedName = NULL;
				for (ULONG ulTarget = 0; ulTarget < cPropNames; ulTarget++)
				{
					// We're looking for any result which doesn't have a mapping
					if (!lpPropTags->aulPropTag[ulTarget])
					{
						// We don't need to reallocate everything - just match the pointers so we can make the call
						lppUncachedPropNames[ulUncachedName] = lppPropNames[ulTarget];
						ulUncachedName++;
					}
				}

				ULONG ulUncachedTags = 0;
				LPSPropTagArray lpUncachedTags = nullptr;

				EC_H_GETPROPS(lpMAPIProp->GetIDsFromNames(
					ulMisses,
					lppUncachedPropNames,
					ulFlags,
					&lpUncachedTags));
				if (SUCCEEDED(hRes) && lpUncachedTags && lpUncachedTags->cValues == ulMisses)
				{
					// Cache the results
					AddMapping(cbSig, lpSig, ulUncachedTags, lppUncachedPropNames, lpUncachedTags);

					// Copy our results over
					// Loop over the target array, looking for empty slots
					// Each empty slot corresponds to one of our results, in order
					ulUncachedName = 0;
					for (ULONG ulTarget = 0; ulTarget < lpPropTags->cValues; ulTarget++)
					{
						// Found an empty slot
						if (!lpPropTags->aulPropTag[ulTarget])
						{
							lpPropTags->aulPropTag[ulTarget] = lpUncachedTags->aulPropTag[ulUncachedName];

							// If we got a hit, decrement the miss counter
							// But only if the hit wasn't an error
							if (lpPropTags->aulPropTag[ulTarget] &&
								PT_ERROR != lpPropTags->aulPropTag[ulTarget])
							{
								ulMisses--;
							}

							// Move on to the next uncached name
							ulUncachedName++;
						}
					}
				}
				MAPIFreeBuffer(lpUncachedTags);
			}
			MAPIFreeBuffer(lppUncachedPropNames);

			if (ulMisses != 0) hRes = MAPI_W_ERRORS_RETURNED;
		}

		*lppPropTags = lpPropTags;
	}

	return hRes;
}

_Check_return_ HRESULT GetIDsFromNames(_In_ LPMAPIPROP lpMAPIProp,
	ULONG cPropNames,
	_In_opt_count_(cPropNames) LPMAPINAMEID* lppPropNames,
	ULONG ulFlags,
	_Out_ _Deref_post_cap_(cPropNames) LPSPropTagArray* lppPropTags)
{
	if (!lpMAPIProp) return MAPI_E_INVALID_PARAMETER;

	// Check if we're bypassing the cache:
	if (!fCacheNamedProps() ||
		// If no names were passed, we have to bypass the cache
		// Should we cache results?
		!cPropNames || !lppPropNames || !*lppPropNames)
	{
		return lpMAPIProp->GetIDsFromNames(cPropNames, lppPropNames, ulFlags, lppPropTags);
	}

	auto hRes = S_OK;

	LPSPropValue lpProp = nullptr;

	WC_MAPI(HrGetOneProp(lpMAPIProp, PR_MAPPING_SIGNATURE, &lpProp));

	if (SUCCEEDED(hRes) && lpProp && PT_BINARY == PROP_TYPE(lpProp->ulPropTag))
	{
		WC_H_GETPROPS(CacheGetIDsFromNames(lpMAPIProp,
			lpProp->Value.bin.cb,
			lpProp->Value.bin.lpb,
			cPropNames,
			lppPropNames,
			ulFlags,
			lppPropTags));
	}
	else
	{
		hRes = S_OK;
		WC_H_GETPROPS(lpMAPIProp->GetIDsFromNames(cPropNames, lppPropNames, ulFlags, lppPropTags));
		// Cache the results
		if (SUCCEEDED(hRes))
		{
			AddMappingWithoutSignature(cPropNames, lppPropNames, *lppPropTags);
		}
	}

	MAPIFreeBuffer(lpProp);

	return hRes;
}