// NamedPropCache.cpp: implementation file
//

#include "stdafx.h"
#include "NamedPropCache.h"

// We keep an array of named prop caches, one for each PR_MAPPING_SIGNATURE
ULONG g_ulNamedPropCacheSize = 0;
ULONG g_ulNamedPropCacheNumEntries = 0;
LPNAMEDPROPCACHEENTRY g_lpNamedPropCache = NULL;
BOOL g_bNamedPropCacheInitialized = false;

// Initial size of the named prop cache
#define NAMED_PROP_CACHE_ENTRY_COUNT 32

void UninitializeNamedPropCache()
{
	ULONG ulMap = 0;

	for (ulMap = 0 ; ulMap < g_ulNamedPropCacheNumEntries ; ulMap++)
	{
		if (g_lpNamedPropCache[ulMap].lpmniName)
		{
			delete g_lpNamedPropCache[ulMap].lpmniName->lpguid;
			if (MNID_STRING == g_lpNamedPropCache[ulMap].lpmniName->ulKind)
			{
				delete[] g_lpNamedPropCache[ulMap].lpmniName->Kind.lpwstrName;
			}
			delete g_lpNamedPropCache[ulMap].lpmniName;
		}
		delete[] g_lpNamedPropCache[ulMap].lpSig;
		delete[] g_lpNamedPropCache[ulMap].lpszPropName;
		delete[] g_lpNamedPropCache[ulMap].lpszPropGUID;
		delete[] g_lpNamedPropCache[ulMap].lpszDASL;
	}
	delete[] g_lpNamedPropCache;
} // UninitializeNamedPropCache

BOOL InitializeNamedPropCache()
{
	// Already initialized
	if (g_bNamedPropCacheInitialized) return true;

	g_lpNamedPropCache = new NamedPropCacheEntry[NAMED_PROP_CACHE_ENTRY_COUNT];
	if (g_lpNamedPropCache)
	{
		memset(g_lpNamedPropCache, 0, sizeof(NamedPropCacheEntry)*NAMED_PROP_CACHE_ENTRY_COUNT);
		g_ulNamedPropCacheSize = NAMED_PROP_CACHE_ENTRY_COUNT;
		g_ulNamedPropCacheNumEntries = 0;
		g_bNamedPropCacheInitialized = true;
	}

	return g_bNamedPropCacheInitialized;
} // InitializeNamedPropCache

// Given a signature and property ID (ulPropID), finds the named prop mapping in the cache
LPNAMEDPROPCACHEENTRY FindCacheEntry(ULONG cbSig, LPBYTE lpSig, ULONG ulPropID)
{
	// No results without a cache
	if (!g_bNamedPropCacheInitialized) return NULL;

	ULONG i = 0;

	for (i = 0 ; i < g_ulNamedPropCacheNumEntries ; i++)
	{
		if (g_lpNamedPropCache[i].ulPropID != ulPropID) continue;
		if (cbSig != g_lpNamedPropCache[i].cbSig) continue;
		if (cbSig && memcmp(lpSig, g_lpNamedPropCache[i].lpSig, cbSig)) continue;

		return &g_lpNamedPropCache[i];
	}
	return NULL;
} // FindCacheEntry (cbSig, lpSig, ulPropID)

// Given a signature, guid, kind, and value, finds the named prop mapping in the cache
LPNAMEDPROPCACHEENTRY FindCacheEntry(ULONG cbSig, LPBYTE lpSig, LPGUID lpguid, ULONG ulKind, LONG lID, LPWSTR lpwstrName)
{
	// No results without a cache
	if (!g_bNamedPropCacheInitialized) return NULL;

	ULONG i = 0;

	for (i = 0 ; i < g_ulNamedPropCacheNumEntries ; i++)
	{
		if (!g_lpNamedPropCache[i].lpmniName) continue;
		if (g_lpNamedPropCache[i].lpmniName->ulKind != ulKind) continue;
		if (MNID_ID == ulKind && g_lpNamedPropCache[i].lpmniName->Kind.lID != lID) continue;
		if (MNID_STRING == ulKind && 0 != lstrcmpW(g_lpNamedPropCache[i].lpmniName->Kind.lpwstrName,lpwstrName)) continue;
		if (0 != memcmp(g_lpNamedPropCache[i].lpmniName->lpguid, lpguid, sizeof(GUID))) continue;
		if (cbSig != g_lpNamedPropCache[i].cbSig) continue;
		if (cbSig && memcmp(lpSig, g_lpNamedPropCache[i].lpSig, cbSig)) continue;

		return &g_lpNamedPropCache[i];
	}
	return NULL;
} // FindCacheEntry (lpguid, ulKind, wID, lpwstrName)

// Given a tag, guid, kind, and value, finds the named prop mapping in the cache
LPNAMEDPROPCACHEENTRY FindCacheEntry(ULONG ulPropID, LPGUID lpguid, ULONG ulKind, LONG lID, LPWSTR lpwstrName)
{
	// No results without a cache
	if (!g_bNamedPropCacheInitialized) return NULL;

	ULONG i = 0;

	for (i = 0 ; i < g_ulNamedPropCacheNumEntries ; i++)
	{
		if (g_lpNamedPropCache[i].ulPropID != ulPropID) continue;
		if (!g_lpNamedPropCache[i].lpmniName) continue;
		if (g_lpNamedPropCache[i].lpmniName->ulKind != ulKind) continue;
		if (MNID_ID == ulKind && g_lpNamedPropCache[i].lpmniName->Kind.lID != lID) continue;
		if (MNID_STRING == ulKind && 0 != lstrcmpW(g_lpNamedPropCache[i].lpmniName->Kind.lpwstrName,lpwstrName)) continue;
		if (0 != memcmp(g_lpNamedPropCache[i].lpmniName->lpguid, lpguid, sizeof(GUID))) continue;

		return &g_lpNamedPropCache[i];
	}
	return NULL;
} // FindCacheEntry (ulPropID, lpguid, ulKind, wID, lpwstrName)

// Go through all the details of copying allocated data to or from a cache entry
void CopyCacheData(
				   LPGUID lpSrcGUID,
				   ULONG ulSrcKind,
				   LONG lSrcID,
				   LPWSTR lpSrcName,
				   LPGUID* lppDstGUID,
				   ULONG* lpulDstKind,
				   LONG* lplDstID,
				   LPWSTR* lppDstName,
				   LPVOID lpMAPIParent) // If passed, allocate using MAPI with this as a parent
{
	if (lpSrcGUID && lppDstGUID)
	{
		LPGUID lpDstGUID = NULL;
		if (lpMAPIParent)
		{
			MAPIAllocateMore(sizeof(GUID), lpMAPIParent, (LPVOID*) &lpDstGUID);
		}
		else lpDstGUID = new GUID;

		if (lpDstGUID)
		{
			memcpy(lpDstGUID, lpSrcGUID, sizeof GUID);
		}
		*lppDstGUID = lpDstGUID;
	}

	if (lpulDstKind) *lpulDstKind = ulSrcKind;
	if (MNID_ID == ulSrcKind)
	{
		if (lplDstID) *lplDstID = lSrcID;
	}
	else if (MNID_STRING == ulSrcKind)
	{
		if (lpSrcName && lppDstName)
		{
			// lpSrcName is LPWSTR which means it's ALWAYS unicode
			// But some folks get it wrong and stuff ANSI data in there
			// So we check the string length both ways to make our best guess
			size_t cchShortLen = NULL;
			size_t cchWideLen = NULL;
			LPWSTR lpDstName = NULL;
			HRESULT hRes = S_OK;
			WC_H(StringCchLengthA((LPSTR)lpSrcName,STRSAFE_MAX_CCH,&cchShortLen));
			WC_H(StringCchLengthW(lpSrcName,STRSAFE_MAX_CCH,&cchWideLen));
			ULONG cbName = NULL;

			if (cchShortLen < cchWideLen)
			{
				// this is the *proper* case
				cbName = (ULONG) (cchWideLen+1) * sizeof(WCHAR);
			}
			else
			{
				// this is the case where ANSI data was shoved into a unicode string.
				// add a couple extra NULL in case we read this as unicode again.
				cbName = (ULONG) (cchShortLen+3) * sizeof(CHAR);
			}
			if (lpMAPIParent)
			{
				MAPIAllocateMore(cbName, lpMAPIParent, (LPVOID*) &lpDstName);
			}
			else lpDstName = (LPWSTR) new BYTE[cbName];

			if (lpDstName)
			{
				memcpy(lpDstName,
					lpSrcName,
					cbName);
			}
			*lppDstName = lpDstName;
		}
	}
} // CopyCacheData

void AddMapping(ULONG cbSig,                // Count bytes of signature
				LPBYTE lpSig,               // Signature
				ULONG ulNumProps,           // Number of mapped names
				LPMAPINAMEID* lppPropNames, // Output from GetNamesFromIDs, input for GetIDsFromNames
				LPSPropTagArray lpTag)      // Input for GetNamesFromIDs, output from GetIDsFromNames
{
	if (!ulNumProps || !lppPropNames || !lpTag) return;
	if (ulNumProps != lpTag->cValues) return; // Wouldn't know what to do with this

	// We'll assume that any mapping being added doesn't already exist
	// Revisit this assumption if it turns out to be a problem

	ULONG ulNewNumEntries = g_ulNamedPropCacheNumEntries + ulNumProps;

	// Check if we need to reallocate
	if (ulNewNumEntries > g_ulNamedPropCacheSize)
	{
		// Naive algorithm - allocate what we need plus a bit more
		ULONG ulNewCacheSize = ulNewNumEntries + NAMED_PROP_CACHE_ENTRY_COUNT;
		LPNAMEDPROPCACHEENTRY lpNewNamedPropCache = new NamedPropCacheEntry[ulNewCacheSize];

		// Can't allocate, nothing to do - fail
		if (!lpNewNamedPropCache) return;

		memset(lpNewNamedPropCache, 0, sizeof(NamedPropCacheEntry) * ulNewCacheSize);
		memcpy(lpNewNamedPropCache, g_lpNamedPropCache, sizeof(NamedPropCacheEntry)*g_ulNamedPropCacheNumEntries);
		delete[] g_lpNamedPropCache; // no need to delete more, everything has been passed to the new array
		g_lpNamedPropCache = lpNewNamedPropCache;
		g_ulNamedPropCacheSize = ulNewCacheSize;
	}

	// Now we know we have enough room - loop the source and add the entries
	ULONG ulCache = 0;
	ULONG ulSource = 0;
	for (ulCache = g_ulNamedPropCacheNumEntries, ulSource = 0 ; ulSource < ulNumProps ; ulCache++, ulSource++)
	{
		if (lppPropNames[ulSource])
		{
			g_lpNamedPropCache[ulCache].lpmniName = new MAPINAMEID;

			if (g_lpNamedPropCache[ulCache].lpmniName)
			{
				g_lpNamedPropCache[ulCache].ulPropID = PROP_ID(lpTag->aulPropTag[ulSource]);
				CopyCacheData(
					lppPropNames[ulSource]->lpguid,
					lppPropNames[ulSource]->ulKind,
					lppPropNames[ulSource]->Kind.lID,
					lppPropNames[ulSource]->Kind.lpwstrName,
					&g_lpNamedPropCache[ulCache].lpmniName->lpguid,
					&g_lpNamedPropCache[ulCache].lpmniName->ulKind,
					&g_lpNamedPropCache[ulCache].lpmniName->Kind.lID,
					&g_lpNamedPropCache[ulCache].lpmniName->Kind.lpwstrName,
					NULL);
				if (cbSig)
				{
					g_lpNamedPropCache[ulCache].lpSig = new BYTE[cbSig];
					if (g_lpNamedPropCache[ulCache].lpSig)
					{
						memcpy(g_lpNamedPropCache[ulCache].lpSig, lpSig, cbSig);
						g_lpNamedPropCache[ulCache].cbSig = cbSig;
					}
				}
			}
		}
	}

	g_ulNamedPropCacheNumEntries = ulNewNumEntries;
} // AddMapping

// Add to the cache entries that don't have a mapping signature
// For each one, we have to check that the item isn't already in the cache
// Since this function should rarely be hit, we'll do it the slow but easy way...
// One entry at a time
void AddMappingWithoutSignature(ULONG ulNumProps,      // Number of mapped names
						   LPMAPINAMEID* lppPropNames, // Output from GetNamesFromIDs, input for GetIDsFromNames
						   LPSPropTagArray lpTag)     // Input for GetNamesFromIDs, output from GetIDsFromNames
{
	if (!ulNumProps || !lppPropNames || !lpTag) return;
	if (ulNumProps != lpTag->cValues) return; // Wouldn't know what to do with this

	ULONG ulProp = 0;
	SPropTagArray sptaTag = {0};
	sptaTag.cValues = 1;
	for (ulProp = 0 ; ulProp < ulNumProps ; ulProp++)
	{
		if (lppPropNames[ulProp])
		{
			LPNAMEDPROPCACHEENTRY lpNamedPropCacheEntry = FindCacheEntry(
				PROP_ID(lpTag->aulPropTag[ulProp]),
				lppPropNames[ulProp]->lpguid,
				lppPropNames[ulProp]->ulKind,
				lppPropNames[ulProp]->Kind.lID,
				lppPropNames[ulProp]->Kind.lpwstrName);
			if (!lpNamedPropCacheEntry)
			{
				sptaTag.aulPropTag[0] = lpTag->aulPropTag[ulProp];
				AddMapping(0,0,1,&lppPropNames[ulProp],&sptaTag);
			}
		}
	}
} // AddMappingWithoutSignature

HRESULT CacheGetNamesFromIDs(LPMAPIPROP lpMAPIProp,
							 ULONG cbSig,
							 LPBYTE lpSig,
							 LPSPropTagArray* lppPropTags,
							 ULONG* lpcPropNames,
							 LPMAPINAMEID** lpppPropNames)
{
	if (!lpMAPIProp || !lppPropTags || !*lppPropTags || !cbSig || !lpSig) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;

	// We're going to walk the cache, looking for the values we need. As soon as we have all the values we need, we're done
	// If we reach the end of the cache and don't have everything, we set up to make a GetNamesFromIDs call.

	// First, allocate our results using MAPI
	LPMAPINAMEID* lppNameIDs = NULL;
	EC_H(MAPIAllocateBuffer(sizeof(MAPINAMEID*) * (*lppPropTags)->cValues, (LPVOID*) &lppNameIDs));

	if (lppNameIDs)
	{
		memset(lppNameIDs, 0, sizeof(MAPINAMEID*) * (*lppPropTags)->cValues);
		ULONG ulTarget = 0;

		// Assume we'll miss on everything
		ULONG ulMisses = (*lppPropTags)->cValues;

		// For each tag we wish to look up...
		for (ulTarget = 0 ; ulTarget < (*lppPropTags)->cValues ; ulTarget++)
		{
			// ...check the cache
			LPNAMEDPROPCACHEENTRY lpEntry = FindCacheEntry(cbSig, lpSig, PROP_ID((*lppPropTags)->aulPropTag[ulTarget]));

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
			LPSPropTagArray lpUncachedTags = NULL;
			EC_H(MAPIAllocateBuffer(CbNewSPropTagArray(ulMisses),(LPVOID*) &lpUncachedTags));
			if (lpUncachedTags)
			{
				memset(lpUncachedTags, 0, CbNewSPropTagArray(ulMisses));
				lpUncachedTags->cValues = ulMisses;
				ULONG ulUncachedTag = NULL;
				for (ulTarget = 0 ; ulTarget < (*lppPropTags)->cValues ; ulTarget++)
				{
					// We're looking for any result which doesn't have a mapping
					if (!lppNameIDs[ulTarget])
					{
						lpUncachedTags->aulPropTag[ulUncachedTag] = (*lppPropTags)->aulPropTag[ulTarget];
						ulUncachedTag++;
					}
				}

				ULONG ulUncachedPropNames = 0;
				LPMAPINAMEID* lppUncachedPropNames = 0;

				WC_H_GETPROPS(lpMAPIProp->GetNamesFromIDs(
					&lpUncachedTags,
					NULL,
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
					for (ulTarget = 0 ; ulTarget < (*lppPropTags)->cValues ; ulTarget++)
					{
						// Found an empty slot
						if (!lppNameIDs[ulTarget])
						{
							// copy the next result into it
							if (lppUncachedPropNames[ulUncachedTag])
							{
								LPMAPINAMEID lpNameID = NULL;

								EC_H(MAPIAllocateMore(sizeof(MAPINAMEID), lppNameIDs, (LPVOID*) &lpNameID));
								if (lpNameID)
								{
									CopyCacheData(
										lppUncachedPropNames[ulUncachedTag]->lpguid,
										lppUncachedPropNames[ulUncachedTag]->ulKind,
										lppUncachedPropNames[ulUncachedTag]->Kind.lID,
										lppUncachedPropNames[ulUncachedTag]->Kind.lpwstrName,
										&lpNameID->lpguid,
										&lpNameID->ulKind,
										&lpNameID->Kind.lID,
										&lpNameID->Kind.lpwstrName,
										lppNameIDs);
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
		if (lpcPropNames) *lpcPropNames = (*lppPropTags)->cValues;
	}

	return hRes;
} // CacheGetNamesFromIDs

HRESULT GetNamesFromIDs(LPMAPIPROP lpMAPIProp,
						  LPSPropTagArray* lppPropTags,
						  LPGUID lpPropSetGuid,
						  ULONG ulFlags,
						  ULONG* lpcPropNames,
						  LPMAPINAMEID** lpppPropNames)
{
	return GetNamesFromIDs(
		lpMAPIProp,
		NULL,
		lppPropTags,
		lpPropSetGuid,
		ulFlags,
		lpcPropNames,
		lpppPropNames);
} // GetNamesFromIDs

HRESULT GetNamesFromIDs(LPMAPIPROP lpMAPIProp,
						LPSBinary lpMappingSignature,
						LPSPropTagArray* lppPropTags,
						LPGUID lpPropSetGuid,
						ULONG ulFlags,
						ULONG* lpcPropNames,
						LPMAPINAMEID** lpppPropNames)
{
	if (!lpMAPIProp) return MAPI_E_INVALID_PARAMETER;

	// Check if we're bypassing the cache:
	if (!fCacheNamedProps() ||
	// Assume an array was passed - none of my calling code passes a NULL tag array
		!lppPropTags || !*lppPropTags ||
	// None of my code uses these flags, but bypass the cache if we see them
		ulFlags || lpPropSetGuid)
	{
		return lpMAPIProp->GetNamesFromIDs(lppPropTags,lpPropSetGuid,ulFlags,lpcPropNames,lpppPropNames);
	}
	if (!InitializeNamedPropCache()) return MAPI_E_NOT_ENOUGH_MEMORY;

	HRESULT hRes = S_OK;
	LPSPropValue lpProp = NULL;

	if (!lpMappingSignature)
	{
		WC_H(HrGetOneProp(lpMAPIProp,PR_MAPPING_SIGNATURE,&lpProp));

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
		WC_H_GETPROPS(lpMAPIProp->GetNamesFromIDs(lppPropTags,lpPropSetGuid,ulFlags,lpcPropNames,lpppPropNames));
		// Cache the results
		if (SUCCEEDED(hRes))
		{
			AddMappingWithoutSignature(*lpcPropNames, *lpppPropNames, *lppPropTags);
		}
	}

	MAPIFreeBuffer(lpProp);
	return hRes;
} // GetNamesFromIDs

HRESULT CacheGetIDsFromNames(LPMAPIPROP lpMAPIProp,
							 ULONG cbSig,
							 LPBYTE lpSig,
							 ULONG cPropNames,
							 LPMAPINAMEID* lppPropNames,
							 ULONG ulFlags,
							 LPSPropTagArray* lppPropTags)
{
	if (!lpMAPIProp || !cPropNames || !*lppPropNames || !lppPropTags) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;

	// We're going to walk the cache, looking for the values we need. As soon as we have all the values we need, we're done
	// If we reach the end of the cache and don't have everything, we set up to make a GetIDsFromNames call.

	// First, allocate our results using MAPI
	LPSPropTagArray lpPropTags = NULL;
	EC_H(MAPIAllocateBuffer(CbNewSPropTagArray(cPropNames), (LPVOID*) &lpPropTags));

	if (lpPropTags)
	{
		memset(lpPropTags, 0, CbNewSPropTagArray(cPropNames));
		lpPropTags->cValues = cPropNames;
		ULONG ulTarget = 0;

		// Assume we'll miss on everything
		ULONG ulMisses = cPropNames;

		// For each tag we wish to look up...
		for (ulTarget = 0 ; ulTarget < cPropNames ; ulTarget++)
		{
			// ...check the cache
			LPNAMEDPROPCACHEENTRY lpEntry = FindCacheEntry(
				cbSig,
				lpSig,
				lppPropNames[ulTarget]->lpguid,
				lppPropNames[ulTarget]->ulKind,
				lppPropNames[ulTarget]->Kind.lID,
				lppPropNames[ulTarget]->Kind.lpwstrName);

			if (lpEntry)
			{
				// We have a hit - copy the data over
				lpPropTags->aulPropTag[ulTarget] = PROP_TAG(PT_UNSPECIFIED,lpEntry->ulPropID);

				// Got a hit, decrement the miss counter
				ulMisses--;
			}
		}

		// Go to MAPI with whatever's left. We set up for a single call to GetIDsFromNames.
		if (0 != ulMisses)
		{
			LPMAPINAMEID* lppUncachedPropNames = NULL;
			EC_H(MAPIAllocateBuffer(sizeof(LPMAPINAMEID) * ulMisses, (LPVOID*) &lppUncachedPropNames));
			if (lppUncachedPropNames)
			{
				memset(lppUncachedPropNames, 0, sizeof(LPMAPINAMEID) * ulMisses);
				ULONG ulUncachedName = NULL;
				for (ulTarget = 0 ; ulTarget < cPropNames ; ulTarget++)
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
				LPSPropTagArray lpUncachedTags = 0;

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
					for (ulTarget = 0 ; ulTarget < lpPropTags->cValues ; ulTarget++)
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
} // CacheGetIDsFromNames

HRESULT GetIDsFromNames(LPMAPIPROP lpMAPIProp,
						ULONG cPropNames,
						LPMAPINAMEID* lppPropNames,
						ULONG ulFlags,
						LPSPropTagArray* lppPropTags)
{
	if (!lpMAPIProp) return MAPI_E_INVALID_PARAMETER;

	// Check if we're bypassing the cache:
	if (!fCacheNamedProps() ||
	// If no names were passed, we have to bypass the cache
	// Should we cache results?
		!cPropNames || !*lppPropNames)
	{
		return lpMAPIProp->GetIDsFromNames(cPropNames,lppPropNames,ulFlags,lppPropTags);
	}
	if (!InitializeNamedPropCache()) return MAPI_E_NOT_ENOUGH_MEMORY;

	HRESULT hRes = S_OK;

	LPSPropValue lpProp = NULL;

	WC_H(HrGetOneProp(lpMAPIProp,PR_MAPPING_SIGNATURE,&lpProp));

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
		WC_H_GETPROPS(lpMAPIProp->GetIDsFromNames(cPropNames,lppPropNames,ulFlags,lppPropTags));
		// Cache the results
		if (SUCCEEDED(hRes))
		{
			AddMappingWithoutSignature(cPropNames, lppPropNames, *lppPropTags);
		}
	}

	MAPIFreeBuffer(lpProp);

	return hRes;
} // GetIDsFromNames