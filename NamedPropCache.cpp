// NamedPropCache.cpp: implementation file
//

#include "stdafx.h"
#include "NamedPropCache.h"

// We keep a vector of named prop cache entries
vector<LPNAMEDPROPCACHEENTRY> g_lpNamedPropCache;

void UninitializeNamedPropCache()
{
	for (LPNAMEDPROPCACHEENTRY& i : g_lpNamedPropCache)
	{
		delete i;
	}
}

// Go through all the details of copying allocated data to or from a cache entry
void CopyCacheData(_In_ LPGUID lpSrcGUID,
	ULONG ulSrcKind,
	LONG lSrcID,
	_In_z_ LPWSTR lpSrcName,
	_In_ LPGUID* lppDstGUID,
	_In_ ULONG* lpulDstKind,
	_In_ LONG* lplDstID,
	_In_z_ LPWSTR* lppDstName,
	_In_opt_ LPVOID lpMAPIParent) // If passed, allocate using MAPI with this as a parent
{
	if (lpSrcGUID && lppDstGUID)
	{
		LPGUID lpDstGUID = NULL;
		if (lpMAPIParent)
		{
			MAPIAllocateMore(sizeof(GUID), lpMAPIParent, (LPVOID*)&lpDstGUID);
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
			WC_H(StringCchLengthA((LPSTR)lpSrcName, STRSAFE_MAX_CCH, &cchShortLen));
			WC_H(StringCchLengthW(lpSrcName, STRSAFE_MAX_CCH, &cchWideLen));
			ULONG cbName = NULL;

			if (cchShortLen < cchWideLen)
			{
				// this is the *proper* case
				cbName = (ULONG)(cchWideLen + 1) * sizeof(WCHAR);
			}
			else
			{
				// this is the case where ANSI data was shoved into a unicode string.
				// add a couple extra NULL in case we read this as unicode again.
				cbName = (ULONG)(cchShortLen + 3) * sizeof(CHAR);
			}

			if (lpMAPIParent)
			{
				MAPIAllocateMore(cbName, lpMAPIParent, (LPVOID*)&lpDstName);
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
}

NamedPropCacheEntry::NamedPropCacheEntry(
	ULONG cbSig,
	_In_opt_count_(cbSig) LPBYTE lpSig,
	LPMAPINAMEID lpPropName,
	ULONG ulPropID)
{
	this->ulPropID = ulPropID;
	this->cbSig = cbSig;
	this->lpmniName = NULL;
	this->lpSig = NULL;
	this->bStringsCached = false;

	if (lpPropName)
	{
		lpmniName = new MAPINAMEID;

		if (lpmniName)
		{
			CopyCacheData(
				lpPropName->lpguid,
				lpPropName->ulKind,
				lpPropName->Kind.lID,
				lpPropName->Kind.lpwstrName,
				&lpmniName->lpguid,
				&lpmniName->ulKind,
				&lpmniName->Kind.lID,
				&lpmniName->Kind.lpwstrName,
				NULL);
		}
	}

	if (cbSig)
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
	for (LPNAMEDPROPCACHEENTRY& i : g_lpNamedPropCache)
	{
		if (i->ulPropID != ulPropID) continue;
		if (i->cbSig != cbSig) continue;
		if (cbSig && memcmp(lpSig, i->lpSig, cbSig)) continue;

		return i;
	}

	return NULL;
}

// Given a signature, guid, kind, and value, finds the named prop mapping in the cache
_Check_return_ LPNAMEDPROPCACHEENTRY FindCacheEntry(ULONG cbSig, _In_count_(cbSig) LPBYTE lpSig, _In_ LPGUID lpguid, ULONG ulKind, LONG lID, _In_z_ LPWSTR lpwstrName)
{
	for (LPNAMEDPROPCACHEENTRY& i : g_lpNamedPropCache)
	{
		if (!i->lpmniName) continue;
		if (i->lpmniName->ulKind != ulKind) continue;
		if (MNID_ID == ulKind && i->lpmniName->Kind.lID != lID) continue;
		if (MNID_STRING == ulKind && 0 != lstrcmpW(i->lpmniName->Kind.lpwstrName, lpwstrName)) continue;
		if (0 != memcmp(i->lpmniName->lpguid, lpguid, sizeof(GUID))) continue;
		if (cbSig != i->cbSig) continue;
		if (cbSig && memcmp(lpSig, i->lpSig, cbSig)) continue;

		return i;
	}

	return NULL;
}

// Given a tag, guid, kind, and value, finds the named prop mapping in the cache
_Check_return_ LPNAMEDPROPCACHEENTRY FindCacheEntry(ULONG ulPropID, _In_ LPGUID lpguid, ULONG ulKind, LONG lID, _In_z_ LPWSTR lpwstrName)
{
	for (LPNAMEDPROPCACHEENTRY& i : g_lpNamedPropCache)
	{
		if (i->ulPropID != ulPropID) continue;
		if (!i->lpmniName) continue;
		if (i->lpmniName->ulKind != ulKind) continue;
		if (MNID_ID == ulKind && i->lpmniName->Kind.lID != lID) continue;
		if (MNID_STRING == ulKind && 0 != lstrcmpW(i->lpmniName->Kind.lpwstrName, lpwstrName)) continue;
		if (0 != memcmp(i->lpmniName->lpguid, lpguid, sizeof(GUID))) continue;

		return i;
	}
	return NULL;
}

void AddMapping(ULONG cbSig, // Count bytes of signature
	_In_opt_count_(cbSig) LPBYTE lpSig, // Signature
	ULONG ulNumProps, // Number of mapped names
	_In_count_(ulNumProps) LPMAPINAMEID* lppPropNames, // Output from GetNamesFromIDs, input for GetIDsFromNames
	_In_ LPSPropTagArray lpTag) // Input for GetNamesFromIDs, output from GetIDsFromNames
{
	if (!ulNumProps || !lppPropNames || !lpTag) return;
	if (ulNumProps != lpTag->cValues) return; // Wouldn't know what to do with this

	ULONG ulSource = 0;
	for (ulSource = 0; ulSource < ulNumProps; ulSource++)
	{
		if (lppPropNames[ulSource])
		{
			g_lpNamedPropCache.push_back(new NamedPropCacheEntry(cbSig, lpSig, lppPropNames[ulSource], PROP_ID(lpTag->aulPropTag[ulSource])));
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

	ULONG ulProp = 0;
	SPropTagArray sptaTag = { 0 };
	sptaTag.cValues = 1;
	for (ulProp = 0; ulProp < ulNumProps; ulProp++)
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
				AddMapping(0, 0, 1, &lppPropNames[ulProp], &sptaTag);
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

	HRESULT hRes = S_OK;

	// We're going to walk the cache, looking for the values we need. As soon as we have all the values we need, we're done
	// If we reach the end of the cache and don't have everything, we set up to make a GetNamesFromIDs call.

	LPSPropTagArray lpPropTags = (*lppPropTags);
	// First, allocate our results using MAPI
	LPMAPINAMEID* lppNameIDs = NULL;
	EC_H(MAPIAllocateBuffer(sizeof(MAPINAMEID*)* lpPropTags->cValues, (LPVOID*)&lppNameIDs));

	if (lppNameIDs)
	{
		memset(lppNameIDs, 0, sizeof(MAPINAMEID*)* lpPropTags->cValues);
		ULONG ulTarget = 0;

		// Assume we'll miss on everything
		ULONG ulMisses = lpPropTags->cValues;

		// For each tag we wish to look up...
		for (ulTarget = 0; ulTarget < lpPropTags->cValues; ulTarget++)
		{
			// ...check the cache
			LPNAMEDPROPCACHEENTRY lpEntry = FindCacheEntry(cbSig, lpSig, PROP_ID(lpPropTags->aulPropTag[ulTarget]));

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
			EC_H(MAPIAllocateBuffer(CbNewSPropTagArray(ulMisses), (LPVOID*)&lpUncachedTags));
			if (lpUncachedTags)
			{
				memset(lpUncachedTags, 0, CbNewSPropTagArray(ulMisses));
				lpUncachedTags->cValues = ulMisses;
				ULONG ulUncachedTag = NULL;
				for (ulTarget = 0; ulTarget < lpPropTags->cValues; ulTarget++)
				{
					// We're looking for any result which doesn't have a mapping
					if (!lppNameIDs[ulTarget])
					{
						lpUncachedTags->aulPropTag[ulUncachedTag] = lpPropTags->aulPropTag[ulTarget];
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
					for (ulTarget = 0; ulTarget < lpPropTags->cValues; ulTarget++)
					{
						// Found an empty slot
						if (!lppNameIDs[ulTarget])
						{
							// copy the next result into it
							if (lppUncachedPropNames[ulUncachedTag])
							{
								LPMAPINAMEID lpNameID = NULL;

								EC_H(MAPIAllocateMore(sizeof(MAPINAMEID), lppNameIDs, (LPVOID*)&lpNameID));
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
		NULL,
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

	HRESULT hRes = S_OK;
	LPSPropValue lpProp = NULL;

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

	HRESULT hRes = S_OK;

	// We're going to walk the cache, looking for the values we need. As soon as we have all the values we need, we're done
	// If we reach the end of the cache and don't have everything, we set up to make a GetIDsFromNames call.

	// First, allocate our results using MAPI
	LPSPropTagArray lpPropTags = NULL;
	EC_H(MAPIAllocateBuffer(CbNewSPropTagArray(cPropNames), (LPVOID*)&lpPropTags));

	if (lpPropTags)
	{
		memset(lpPropTags, 0, CbNewSPropTagArray(cPropNames));
		lpPropTags->cValues = cPropNames;
		ULONG ulTarget = 0;

		// Assume we'll miss on everything
		ULONG ulMisses = cPropNames;

		// For each tag we wish to look up...
		for (ulTarget = 0; ulTarget < cPropNames; ulTarget++)
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
				lpPropTags->aulPropTag[ulTarget] = PROP_TAG(PT_UNSPECIFIED, lpEntry->ulPropID);

				// Got a hit, decrement the miss counter
				ulMisses--;
			}
		}

		// Go to MAPI with whatever's left. We set up for a single call to GetIDsFromNames.
		if (0 != ulMisses)
		{
			LPMAPINAMEID* lppUncachedPropNames = NULL;
			EC_H(MAPIAllocateBuffer(sizeof(LPMAPINAMEID)* ulMisses, (LPVOID*)&lppUncachedPropNames));
			if (lppUncachedPropNames)
			{
				memset(lppUncachedPropNames, 0, sizeof(LPMAPINAMEID)* ulMisses);
				ULONG ulUncachedName = NULL;
				for (ulTarget = 0; ulTarget < cPropNames; ulTarget++)
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
					for (ulTarget = 0; ulTarget < lpPropTags->cValues; ulTarget++)
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

	HRESULT hRes = S_OK;

	LPSPropValue lpProp = NULL;

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