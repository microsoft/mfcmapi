#pragma once
// Named Property Cache

// NamedPropCacheEntry
// =====================
//   An entry in the named property cache
typedef struct _NamedPropCacheEntry
{
	ULONG ulPropID;         // MAPI ID (ala PROP_ID) for a named property
	LPMAPINAMEID lpmniName; // guid, kind, value
	ULONG cbSig;            // Size and...
	LPBYTE lpSig;           // Value of PR_MAPPING_SIGNATURE
	BOOL bStringsCached;    // We have cached strings
	LPTSTR lpszPropName;    // Cached strings
	LPTSTR lpszPropGUID;    //
	LPTSTR lpszDASL;        //
} NamedPropCacheEntry, * LPNAMEDPROPCACHEENTRY;

// Cache initializes on demand. Uninitialize on program shutdown.
void UninitializeNamedPropCache();

LPNAMEDPROPCACHEENTRY FindCacheEntry(ULONG ulPropID, LPGUID lpguid, ULONG ulKind, LONG lID, LPWSTR lpwstrName);

HRESULT GetNamesFromIDs(LPMAPIPROP lpMAPIProp,
						LPSPropTagArray* lppPropTags,
						LPGUID lpPropSetGuid,
						ULONG ulFlags,
						ULONG* lpcPropNames,
						LPMAPINAMEID** lpppPropNames);
HRESULT GetNamesFromIDs(LPMAPIPROP lpMAPIProp,
						LPSBinary lpMappingSignature,
						LPSPropTagArray* lppPropTags,
						LPGUID lpPropSetGuid,
						ULONG ulFlags,
						ULONG* lpcPropNames,
						LPMAPINAMEID** lpppPropNames);
HRESULT GetIDsFromNames(LPMAPIPROP lpMAPIProp,
						ULONG cPropNames,
						LPMAPINAMEID* lppPropNames,
						ULONG ulFlags,
						LPSPropTagArray* lppPropTags);

inline BOOL fCacheNamedProps()
{
    return RegKeys[regkeyCACHE_NAME_DPROPS].ulCurDWORD != 0;
}