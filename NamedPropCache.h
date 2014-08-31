#pragma once
// Named Property Cache
#include <string>

class NamedPropCacheEntry
{
public:
	NamedPropCacheEntry::NamedPropCacheEntry(
		ULONG cbSig,
		_In_opt_count_(cbSig) LPBYTE lpSig,
		LPMAPINAMEID lpPropName,
		ULONG ulPropID);
	~NamedPropCacheEntry();

	ULONG ulPropID;         // MAPI ID (ala PROP_ID) for a named property
	LPMAPINAMEID lpmniName; // guid, kind, value
	ULONG cbSig;            // Size and...
	LPBYTE lpSig;           // Value of PR_MAPPING_SIGNATURE
	bool bStringsCached;    // We have cached strings
	std::wstring lpszPropName;    // Cached strings
	std::wstring lpszPropGUID;    //
	std::wstring lpszDASL;        //

private:
};
typedef NamedPropCacheEntry *LPNAMEDPROPCACHEENTRY;

void UninitializeNamedPropCache();

_Check_return_ LPNAMEDPROPCACHEENTRY FindCacheEntry(ULONG ulPropID, _In_ LPGUID lpguid, ULONG ulKind, LONG lID, _In_z_ LPWSTR lpwstrName);

_Check_return_ HRESULT GetNamesFromIDs(_In_ LPMAPIPROP lpMAPIProp,
	_In_ LPSPropTagArray* lppPropTags,
	_In_opt_ LPGUID lpPropSetGuid,
	ULONG ulFlags,
	_Out_ ULONG* lpcPropNames,
	_Out_ _Deref_post_cap_(*lpcPropNames) LPMAPINAMEID** lpppPropNames);
_Check_return_ HRESULT GetNamesFromIDs(_In_ LPMAPIPROP lpMAPIProp,
	_In_opt_ LPSBinary lpMappingSignature,
	_In_ LPSPropTagArray* lppPropTags,
	_In_opt_ LPGUID lpPropSetGuid,
	ULONG ulFlags,
	_Out_ ULONG* lpcPropNames,
	_Out_ _Deref_post_cap_(*lpcPropNames) LPMAPINAMEID** lpppPropNames);
_Check_return_ HRESULT GetIDsFromNames(_In_ LPMAPIPROP lpMAPIProp,
	ULONG cPropNames,
	_In_opt_count_(cPropNames) LPMAPINAMEID* lppPropNames,
	ULONG ulFlags,
	_Out_ _Deref_post_cap_(cPropNames) LPSPropTagArray* lppPropTags);

_Check_return_ inline bool fCacheNamedProps()
{
	return RegKeys[regkeyCACHE_NAME_DPROPS].ulCurDWORD != 0;
}