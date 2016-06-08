#pragma once

// Base64 functions
_Check_return_ HRESULT Base64Decode(_In_z_ LPCTSTR szEncodedStr, _Inout_ size_t* cbBuf, _Out_ _Deref_post_cap_(*cbBuf) LPBYTE* lpDecodedBuffer);
_Check_return_ HRESULT Base64Encode(size_t cbSourceBuf, _In_count_(cbSourceBuf) LPBYTE lpSourceBuffer, _Inout_ size_t* cchEncodedStr, _Out_ _Deref_post_cap_(*cchEncodedStr) LPTSTR* szEncodedStr);

void FileTimeToString(_In_ FILETIME* lpFileTime, _In_ wstring& PropString, _In_opt_ wstring& AltPropString);

#define TAG_MAX_LEN 1024 // Max I've seen in testing is 546 - bit more to be safe
wstring TagToString(ULONG ulPropTag, _In_opt_ LPMAPIPROP lpObj, bool bIsAB, bool bSingleLine);
wstring TypeToString(ULONG ulPropTag);
wstring ProblemArrayToString(_In_ LPSPropProblemArray lpProblems);
wstring MAPIErrToString(ULONG ulFlags, _In_ LPMAPIERROR lpErr);
wstring TnefProblemArrayToString(_In_ LPSTnefProblemArray lpError);

void NameIDToStrings(
	ULONG ulPropTag, // optional 'original' prop tag
	_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
	_In_opt_ LPMAPINAMEID lpNameID, // optional named property information to avoid GetNamesFromIDs call
	_In_opt_ LPSBinary lpMappingSignature, // optional mapping signature for object to speed named prop lookups
	bool bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
	_In_ wstring& lpszNamedPropName, // Built from ulPropTag & lpMAPIProp
	_In_ wstring& lpszNamedPropGUID, // Built from ulPropTag & lpMAPIProp
	_In_ wstring& lpszNamedPropDASL); // Built from ulPropTag & lpMAPIProp

wstring CurrencyToString(CURRENCY curVal);

wstring RestrictionToString(_In_ LPSRestriction lpRes, _In_opt_ LPMAPIPROP lpObj);
wstring ActionsToString(_In_ ACTIONS* lpActions);

wstring AdrListToString(_In_ LPADRLIST lpAdrList);

void InterpretProp(_In_ LPSPropValue lpProp, _In_opt_ wstring* PropString, _In_opt_ wstring* AltPropString);