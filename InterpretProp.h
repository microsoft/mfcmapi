#pragma once

#include <string>
using namespace std;

// Base64 functions
_Check_return_ HRESULT Base64Decode(_In_z_ LPCTSTR szEncodedStr, _Inout_ size_t* cbBuf, _Out_ _Deref_post_cap_(*cbBuf) LPBYTE* lpDecodedBuffer);
_Check_return_ HRESULT Base64Encode(size_t cbSourceBuf, _In_count_(cbSourceBuf) LPBYTE lpSourceBuffer, _Inout_ size_t* cchEncodedStr, _Out_ _Deref_post_cap_(*cchEncodedStr) LPTSTR* szEncodedStr);

void FileTimeToString(_In_ FILETIME* lpFileTime, _In_ CString *PropString, _In_opt_ CString *AltPropString);

#define TAG_MAX_LEN 1024 // Max I've seen in testing is 546 - bit more to be safe
_Check_return_ CString TagToString(ULONG ulPropTag, _In_opt_ LPMAPIPROP lpObj, bool bIsAB, bool bSingleLine);
_Check_return_ CString TypeToString(ULONG ulPropTag);
_Check_return_ CString ProblemArrayToString(_In_ LPSPropProblemArray lpProblems);
_Check_return_ CString MAPIErrToString(ULONG ulFlags, _In_ LPMAPIERROR lpErr);
_Check_return_ CString TnefProblemArrayToString(_In_ LPSTnefProblemArray lpError);

// Free with FreeNameIDStrings
void NameIDToStrings(
	ULONG ulPropTag, // optional 'original' prop tag
	_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
	_In_opt_ LPMAPINAMEID lpNameID, // optional named property information to avoid GetNamesFromIDs call
	_In_opt_ LPSBinary lpMappingSignature, // optional mapping signature for object to speed named prop lookups
	bool bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
	_Deref_opt_out_opt_z_ LPTSTR* lpszNamedPropName, // Built from ulPropTag & lpMAPIProp
	_Deref_opt_out_opt_z_ LPTSTR* lpszNamedPropGUID, // Built from ulPropTag & lpMAPIProp
	_Deref_opt_out_opt_z_ LPTSTR* lpszNamedPropDASL); // Built from ulPropTag & lpMAPIProp

void FreeNameIDStrings(_In_opt_z_ LPTSTR lpszPropName,
					   _In_opt_z_ LPTSTR lpszPropGUID,
					   _In_opt_z_ LPTSTR lpszDASL);

_Check_return_ wstring GUIDToString(_In_opt_ LPCGUID lpGUID);

_Check_return_ HRESULT StringToGUID(_In_z_ LPCTSTR szGUID, _Inout_ LPGUID lpGUID);
_Check_return_ HRESULT StringToGUID(_In_z_ LPCTSTR szGUID, bool bByteSwapped, _Inout_ LPGUID lpGUID);

_Check_return_ wstring CurrencyToString(CURRENCY curVal);

_Check_return_ wstring RestrictionToString(_In_ LPSRestriction lpRes, _In_opt_ LPMAPIPROP lpObj);
_Check_return_ wstring ActionsToString(_In_ ACTIONS* lpActions);

_Check_return_ wstring AdrListToString(_In_ LPADRLIST lpAdrList);

void InterpretProp(_In_ LPSPropValue lpProp, _In_opt_  wstring* PropString, _In_opt_  wstring* AltPropString);