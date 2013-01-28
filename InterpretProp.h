#pragma once

// Base64 functions
_Check_return_ HRESULT Base64Decode(_In_z_ LPCTSTR szEncodedStr, _Inout_ size_t* cbBuf, _Out_ _Deref_post_cap_(*cbBuf) LPBYTE* lpDecodedBuffer);
_Check_return_ HRESULT Base64Encode(size_t cbSourceBuf, _In_count_(cbSourceBuf) LPBYTE lpSourceBuffer, _Inout_ size_t* cchEncodedStr, _Out_ _Deref_post_cap_(*cchEncodedStr) LPTSTR* szEncodedStr);

// Function to create strings representing properties
_Check_return_ CString BinToTextString(_In_ LPSBinary lpBin, bool bMultiLine);

_Check_return_ CString BinToHexString(_In_opt_ LPSBinary lpBin, bool bPrependCB);

void FileTimeToString(_In_ FILETIME* lpFileTime, _In_ CString *PropString, _In_opt_ CString *AltPropString);

#define TAG_MAX_LEN 1024 // Max I've seen in testing is 546 - bit more to be safe
_Check_return_ CString TagToString(ULONG ulPropTag, _In_opt_ LPMAPIPROP lpObj, bool bIsAB, bool bSingleLine);
_Check_return_ CString TypeToString(ULONG ulPropTag);
_Check_return_ CString ProblemArrayToString(_In_ LPSPropProblemArray lpProblems);
_Check_return_ CString MAPIErrToString(ULONG ulFlags, _In_ LPMAPIERROR lpErr);
_Check_return_ CString TnefProblemArrayToString(_In_ LPSTnefProblemArray lpError);

// Allocates strings with new
// Free with FreeNameIDStrings
//
// lpszDASL string for a named prop will look like this:
// id/{12345678-1234-1234-1234-12345678ABCD}/80010003
// string/{12345678-1234-1234-1234-12345678ABCD}/MyProp
// So the following #defines give the size of the buffers we need, in TCHARS, including 1 for the null terminator
// CCH_DASL_ID gets an extra digit to handle some AB props with name IDs of five digits
#define CCH_DASL_ID 2+1+38+1+8+1+1
#define CCH_DASL_STRING 6+1+38+1+1
// TagToString will prepend the http://schemas.microsoft.com/MAPI/ for us since it's a constant
// We don't compute a DASL string for non-named props as FormatMessage in TagToString can handle those
void NameIDToStrings(_In_ LPMAPINAMEID lpNameID,
					 ULONG ulPropTag,
					 _Deref_opt_out_opt_z_ LPTSTR* lpszPropName,
					 _Deref_opt_out_opt_z_ LPTSTR* lpszPropGUID,
					 _Deref_opt_out_opt_z_ LPTSTR* lpszDASL);
void FreeNameIDStrings(_In_opt_z_ LPTSTR lpszPropName,
					   _In_opt_z_ LPTSTR lpszPropGUID,
					   _In_opt_z_ LPTSTR lpszDASL);

_Check_return_ LPTSTR GUIDToString(_In_opt_ LPCGUID lpGUID);

_Check_return_ HRESULT StringToGUID(_In_z_ LPCTSTR szGUID, _Inout_ LPGUID lpGUID);
_Check_return_ HRESULT StringToGUID(_In_z_ LPCTSTR szGUID, bool bByteSwapped, _Inout_ LPGUID lpGUID);

_Check_return_ CString CurrencyToString(CURRENCY curVal);

_Check_return_ CString RestrictionToString(_In_ LPSRestriction lpRes, _In_opt_ LPMAPIPROP lpObj);

void ActionsToString(_In_ ACTIONS* lpActions, _In_ CString* PropString);

void AdrListToString(_In_ LPADRLIST lpAdrList, _In_ CString *PropString);

void InterpretMVProp(_In_ LPSPropValue lpProp, ULONG ulMVRow, _In_ CString *PropString, _In_ CString *AltPropString);
void InterpretProp(_In_ LPSPropValue lpProp, _In_opt_ CString *PropString, _In_opt_ CString *AltPropString);