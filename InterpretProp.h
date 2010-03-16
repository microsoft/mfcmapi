#pragma once

// Base64 functions
HRESULT Base64Decode(LPCTSTR szEncodedStr, size_t* cbBuf, LPBYTE* lpDecodedBuffer);
HRESULT Base64Encode(size_t cbSourceBuf, LPBYTE lpSourceBuffer, size_t* cchEncodedStr, LPTSTR* szEncodedStr);

// Function to create strings representing properties
CString BinToTextString(LPSBinary lpBin, BOOL bMultiLine);

CString BinToHexString(LPSBinary lpBin, BOOL bPrependCB);

void FileTimeToString(FILETIME* lpFileTime, CString *PropString, CString *AltPropString);
CString RTimeToString(DWORD rTime);

#define TAG_MAX_LEN 1024 // Max I've seen in testing is 546 - bit more to be safe
CString TagToString(ULONG ulPropTag, LPMAPIPROP lpObj, BOOL bIsAB, BOOL bSingleLine);
CString TypeToString(ULONG ulPropTag);
CString ProblemArrayToString(LPSPropProblemArray lpProblems);
CString MAPIErrToString(ULONG ulFlags, LPMAPIERROR lpErr);
CString TnefProblemArrayToString(LPSTnefProblemArray lpError);

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
void NameIDToStrings(LPMAPINAMEID lpNameID,
					 ULONG ulPropTag,
					 LPTSTR* lpszPropName,
					 LPTSTR* lpszPropGUID,
					 LPTSTR* lpszDASL);
void FreeNameIDStrings(LPTSTR lpszPropName,
					   LPTSTR lpszPropGUID,
					   LPTSTR lpszDASL);

LPTSTR GUIDToString(LPCGUID lpGUID);

HRESULT StringToGUID(LPCTSTR szGUID, LPGUID lpGUID);

CString CurrencyToString(CURRENCY curVal);

CString RestrictionToString(LPSRestriction lpRes,LPMAPIPROP lpObj);

void ActionsToString(ACTIONS* lpActions, CString* PropString);

void AdrListToString(LPADRLIST lpAdrList,CString *PropString);

void InterpretMVProp(LPSPropValue lpProp, ULONG ulMVRow, CString *PropString, CString *AltPropString);
void InterpretProp(LPSPropValue lpProp, CString *PropString, CString *AltPropString);