#pragma once

#include <MapiX.h>

//Base64 functions
HRESULT Base64Decode(LPCTSTR szEncodedStr, size_t* cbBuf, LPBYTE* lpDecodedBuffer);
HRESULT Base64Encode(size_t cbSourceBuf, LPBYTE lpSourceBuffer, size_t* cchEncodedStr, LPTSTR* szEncodedStr);

//Function to create strings representing properties
CString BinToTextString(LPSBinary lpBin, BOOL bMultiLine);

CString BinToHexString(LPSBinary lpBin, BOOL bPrependCB);

void FileTimeToString(FILETIME* lpFileTime, CString *PropString, CString *AltPropString);

#define TAG_MAX_LEN 1024 // Max I've seen in testing is 546 - bit more to be safe
CString TagToString(ULONG ulPropTag, LPMAPIPROP lpObj, BOOL bIsAB, BOOL bSingleLine);
CString TypeToString(ULONG ulPropTag);
CString ProblemArrayToString(LPSPropProblemArray lpProblems);
CString MAPIErrToString(ULONG ulFlags, LPMAPIERROR lpErr);
CString TnefProblemArrayToString(LPSTnefProblemArray lpError);

void GetPropName(LPMAPIPROP lpMAPIProp,
				 ULONG ulPropTag,
				 LPTSTR *lpszPropName,
				 LPTSTR *lpszPropGUID);

void GetPropName(LPMAPIPROP lpMAPIProp,
				 ULONG ulPropTag,
				 LPTSTR *lpszPropName,
				 LPTSTR *lpszPropGUID,
				 LPTSTR *lpszDASL);

LPTSTR GUIDToString(LPCGUID lpGUID);

HRESULT StringToGUID(LPCTSTR szGUID, LPGUID lpGUID);

CString CurrencyToString(CURRENCY curVal);

CString RestrictionToString(LPSRestriction lpRes,LPMAPIPROP lpObj);

void ActionsToString(ACTIONS* lpActions, CString* PropString);

CString EntryListToString(LPENTRYLIST lpEntryList);
void AdrListToString(LPADRLIST lpAdrList,CString *PropString);

void InterpretMVProp(LPSPropValue lpProp, ULONG ulMVRow, CString *PropString, CString *AltPropString);
void InterpretProp(LPSPropValue lpProp, CString *PropString, CString *AltPropString);