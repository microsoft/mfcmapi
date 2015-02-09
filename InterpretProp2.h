#pragma once
#include <vector>
#include <string>
using namespace std;

#define PROP_TAG_MASK 0xffff0000
void FindTagArrayMatches(_In_ ULONG ulTarget,
	bool bIsAB,
	_In_count_(ulMyArray) NAME_ARRAY_ENTRY_V2* MyArray,
	_In_ ULONG ulMyArray,
	vector<ULONG>& ulExacts,
	vector<ULONG>& ulPartials);

// Function to convert property tags to their names
// Free lpszExactMatch and lpszPartialMatches with MAPIFreeBuffer
void PropTagToPropName(ULONG ulPropTag, bool bIsAB, _Deref_opt_out_opt_z_ LPTSTR* lpszExactMatch, _Deref_opt_out_opt_z_ LPTSTR* lpszPartialMatches);

// Strictly does a lookup in the array. Does not convert otherwise
_Check_return_ HRESULT LookupPropName(_In_z_ LPCWSTR lpszPropName, _Out_ ULONG* ulPropTag);

_Check_return_ HRESULT PropNameToPropTagW(_In_z_ LPCWSTR lpszPropName, _Out_ ULONG* ulPropTag);
_Check_return_ HRESULT PropNameToPropTagA(_In_z_ LPCSTR lpszPropName, _Out_ ULONG* ulPropTag);
#ifdef UNICODE
#define PropNameToPropTag PropNameToPropTagW
#else
#define PropNameToPropTag PropNameToPropTagA
#endif
_Check_return_ ULONG PropTypeNameToPropTypeW(_In_z_ LPCWSTR lpszPropType);
_Check_return_ ULONG PropTypeNameToPropTypeA(_In_z_ LPCSTR lpszPropType);
#ifdef UNICODE
#define PropTypeNameToPropType PropTypeNameToPropTypeW
#else
#define PropTypeNameToPropType PropTypeNameToPropTypeA
#endif

_Check_return_ wstring GUIDToString(_In_opt_ LPCGUID lpGUID);
_Check_return_ wstring GUIDToStringAndName(_In_opt_ LPCGUID lpGUID);
LPCGUID GUIDNameToGUIDW(_In_z_ LPCWSTR szGUID, bool bByteSwapped);
LPCGUID GUIDNameToGUIDA(_In_z_ LPCSTR szGUID, bool bByteSwapped);
#ifdef UNICODE
#define GUIDNameToGUID GUIDNameToGUIDW
#else
#define GUIDNameToGUID GUIDNameToGUIDA
#endif

wstring NameIDToPropName(_In_ LPMAPINAMEID lpNameID);

_Check_return_  wstring InterpretFlags(const ULONG ulFlagName, const LONG lFlagValue);
_Check_return_ CString AllFlagsToString(const ULONG ulFlagName, bool bHex);

_Check_return_ HRESULT GetLargeBinaryProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_ LPSPropValue* lppProp);
_Check_return_ HRESULT GetLargeStringProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_ LPSPropValue* lppProp);