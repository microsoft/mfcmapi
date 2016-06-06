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
_Check_return_ HRESULT LookupPropName(_In_ wstring lpszPropName, _Out_ ULONG* ulPropTag);

_Check_return_ HRESULT PropNameToPropTag(_In_ wstring lpszPropName, _Out_ ULONG* ulPropTag);
_Check_return_ ULONG PropTypeNameToPropType(_In_ wstring lpszPropType);

wstring GUIDToString(_In_opt_ LPCGUID lpGUID);
wstring GUIDToStringAndName(_In_opt_ LPCGUID lpGUID);
LPCGUID GUIDNameToGUID(_In_ wstring szGUID, bool bByteSwapped);

wstring NameIDToPropName(_In_ LPMAPINAMEID lpNameID);

wstring InterpretFlags(const ULONG ulFlagName, const LONG lFlagValue);
wstring AllFlagsToString(const ULONG ulFlagName, bool bHex);

_Check_return_ HRESULT GetLargeBinaryProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_ LPSPropValue* lppProp);
_Check_return_ HRESULT GetLargeStringProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_ LPSPropValue* lppProp);