#pragma once

#define PROP_TAG_MASK 0xffff0000
void FindTagArrayMatches(_In_ ULONG ulTarget,
	bool bIsAB,
	_In_count_(ulMyArray) NAME_ARRAY_ENTRY_V2* MyArray,
	_In_ ULONG ulMyArray,
	vector<ULONG>& ulExacts,
	vector<ULONG>& ulPartials);

// Function to convert property tags to their names
void PropTagToPropName(ULONG ulPropTag, bool bIsAB, _In_opt_  wstring& lpszExactMatch, _In_opt_  wstring& lpszPartialMatches);

// Strictly does a lookup in the array. Does not convert otherwise
_Check_return_ ULONG LookupPropName(_In_ wstring lpszPropName);
_Check_return_ ULONG PropNameToPropTag(_In_ wstring lpszPropName);
_Check_return_ ULONG PropTypeNameToPropType(_In_ wstring lpszPropType);

wstring GUIDToString(_In_opt_ LPCGUID lpGUID);
wstring GUIDToStringAndName(_In_opt_ LPCGUID lpGUID);
LPCGUID GUIDNameToGUID(_In_ wstring szGUID, bool bByteSwapped);
_Check_return_ GUID StringToGUID(_In_ wstring szGUID);
_Check_return_ GUID StringToGUID(_In_ wstring szGUID, bool bByteSwapped);

wstring NameIDToPropName(_In_ LPMAPINAMEID lpNameID);

wstring InterpretFlags(const ULONG ulFlagName, const LONG lFlagValue);
wstring AllFlagsToString(const ULONG ulFlagName, bool bHex);

_Check_return_ HRESULT GetLargeBinaryProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_ LPSPropValue* lppProp);
_Check_return_ HRESULT GetLargeStringProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_ LPSPropValue* lppProp);