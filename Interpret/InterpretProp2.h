#pragma once

#define PROP_TAG_MASK 0xffff0000
void FindTagArrayMatches(_In_ ULONG ulTarget,
	bool bIsAB,
	const std::vector<NAME_ARRAY_ENTRY_V2>& MyArray,
	std::vector<ULONG>& ulExacts,
	std::vector<ULONG>& ulPartials);

struct PropTagNames
{
	std::wstring bestGuess;
	std::wstring otherMatches;
};

// Function to convert property tags to their names
PropTagNames PropTagToPropName(ULONG ulPropTag, bool bIsAB);

// Strictly does a lookup in the array. Does not convert otherwise
_Check_return_ ULONG LookupPropName(_In_ const std::wstring& lpszPropName);
_Check_return_ ULONG PropNameToPropTag(_In_ const std::wstring& lpszPropName);
_Check_return_ ULONG PropTypeNameToPropType(_In_ const std::wstring& lpszPropType);

std::wstring GUIDToString(_In_opt_ LPCGUID lpGUID);
std::wstring GUIDToStringAndName(_In_opt_ LPCGUID lpGUID);
LPCGUID GUIDNameToGUID(_In_ const std::wstring& szGUID, bool bByteSwapped);
_Check_return_ GUID StringToGUID(_In_ const std::wstring& szGUID);
_Check_return_ GUID StringToGUID(_In_ const std::wstring& szGUID, bool bByteSwapped);

std::vector<std::wstring> NameIDToPropNames(_In_ const LPMAPINAMEID lpNameID);

std::wstring InterpretFlags(ULONG ulFlagName, LONG lFlagValue);
std::wstring AllFlagsToString(ULONG ulFlagName, bool bHex);

_Check_return_ HRESULT GetLargeBinaryProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_ LPSPropValue* lppProp);
_Check_return_ HRESULT GetLargeStringProp(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_ LPSPropValue* lppProp);