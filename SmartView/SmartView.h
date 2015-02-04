#pragma once
#include <string>
using namespace std;

wstring InterpretPropSmartView(_In_ LPSPropValue lpProp, // required property value
	_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
	_In_opt_ LPMAPINAMEID lpNameID, // optional named property information to avoid GetNamesFromIDs call
	_In_opt_ LPSBinary lpMappingSignature, // optional mapping signature for object to speed named prop lookups
	bool bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
	bool bMVRow); // did the row come from a MV prop?

wstring InterpretBinaryAsString(SBinary myBin, __ParsingTypeEnum iStructType, _In_opt_ LPMAPIPROP lpMAPIProp);
wstring InterpretMVBinaryAsString(SBinaryArray myBinArray, __ParsingTypeEnum iStructType, _In_opt_ LPMAPIPROP lpMAPIProp);
wstring InterpretNumberAsString(_PV pV, ULONG ulPropTag, ULONG ulPropNameID, _In_opt_z_ LPWSTR lpszPropNameString, _In_opt_ LPCGUID lpguidNamedProp, bool bLabel);
wstring InterpretNumberAsStringProp(ULONG ulVal, ULONG ulPropTag);
wstring InterpretNumberAsStringNamedProp(ULONG ulVal, ULONG ulPropNameID, _In_opt_ LPCGUID lpguidNamedProp);
_Check_return_ wstring RTimeToString(DWORD rTime);
_Check_return_ wstring FidMidToSzString(LONGLONG llID, bool bLabel);

_Check_return_ __ParsingTypeEnum FindSmartViewParserForProp(const ULONG ulPropTag, const ULONG ulPropNameID, _In_opt_ const LPCGUID lpguidNamedProp, bool bMVRow);