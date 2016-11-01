#pragma once

// Base64 functions
vector<BYTE> Base64Decode(wstring const& szEncodedStr);
wstring Base64Encode(size_t cbSourceBuf, _In_count_(cbSourceBuf) LPBYTE lpSourceBuffer);

void FileTimeToString(_In_ const FILETIME& fileTime, _In_ wstring& PropString, _In_opt_ wstring& AltPropString);

wstring TagToString(ULONG ulPropTag, _In_opt_ LPMAPIPROP lpObj, bool bIsAB, bool bSingleLine);
wstring TypeToString(ULONG ulPropTag);
wstring ProblemArrayToString(_In_ const SPropProblemArray& problems);
wstring MAPIErrToString(ULONG ulFlags, _In_ const MAPIERROR& err);
wstring TnefProblemArrayToString(_In_ const STnefProblemArray& error);

struct NamePropNames
{
	wstring name;
	wstring guid;
	wstring dasl;
	wstring bestPidLid;
	wstring otherPidLid;
};

NamePropNames NameIDToStrings(
	ULONG ulPropTag, // optional 'original' prop tag
	_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
	_In_opt_ LPMAPINAMEID lpNameID, // optional named property information to avoid GetNamesFromIDs call
	_In_opt_ LPSBinary lpMappingSignature, // optional mapping signature for object to speed named prop lookups
	bool bIsAB); // true if we know we're dealing with an address book property (they can be > 8000 and not named props)

wstring CurrencyToString(const CURRENCY& curVal);

wstring RestrictionToString(_In_ const LPSRestriction lpRes, _In_opt_ LPMAPIPROP lpObj);
wstring ActionsToString(_In_ const ACTIONS& actions);

wstring AdrListToString(_In_ const ADRLIST& adrList);

void InterpretProp(_In_ LPSPropValue lpProp, _In_opt_ wstring* PropString, _In_opt_ wstring* AltPropString);