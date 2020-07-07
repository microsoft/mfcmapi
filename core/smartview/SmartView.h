#pragma once
#include <core/smartview/block/block.h>

// Forward declarations
enum class parserType;

namespace smartview
{
	std::shared_ptr<block> GetSmartViewParser(parserType type, _In_opt_ LPMAPIPROP lpMAPIProp);
	_Check_return_ parserType FindSmartViewParserForProp(
		_In_opt_ const ULONG ulPropTag, // required property tag
		_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
		_In_opt_ const MAPINAMEID* lpNameID, // optional named property information to avoid GetNamesFromIDs call
		_In_opt_ const SBinary* lpMappingSignature, // optional mapping signature for object to speed named prop lookups
		bool
			bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
		bool bMVRow); // did the row come from a MV prop?
	_Check_return_ parserType FindSmartViewParserForProp(
		_In_opt_ const _SPropValue* lpProp, // required property value
		_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
		_In_opt_ const MAPINAMEID* lpNameID, // optional named property information to avoid GetNamesFromIDs call
		_In_opt_ const SBinary* lpMappingSignature, // optional mapping signature for object to speed named prop lookups
		bool
			bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
		bool bMVRow); // did the row come from a MV prop?
	std::pair<ULONG, GUID> GetNamedPropInfo(
		_In_opt_ ULONG ulPropTag,
		_In_opt_ LPMAPIPROP lpMAPIProp,
		_In_opt_ const MAPINAMEID* lpNameID,
		_In_opt_ const SBinary* lpMappingSignature,
		bool bIsAB);

	std::wstring parsePropertySmartView(
		_In_opt_ const SPropValue* lpProp, // required property value
		_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
		_In_opt_ const MAPINAMEID*, // optional named property information to avoid GetNamesFromIDs call
		_In_opt_ const SBinary* lpMappingSignature, // optional mapping signature for object to speed named prop lookups
		bool
			bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
		bool bMVRow); // did the row come from a MV prop?

	std::wstring InterpretBinaryAsString(SBinary myBin, parserType parser, _In_opt_ LPMAPIPROP lpMAPIProp);
	std::wstring InterpretMVLongAsString(
		std::vector<LONG> rows,
		_In_opt_ ULONG ulPropTag,
		_In_opt_ ULONG ulPropNameID,
		_In_opt_ LPGUID lpguidNamedProp);
	std::wstring InterpretMVLongAsString(
		SLongArray myLongArray,
		_In_opt_ ULONG ulPropTag,
		_In_opt_ ULONG ulPropNameID,
		_In_opt_ LPCGUID lpguidNamedProp);
	std::wstring InterpretMVBinaryAsString(SBinaryArray myBinArray, parserType parser, _In_opt_ LPMAPIPROP lpMAPIProp);
	std::wstring InterpretNumberAsString(
		LONGLONG val,
		ULONG ulPropTag,
		ULONG ulPropNameID,
		_In_opt_z_ LPWSTR lpszPropNameString,
		_In_opt_ LPCGUID lpguidNamedProp,
		bool bLabel);
	std::wstring InterpretNumberAsString(
		_PV pV,
		ULONG ulPropTag,
		ULONG ulPropNameID,
		_In_opt_z_ LPWSTR lpszPropNameString,
		_In_opt_ LPCGUID lpguidNamedProp,
		bool bLabel);
	std::wstring InterpretNumberAsStringProp(ULONG ulVal, ULONG ulPropTag);
	std::wstring InterpretNumberAsStringNamedProp(ULONG ulVal, ULONG ulPropNameID, _In_opt_ LPCGUID lpguidNamedProp);
	_Check_return_ std::wstring RTimeToString(DWORD rTime);
	_Check_return_ std::wstring FidMidToSzString(LONGLONG llID, bool bLabel);
	std::shared_ptr<block> createBlock(parserType svParser, const SBinary bin);
} // namespace smartview