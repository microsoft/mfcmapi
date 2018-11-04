#pragma once
#include <MFCMAPI.h>

namespace interpretprop
{
	std::wstring TagToString(ULONG ulPropTag, _In_opt_ LPMAPIPROP lpObj, bool bIsAB, bool bSingleLine);
	std::wstring TypeToString(ULONG ulPropTag);
	std::wstring ProblemArrayToString(_In_ const SPropProblemArray& problems);
	std::wstring MAPIErrToString(ULONG ulFlags, _In_ const MAPIERROR& err);
	std::wstring TnefProblemArrayToString(_In_ const STnefProblemArray& error);

	std::wstring RestrictionToString(_In_ const _SRestriction* lpRes, _In_opt_ LPMAPIPROP lpObj);
	std::wstring ActionsToString(_In_ const ACTIONS& actions);

	std::wstring AdrListToString(_In_ const ADRLIST& adrList);

	void InterpretProp(
		_In_ const _SPropValue* lpProp,
		_In_opt_ std::wstring* PropString,
		_In_opt_ std::wstring* AltPropString);

#define PROP_TAG_MASK 0xffff0000
	void FindTagArrayMatches(
		_In_ ULONG ulTarget,
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

	std::vector<std::wstring> NameIDToPropNames(_In_ const MAPINAMEID* lpNameID);

	std::wstring InterpretFlags(ULONG ulFlagName, LONG lFlagValue);
	std::wstring AllFlagsToString(ULONG ulFlagName, bool bHex);
} // namespace interpretprop
