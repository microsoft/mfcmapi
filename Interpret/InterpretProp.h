#pragma once

namespace interpretprop
{
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

	// Strictly does a lookup in the array. Does not convert otherwise
	_Check_return_ ULONG LookupPropName(_In_ const std::wstring& lpszPropName);
	_Check_return_ ULONG PropNameToPropTag(_In_ const std::wstring& lpszPropName);
	_Check_return_ ULONG PropTypeNameToPropType(_In_ const std::wstring& lpszPropType);
} // namespace interpretprop
