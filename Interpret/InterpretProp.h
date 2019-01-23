#pragma once

namespace interpretprop
{
	std::wstring TnefProblemArrayToString(_In_ const STnefProblemArray& error);

	std::wstring RestrictionToString(_In_ const _SRestriction* lpRes, _In_opt_ LPMAPIPROP lpObj);
	std::wstring ActionsToString(_In_ const ACTIONS& actions);

	std::wstring AdrListToString(_In_ const ADRLIST& adrList);

	void InterpretProp(
		_In_ const _SPropValue* lpProp,
		_In_opt_ std::wstring* PropString,
		_In_opt_ std::wstring* AltPropString);
} // namespace interpretprop
