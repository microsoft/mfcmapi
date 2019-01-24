#pragma once
#include <core/property/property.h>

namespace property
{
	void parseProperty(
		_In_ const _SPropValue* lpProp,
		_In_opt_ std::wstring* PropString,
		_In_opt_ std::wstring* AltPropString);

	Property parseProperty(_In_ const _SPropValue* lpProp);

	std::wstring RestrictionToString(_In_ const _SRestriction* lpRes, _In_opt_ LPMAPIPROP lpObj);
	std::wstring ActionsToString(_In_ const ACTIONS& actions);

	std::wstring AdrListToString(_In_ const ADRLIST& adrList);
} // namespace property