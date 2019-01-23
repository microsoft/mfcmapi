#pragma once

namespace proptype
{
	std::wstring TypeToString(ULONG ulPropTag);
	_Check_return_ ULONG PropTypeNameToPropType(_In_ const std::wstring& lpszPropType);
} // namespace proptype