#pragma once

namespace flags
{
	std::wstring InterpretFlags(ULONG ulFlagName, LONG lFlagValue);
	std::wstring AllFlagsToString(ULONG ulFlagName, bool bHex);
} // namespace flags
