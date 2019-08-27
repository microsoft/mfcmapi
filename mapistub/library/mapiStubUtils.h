#pragma once
#include <Windows.h>

namespace mapistub
{
	extern volatile ULONG g_ulDllSequenceNum;
	HMODULE GetMAPIHandle();
	HMODULE GetPrivateMAPI();
} // namespace mapistub
