#pragma once
#include <Windows.h>

namespace mapistub
{
	extern volatile ULONG g_ulDllSequenceNum;
	extern volatile HMODULE g_hinstMAPI;
	HMODULE GetMAPIHandle();
	HMODULE GetPrivateMAPI();
} // namespace mapistub
