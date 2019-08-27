#pragma once
#include <Windows.h>

namespace mapistub
{
	extern volatile ULONG g_ulDllSequenceNum;
	static volatile HMODULE g_hinstMAPI = nullptr;
	HMODULE GetMAPIHandle();
	HMODULE GetPrivateMAPI();
} // namespace mapistub
