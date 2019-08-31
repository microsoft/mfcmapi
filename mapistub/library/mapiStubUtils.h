#pragma once
#include <Windows.h>
#include <functional>

namespace mapistub
{
	extern std::function<void(LPCWSTR szMsg, va_list argList)> logLoadMapiCallback;
	void __cdecl logLoadMapi(LPCWSTR szMsg, ...);

	extern volatile ULONG g_ulDllSequenceNum;
	extern volatile HMODULE g_hinstMAPI;
	HMODULE GetMAPIHandle();
	HMODULE GetPrivateMAPI();
} // namespace mapistub
