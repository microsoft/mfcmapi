#include <mapistub/library/mapiStubUtils.h>

namespace mapistub
{
	// Sequence number which is incremented every time we set our MAPI handle which will
	// cause a re-fetch of all stored function pointers
	volatile ULONG g_ulDllSequenceNum = 1;
	volatile HMODULE g_hinstMAPI = nullptr;

	HMODULE GetMAPIHandle() { return g_hinstMAPI; }

	std::function<void(LPCWSTR szMsg, va_list argList)> debugPrintCallback;

	void __cdecl DebugPrint(LPCWSTR szMsg, ...)
	{
		if (debugPrintCallback)
		{
			va_list argList = nullptr;
			va_start(argList, szMsg);
			debugPrintCallback(szMsg, argList);
			va_end(argList);
		}
	}
} // namespace mapistub
