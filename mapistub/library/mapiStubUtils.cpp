#include <mapistub/library/mapiStubUtils.h>

namespace mapistub
{
	// Sequence number which is incremented every time we set our MAPI handle which will
	// cause a re-fetch of all stored function pointers
	volatile ULONG g_ulDllSequenceNum = 1;
	volatile HMODULE g_hinstMAPI = nullptr;

	HMODULE GetMAPIHandle() { return g_hinstMAPI; }

	std::function<void(LPCWSTR szMsg, va_list argList)> logLoadMapiCallback;
	std::function<void(LPCWSTR szMsg, va_list argList)> logLoadLibraryCallback;

	void __cdecl logLoadMapi(LPCWSTR szMsg, ...)
	{
		if (logLoadMapiCallback)
		{
			va_list argList = nullptr;
			va_start(argList, szMsg);
			logLoadMapiCallback(szMsg, argList);
			va_end(argList);
		}
	}

	void __cdecl logLoadLibrary(LPCWSTR szMsg, ...)
	{
		if (logLoadLibraryCallback)
		{
			va_list argList = nullptr;
			va_start(argList, szMsg);
			logLoadLibraryCallback(szMsg, argList);
			va_end(argList);
		}
	}
} // namespace mapistub
