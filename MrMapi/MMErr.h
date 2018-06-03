#pragma once
// Error code parsing for MrMAPI

namespace error
{
	extern ERROR_ARRAY_ENTRY g_ErrorArray[];
	extern ULONG g_ulErrorArray;
}

void DoErrorParse(_In_ MYOPTIONS ProgOpts);