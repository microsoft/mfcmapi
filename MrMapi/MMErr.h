#pragma once
// Error code parsing for MrMAPI

namespace error
{
	extern ERROR_ARRAY_ENTRY g_ErrorArray[];
	extern ULONG g_ulErrorArray;
} // namespace error

void DoErrorParse(_In_ cli::MYOPTIONS ProgOpts);