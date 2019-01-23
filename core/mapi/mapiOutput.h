#pragma once

namespace output
{
	void outputBinary(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ const SBinary& bin);
	void outputEntryList(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPENTRYLIST lpEntryList);
	void outputFormPropArray(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPIFORMPROPARRAY lpMAPIFormPropArray);
	void outputNamedPropID(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPINAMEID lpName);
	void outputPropTagArray(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSPropTagArray lpTagsToDump);
} // namespace output