#pragma once
#include <core/utility/output.h>

namespace output
{
	// Template for the Output functions
	// void Output(ULONG ulDbgLvl, FILE* fFile,stufftooutput)

	// If the first parameter is not DBGNoDebug, we debug print the output
	// If the second parameter is a file name, we print the output to the file
	// Doing both is OK, if we ask to do neither, ASSERT

	// We'll use macros to make these calls so the code will read right

	void outputBinary(DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile, _In_ const SBinary& bin);
	void outputEntryList(DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPENTRYLIST lpEntryList);
	void outputFormPropArray(DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPIFORMPROPARRAY lpMAPIFormPropArray);
	void outputNamedPropID(DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPINAMEID lpName);
	void outputPropTagArray(DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSPropTagArray lpTagsToDump);

	void outputProperty(
		DBGLEVEL ulDbgLvl,
		_In_opt_ FILE* fFile,
		_In_ LPSPropValue lpProp,
		_In_opt_ LPMAPIPROP lpObj,
		bool bRetryStreamProps);
	void outputProperties(
		DBGLEVEL ulDbgLvl,
		_In_opt_ FILE* fFile,
		ULONG cProps,
		_In_count_(cProps) LPSPropValue lpProps,
		_In_opt_ LPMAPIPROP lpObj,
		bool bRetryStreamProps);
	void outputSRow(DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile, _In_ const _SRow* lpSRow, _In_opt_ LPMAPIPROP lpObj);
	void outputSRowSet(DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSRowSet lpRowSet, _In_opt_ LPMAPIPROP lpObj);
	void outputRestriction(
		DBGLEVEL ulDbgLvl,
		_In_opt_ FILE* fFile,
		_In_opt_ const _SRestriction* lpRes,
		_In_opt_ LPMAPIPROP lpObj);
	void outputFormInfo(DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPIFORMINFO lpMAPIFormInfo);
	void outputTable(DBGLEVEL ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPITABLE lpMAPITable);
	void outputNotifications(
		DBGLEVEL ulDbgLvl,
		_In_opt_ FILE* fFile,
		ULONG cNotify,
		_In_count_(cNotify) LPNOTIFICATION lpNotifications,
		_In_opt_ LPMAPIPROP lpObj);
} // namespace output