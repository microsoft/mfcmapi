#pragma once
// Output (to File/Debug) functions

namespace output
{
	// Template for the Output functions
	// void Output(ULONG ulDbgLvl, FILE* fFile,stufftooutput)

	// If the first parameter is not DBGNoDebug, we debug print the output
	// If the second parameter is a file name, we print the output to the file
	// Doing both is OK, if we ask to do neither, ASSERT

	// We'll use macros to make these calls so the code will read right

	void _OutputProperty(
		ULONG ulDbgLvl,
		_In_opt_ FILE* fFile,
		_In_ LPSPropValue lpProp,
		_In_opt_ LPMAPIPROP lpObj,
		bool bRetryStreamProps);
	void _OutputProperties(
		ULONG ulDbgLvl,
		_In_opt_ FILE* fFile,
		ULONG cProps,
		_In_count_(cProps) LPSPropValue lpProps,
		_In_opt_ LPMAPIPROP lpObj,
		bool bRetryStreamProps);
	void _OutputSRow(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ const _SRow* lpSRow, _In_opt_ LPMAPIPROP lpObj);
	void _OutputSRowSet(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPSRowSet lpRowSet, _In_opt_ LPMAPIPROP lpObj);
	void _OutputRestriction(
		ULONG ulDbgLvl,
		_In_opt_ FILE* fFile,
		_In_opt_ const _SRestriction* lpRes,
		_In_opt_ LPMAPIPROP lpObj);
	void _OutputFormInfo(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPIFORMINFO lpMAPIFormInfo);
	void _OutputTable(ULONG ulDbgLvl, _In_opt_ FILE* fFile, _In_ LPMAPITABLE lpMAPITable);
	void _OutputNotifications(
		ULONG ulDbgLvl,
		_In_opt_ FILE* fFile,
		ULONG cNotify,
		_In_count_(cNotify) LPNOTIFICATION lpNotifications,
		_In_opt_ LPMAPIPROP lpObj);

#define DebugPrintProperties(ulDbgLvl, cProps, lpProps, lpObj) \
	_OutputProperties((ulDbgLvl), nullptr, (cProps), (lpProps), (lpObj), false)
#define DebugPrintRestriction(ulDbgLvl, lpRes, lpObj) _OutputRestriction((ulDbgLvl), nullptr, (lpRes), (lpObj))
#define DebugPrintFormInfo(ulDbgLvl, lpMAPIFormInfo) _OutputFormInfo((ulDbgLvl), nullptr, (lpMAPIFormInfo))
#define DebugPrintNotifications(ulDbgLvl, cNotify, lpNotifications, lpObj) \
	_OutputNotifications((ulDbgLvl), nullptr, (cNotify), (lpNotifications), (lpObj))
#define DebugPrintSRowSet(ulDbgLvl, lpRowSet, lpObj) _OutputSRowSet((ulDbgLvl), nullptr, (lpRowSet), (lpObj))

#define OutputStreamToFile(fFile, lpStream) _OutputStream(DBGNoDebug, (fFile), (lpStream))
#define OutputTableToFile(fFile, lpMAPITable) _OutputTable(DBGNoDebug, (fFile), (lpMAPITable))
#define OutputSRowToFile(fFile, lpSRow, lpObj) _OutputSRow(DBGNoDebug, fFile, lpSRow, lpObj)
#define OutputPropertiesToFile(fFile, cProps, lpProps, lpObj, bRetry) \
	_OutputProperties(DBGNoDebug, fFile, cProps, lpProps, lpObj, bRetry)
#define OutputPropertyToFile(fFile, lpProp, lpObj, bRetry) _OutputProperty(DBGNoDebug, fFile, lpProp, lpObj, bRetry)
} // namespace output