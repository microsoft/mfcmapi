#include "stdafx.h"
#include "Error.h"
#include "Editor.h"
#include "ErrorArray.h"

void LogFunctionCall(
	HRESULT hRes,
	HRESULT hrIgnore,
	bool bShowDialog,
	bool bMAPICall,
	bool bSystemCall,
	UINT uidErrorMsg,
	_In_opt_z_ LPCSTR szFunction,
	_In_z_ LPCSTR szFile,
	int iLine)
{
	if (fIsSet(DBGMAPIFunctions) && bMAPICall)
	{
		wstring szFunctionString = formatmessage(
			IDS_FUNCTION,
			szFile,
			iLine,
			szFunction);

		Output(DBGMAPIFunctions, NULL, true, szFunctionString);
		Output(DBGMAPIFunctions, NULL, false, L"\n");
	}

	// Check if we have no work to do
	if (S_OK == hRes || hrIgnore == hRes) return;
#ifdef MRMAPI
	if (!fIsSet(DBGHRes)) return;
#else
	if (!fIsSet(DBGHRes) && !bShowDialog) return;
#endif

	// Get our error message if we have one
	wstring szErrorMsg = bSystemCall ? formatmessagesys(uidErrorMsg) : uidErrorMsg ? loadstring(uidErrorMsg) : L"";

	wstring szErrString = formatmessage(
		FAILED(hRes) ? IDS_ERRFORMATSTRING : IDS_WARNFORMATSTRING,
		szErrorMsg.c_str(),
		ErrorNameFromErrorCode(hRes),
		hRes,
		szFunction,
		szFile,
		iLine);

	Output(DBGHRes, NULL, true, StripCarriage(szErrString));
	Output(DBGHRes, NULL, false, L"\n");

	if (bShowDialog)
	{
#ifndef MRMAPI
		CEditor Err(
			NULL,
			ID_PRODUCTNAME,
			NULL,
			(ULONG)0,
			CEDITOR_BUTTON_OK);
		LPTSTR lpszErr = wstringToLPTSTR(szErrString);
		Err.SetPromptPostFix(lpszErr);
		delete[] lpszErr;
		(void)Err.DisplayDialog();
#endif
	}
}

_Check_return_ HRESULT CheckWin32Error(bool bDisplayDialog, _In_z_ LPCSTR szFile, int iLine, _In_z_ LPCSTR szFunction)
{
	DWORD dwErr = GetLastError();
	HRESULT hRes = HRESULT_FROM_WIN32(dwErr);
	LogFunctionCall(hRes, NULL, bDisplayDialog, false, true, dwErr, szFunction, szFile, iLine);
	return hRes;
}

void __cdecl ErrDialog(_In_z_ LPCSTR szFile, int iLine, UINT uidErrorFmt, ...)
{
	wstring szErrorFmt = loadstring(uidErrorFmt);

	// Build out error message from the variant argument list
	va_list argList = NULL;
	va_start(argList, uidErrorFmt);
	wstring szErrorBegin = formatV(szErrorFmt, argList);
	va_end(argList);

	wstring szCombo = szErrorBegin + formatmessage(IDS_INFILEONLINE, szFile, iLine);

	Output(DBGHRes, NULL, true, szCombo);
	Output(DBGHRes, NULL, false, L"\n");

#ifndef MRMAPI
	CEditor Err(
		NULL,
		ID_PRODUCTNAME,
		NULL,
		(ULONG)0,
		CEDITOR_BUTTON_OK);
	LPTSTR lpszCombo = wstringToLPTSTR(szCombo);
	Err.SetPromptPostFix(lpszCombo);
	delete[] lpszCombo;
	(void)Err.DisplayDialog();
#endif
} // ErrDialog

#define RETURN_ERR_CASE(err) case (err): return(_T(#err))

// Function to convert error codes to their names
_Check_return_ LPWSTR ErrorNameFromErrorCode(ULONG hrErr)
{
	ULONG i = 0;

	for (i = 0; i < g_ulErrorArray; i++)
	{
		if (g_ErrorArray[i].ulErrorName == hrErr) return (LPWSTR)g_ErrorArray[i].lpszName;
	}

	HRESULT hRes = S_OK;
	static WCHAR szErrCode[35];
	EC_H(StringCchPrintfW(szErrCode, _countof(szErrCode), L"0x%08X", hrErr)); // STRING_OK

	return(szErrCode);
}

#ifdef _DEBUG
void PrintSkipNote(HRESULT hRes, _In_z_ LPCSTR szFunc)
{
	DebugPrint(DBGHRes,
		L"Skipping %hs because hRes = 0x%8x = %ws.\n", // STRING_OK
		szFunc,
		hRes,
		ErrorNameFromErrorCode(hRes));
}
#endif