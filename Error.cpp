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
		CString szFunctionString;
		szFunctionString.FormatMessage(
			IDS_FUNCTION,
			szFile,
			iLine,
			szFunction);

		_Output(DBGMAPIFunctions, NULL, true, szFunctionString);
		_Output(DBGMAPIFunctions, NULL, false, _T("\n"));
	}

	// Check if we have no work to do
	if (S_OK == hRes || hrIgnore == hRes) return;
#ifdef MRMAPI
	if (!fIsSet(DBGHRes)) return;
#else
	if (!fIsSet(DBGHRes) && !bShowDialog) return;
#endif

	// Get our error message if we have one
	CString szErrorMsg;
	if (bSystemCall)
	{
		LPTSTR szErr = NULL;
		DWORD dw = FormatMessage(
			FORMAT_MESSAGE_ALLOCATE_BUFFER|FORMAT_MESSAGE_FROM_SYSTEM,
			0,
			uidErrorMsg,
			0,
			(LPTSTR)&szErr,
			0,
			0);
		if (dw)
		{
			szErrorMsg = szErr;
			LocalFree(szErr);
		}
	}
	else if (uidErrorMsg) (void) szErrorMsg.LoadString(uidErrorMsg);

	CString szErrString;
	szErrString.FormatMessage(
		FAILED(hRes)?IDS_ERRFORMATSTRING:IDS_WARNFORMATSTRING,
		szErrorMsg,
		ErrorNameFromErrorCode(hRes),
		hRes,
		szFunction,
		szFile,
		iLine);

	_Output(DBGHRes,NULL, true, szErrString);
	_Output(DBGHRes,NULL, false, _T("\n"));

	if (bShowDialog)
	{
#ifndef MRMAPI
		CEditor Err(
			NULL,
			ID_PRODUCTNAME,
			NULL,
			(ULONG) 0,
			CEDITOR_BUTTON_OK);
		Err.SetPromptPostFix(szErrString);
		(void) Err.DisplayDialog();
#endif
	}
} // LogFunctionCall

_Check_return_ HRESULT CheckWin32Error(bool bDisplayDialog, _In_z_ LPCSTR szFile, int iLine, _In_z_ LPCSTR szFunction)
{
	DWORD dwErr = GetLastError();
	HRESULT hRes = HRESULT_FROM_WIN32(dwErr);
	LogFunctionCall(hRes, NULL, bDisplayDialog, false, true, dwErr, szFunction, szFile, iLine);
	return hRes;
} // CheckWin32Error

void __cdecl ErrDialog(_In_z_ LPCSTR szFile, int iLine, UINT uidErrorFmt, ...)
{
	CString szErrorFmt;
	(void) szErrorFmt.LoadString(uidErrorFmt);
	CString szErrorBegin;
	CString szErrorEnd;
	CString szCombo;

	// Build out error message from the variant argument list
	va_list argList = NULL;
	va_start(argList, uidErrorFmt);
	szErrorBegin.FormatV(szErrorFmt,argList);
	va_end(argList);

	szErrorEnd.FormatMessage(IDS_INFILEONLINE,szFile,iLine);

	szCombo = szErrorBegin+szErrorEnd;

	_Output(DBGHRes,NULL, true, szCombo);
	_Output(DBGHRes,NULL, false, _T("\n"));

#ifndef MRMAPI
	CEditor Err(
		NULL,
		ID_PRODUCTNAME,
		NULL,
		(ULONG) 0,
		CEDITOR_BUTTON_OK);
	Err.SetPromptPostFix(szCombo);
	(void) Err.DisplayDialog();
#endif
} // ErrDialog

#define RETURN_ERR_CASE(err) case (err): return(_T(#err))

// Function to convert error codes to their names
_Check_return_ LPWSTR ErrorNameFromErrorCode(ULONG hrErr)
{
	ULONG i = 0;

	for (i = 0;i < g_ulErrorArray;i++)
	{
		if (g_ErrorArray[i].ulErrorName == hrErr) return (LPWSTR) g_ErrorArray[i].lpszName;
	}

	HRESULT hRes = S_OK;
	static WCHAR szErrCode[35];
	EC_H(StringCchPrintfW(szErrCode, _countof(szErrCode), L"0x%08X", hrErr)); // STRING_OK

	return(szErrCode);
} // ErrorNameFromErrorCode

#ifdef _DEBUG
void PrintSkipNote(HRESULT hRes, _In_z_ LPCSTR szFunc)
{
	DebugPrint(DBGHRes,
		L"Skipping %hs because hRes = 0x%8x = %ws.\n", // STRING_OK
		szFunc,
		hRes,
		ErrorNameFromErrorCode(hRes));
} // PrintSkipNote
#endif