#include "stdafx.h"
#include "Error.h"
#include "Editor.h"
#include "ErrorArray.h"

// This function WILL _Output if it is called
void LogError(
			  HRESULT hRes,
			  UINT uidErrorMsg,
			  _In_z_ LPCSTR szFile,
			  int iLine,
			  _In_opt_z_ LPCSTR szFunction,
			  bool bShowDialog,
			  bool bErrorMsgFromSystem)
{
	// Get our error message if we have one
	CString szErrorMsg;
	if (bErrorMsgFromSystem)
	{
		// We do this ourselves because CString doesn't know how to
		LPTSTR szErr = NULL;
		FormatMessage(
			FORMAT_MESSAGE_ALLOCATE_BUFFER|FORMAT_MESSAGE_FROM_SYSTEM,
			0,
			uidErrorMsg,
			0,
			(LPTSTR)&szErr,
			0,
			0);
		szErrorMsg = szErr;
		LocalFree(szErr);
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
} // LogError

void CheckHResFn(HRESULT hRes, _In_opt_z_ LPCSTR szFunction, UINT uidErrorMsg, _In_z_ LPCSTR szFile, int iLine)
{
	if (S_OK == hRes) return;
	LogError(hRes,uidErrorMsg,szFile,iLine,szFunction,true,false);
} // CheckHResFn

// Warn logs an error but never displays a dialog
// We can log MAPI_W errors along with normal ones
void WarnHResFn(HRESULT hRes, _In_opt_z_ LPCSTR szFunction, UINT uidErrorMsg, _In_z_ LPCSTR szFile, int iLine)
{
	if (fIsSet(DBGHRes) && S_OK != hRes)
	{
		LogError(hRes,uidErrorMsg,szFile,iLine,szFunction,false,false);
	}
} // WarnHResFn

_Check_return_ HRESULT DialogOnWin32Error(_In_z_ LPCSTR szFile, int iLine, _In_z_ LPCSTR szFunction)
{
	DWORD dwErr = GetLastError();
	if (0 == dwErr) return S_OK;

	HRESULT hRes = HRESULT_FROM_WIN32(dwErr);
	if (S_OK == hRes) return S_OK;
	LogError(hRes,dwErr,szFile,iLine,szFunction,true,true);

	return hRes;
} // DialogOnWin32Error

_Check_return_ HRESULT WarnOnWin32Error(_In_z_ LPCSTR szFile, int iLine, _In_z_ LPCSTR szFunction)
{
	DWORD dwErr = GetLastError();
	HRESULT hRes = HRESULT_FROM_WIN32(dwErr);

	if (fIsSet(DBGHRes) && S_OK != hRes)
	{
		LogError(hRes,dwErr,szFile,iLine,szFunction,false,true);
	}

	return hRes;
} // WarnOnWin32Error

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
		_T("Skipping %hs because hRes = 0x%8x = %ws.\n"), // STRING_OK
		szFunc,
		hRes,
		ErrorNameFromErrorCode(hRes));
} // PrintSkipNote
#endif