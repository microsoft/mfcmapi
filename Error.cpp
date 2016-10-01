#include "stdafx.h"
#include "Error.h"
#include <Dialogs/Editors/Editor.h>
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
		auto szFunctionString = formatmessage(
			IDS_FUNCTION,
			szFile,
			iLine,
			szFunction);

		Output(DBGMAPIFunctions, nullptr, true, szFunctionString);
		Output(DBGMAPIFunctions, nullptr, false, L"\n");
	}

	// Check if we have no work to do
	if (S_OK == hRes || hrIgnore == hRes) return;
#ifdef MRMAPI
	if (!fIsSet(DBGHRes)) return;
#else
	if (!fIsSet(DBGHRes) && !bShowDialog) return;
#endif

	// Get our error message if we have one
	auto szErrorMsg = bSystemCall ? formatmessagesys(uidErrorMsg) : uidErrorMsg ? loadstring(uidErrorMsg) : L"";

	auto szErrString = formatmessage(
		FAILED(hRes) ? IDS_ERRFORMATSTRING : IDS_WARNFORMATSTRING,
		szErrorMsg.c_str(),
		ErrorNameFromErrorCode(hRes).c_str(),
		hRes,
		szFunction,
		szFile,
		iLine);

	Output(DBGHRes, nullptr, true, StripCarriage(szErrString));
	Output(DBGHRes, nullptr, false, L"\n");

	if (bShowDialog)
	{
#ifndef MRMAPI
		CEditor Err(
			nullptr,
			ID_PRODUCTNAME,
			NULL,
			static_cast<ULONG>(0),
			CEDITOR_BUTTON_OK);
		Err.SetPromptPostFix(szErrString);
		(void)Err.DisplayDialog();
#endif
	}
}

_Check_return_ HRESULT CheckWin32Error(bool bDisplayDialog, _In_z_ LPCSTR szFile, int iLine, _In_z_ LPCSTR szFunction)
{
	auto dwErr = GetLastError();
	auto hRes = HRESULT_FROM_WIN32(dwErr);
	LogFunctionCall(hRes, NULL, bDisplayDialog, false, true, dwErr, szFunction, szFile, iLine);
	return hRes;
}

void __cdecl ErrDialog(_In_z_ LPCSTR szFile, int iLine, UINT uidErrorFmt, ...)
{
	auto szErrorFmt = loadstring(uidErrorFmt);

	// Build out error message from the variant argument list
	va_list argList = nullptr;
	va_start(argList, uidErrorFmt);
	auto szErrorBegin = formatV(szErrorFmt.c_str(), argList);
	va_end(argList);

	auto szCombo = szErrorBegin + formatmessage(IDS_INFILEONLINE, szFile, iLine);

	Output(DBGHRes, nullptr, true, szCombo);
	Output(DBGHRes, nullptr, false, L"\n");

#ifndef MRMAPI
	CEditor Err(
		nullptr,
		ID_PRODUCTNAME,
		NULL,
		static_cast<ULONG>(0),
		CEDITOR_BUTTON_OK);
	Err.SetPromptPostFix(szCombo);
	(void)Err.DisplayDialog();
#endif
} // ErrDialog

#define RETURN_ERR_CASE(err) case (err): return(_T(#err))

// Function to convert error codes to their names
wstring ErrorNameFromErrorCode(ULONG hrErr)
{
	for (ULONG i = 0; i < g_ulErrorArray; i++)
	{
		if (g_ErrorArray[i].ulErrorName == hrErr) return const_cast<LPWSTR>(g_ErrorArray[i].lpszName);
	}

	return format(L"0x%08X", hrErr); // STRING_OK
}

#ifdef _DEBUG
void PrintSkipNote(HRESULT hRes, _In_z_ LPCSTR szFunc)
{
	DebugPrint(DBGHRes,
		L"Skipping %hs because hRes = 0x%8x = %ws.\n", // STRING_OK
		szFunc,
		hRes,
		ErrorNameFromErrorCode(hRes).c_str());
}
#endif