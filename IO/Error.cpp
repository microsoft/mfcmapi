#include "StdAfx.h"
#include <IO/Error.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <Interpret/ErrorArray.h>

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
		const auto szFunctionString = strings::formatmessage(
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
	auto szErrorMsg = bSystemCall ? strings::formatmessagesys(uidErrorMsg) : uidErrorMsg ? strings::loadstring(uidErrorMsg) : L"";

	const auto szErrString = strings::formatmessage(
		FAILED(hRes) ? IDS_ERRFORMATSTRING : IDS_WARNFORMATSTRING,
		szErrorMsg.c_str(),
		ErrorNameFromErrorCode(hRes).c_str(),
		hRes,
		szFunction,
		szFile,
		iLine);

	Output(DBGHRes, nullptr, true, strings::StripCarriage(szErrString));
	Output(DBGHRes, nullptr, false, L"\n");

	if (bShowDialog)
	{
#ifndef MRMAPI
		CEditor Err(
			nullptr,
			ID_PRODUCTNAME,
			NULL,
			CEDITOR_BUTTON_OK);
		Err.SetPromptPostFix(szErrString);
		(void)Err.DisplayDialog();
#endif
	}
}

_Check_return_ HRESULT CheckWin32Error(bool bDisplayDialog, _In_z_ LPCSTR szFile, int iLine, _In_z_ LPCSTR szFunction)
{
	const auto dwErr = GetLastError();
	const auto hRes = HRESULT_FROM_WIN32(dwErr);
	LogFunctionCall(hRes, NULL, bDisplayDialog, false, true, dwErr, szFunction, szFile, iLine);
	return hRes;
}

void __cdecl ErrDialog(_In_z_ LPCSTR szFile, int iLine, UINT uidErrorFmt, ...)
{
	auto szErrorFmt = strings::loadstring(uidErrorFmt);

	// Build out error message from the variant argument list
	va_list argList = nullptr;
	va_start(argList, uidErrorFmt);
	const auto szErrorBegin = strings::formatV(szErrorFmt.c_str(), argList);
	va_end(argList);

	const auto szCombo = szErrorBegin + strings::formatmessage(IDS_INFILEONLINE, szFile, iLine);

	Output(DBGHRes, nullptr, true, szCombo);
	Output(DBGHRes, nullptr, false, L"\n");

#ifndef MRMAPI
	CEditor Err(
		nullptr,
		ID_PRODUCTNAME,
		NULL,
		CEDITOR_BUTTON_OK);
	Err.SetPromptPostFix(szCombo);
	(void)Err.DisplayDialog();
#endif
}

#define RETURN_ERR_CASE(err) case (err): return(_T(#err))

// Function to convert error codes to their names
std::wstring ErrorNameFromErrorCode(ULONG hrErr)
{
	for (ULONG i = 0; i < g_ulErrorArray; i++)
	{
		if (g_ErrorArray[i].ulErrorName == hrErr) return const_cast<LPWSTR>(g_ErrorArray[i].lpszName);
	}

	return strings::format(L"0x%08X", hrErr); // STRING_OK
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