#include <StdAfx.h>
#include <IO/Error.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <Interpret/ErrorArray.h>

namespace error
{
	std::function<void(const std::wstring& errString)> displayError;

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
			const auto szFunctionString = strings::formatmessage(IDS_FUNCTION, szFile, iLine, szFunction);

			output::Output(DBGMAPIFunctions, nullptr, true, szFunctionString);
			output::Output(DBGMAPIFunctions, nullptr, false, L"\n");
		}

		// Check if we have no work to do
		if (hRes == S_OK || hRes == hrIgnore) return;
		if (!fIsSet(DBGHRes)) return;

		// Get our error message if we have one
		auto szErrorMsg =
			bSystemCall ? strings::formatmessagesys(uidErrorMsg) : uidErrorMsg ? strings::loadstring(uidErrorMsg) : L"";

		const auto szErrString = strings::formatmessage(
			FAILED(hRes) ? IDS_ERRFORMATSTRING : IDS_WARNFORMATSTRING,
			szErrorMsg.c_str(),
			ErrorNameFromErrorCode(hRes).c_str(),
			hRes,
			szFunction,
			szFile,
			iLine);

		output::Output(DBGHRes, nullptr, true, strings::StripCarriage(szErrString));
		output::Output(DBGHRes, nullptr, false, L"\n");

		if (bShowDialog && displayError)
		{
			displayError(szErrString);
		}
	}

	_Check_return_ HRESULT
	CheckWin32Error(bool bDisplayDialog, _In_z_ LPCSTR szFile, int iLine, _In_z_ LPCSTR szFunction)
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

		output::Output(DBGHRes, nullptr, true, szCombo);
		output::Output(DBGHRes, nullptr, false, L"\n");

		if (displayError)
		{
			displayError(szCombo);
		}
	}

#define RETURN_ERR_CASE(err) \
	case (err): \
		return (_T(#err))

	// Function to convert error codes to their names
	std::wstring ErrorNameFromErrorCode(ULONG hrErr)
	{
		for (ULONG i = 0; i < g_ulErrorArray; i++)
		{
			if (g_ErrorArray[i].ulErrorName == hrErr) return g_ErrorArray[i].lpszName;
		}

		return strings::format(L"0x%08X", hrErr); // STRING_OK
	}
} // namespace error