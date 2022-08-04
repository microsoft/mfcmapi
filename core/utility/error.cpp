#include <core/stdafx.h>
#include <core/utility/error.h>
#include <core/interpret/errorArray.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/utility/registry.h>
#include <core/interpret/proptags.h>

namespace error
{
	std::function<void(const std::wstring& errString)> displayError;

	void LogFunctionCall(
		HRESULT hRes,
		std::list<HRESULT> hrIgnore,
		bool bShowDialog,
		bool bMAPICall,
		bool bSystemCall,
		UINT uidErrorMsg,
		_In_opt_z_ LPCSTR szFunction,
		_In_z_ LPCSTR szFile,
		int iLine)
	{
		if (fIsSet(output::dbgLevel::MAPIFunctions) && bMAPICall)
		{
			const auto szFunctionString = strings::formatmessage(IDS_FUNCTION, szFile, iLine, szFunction);

			output::Output(output::dbgLevel::MAPIFunctions, nullptr, true, szFunctionString);
			output::Output(output::dbgLevel::MAPIFunctions, nullptr, false, L"\n");
		}

		// Check if we have no work to do
		if (hRes == S_OK || std::contains(hrIgnore, hRes)) return;
		if (!(bShowDialog && displayError) && !fIsSet(output::dbgLevel::HRes)) return;

		// Get our error message if we have one
		const auto szErrorMsg = bSystemCall	  ? strings::formatmessagesys(uidErrorMsg)
								: uidErrorMsg ? strings::loadstring(uidErrorMsg)
											  : L"";

		const auto szErrString = strings::formatmessage(
			FAILED(hRes) ? IDS_ERRFORMATSTRING : IDS_WARNFORMATSTRING,
			szErrorMsg.c_str(),
			ErrorNameFromErrorCode(hRes).c_str(),
			hRes,
			szFunction,
			szFile,
			iLine);

		if (fIsSet(output::dbgLevel::HRes))
		{
			output::Output(output::dbgLevel::HRes, nullptr, true, strings::StripCarriage(szErrString));
			output::Output(output::dbgLevel::HRes, nullptr, false, L"\n");
		}

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
		LogFunctionCall(hRes, {}, bDisplayDialog, false, true, dwErr, szFunction, szFile, iLine);
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

		output::Output(output::dbgLevel::HRes, nullptr, true, szCombo);
		output::Output(output::dbgLevel::HRes, nullptr, false, L"\n");

		if (displayError)
		{
			displayError(szCombo);
		}
	}

	// Function to convert error codes to their names
	std::wstring ErrorNameFromErrorCode(ULONG hrErr)
	{
		for (ULONG i = 0; i < g_ulErrorArray; i++)
		{
			if (g_ErrorArray[i].ulErrorName == hrErr) return g_ErrorArray[i].lpszName;
		}

		return strings::format(L"0x%08X", hrErr); // STRING_OK
	}

	std::wstring ProblemArrayToString(_In_ const SPropProblemArray& problems)
	{
		std::wstring szOut;
		for (ULONG i = 0; i < problems.cProblem; i++)
		{
			szOut += strings::formatmessage(
				IDS_PROBLEMARRAY,
				problems.aProblem[i].ulIndex,
				proptags::TagToString(problems.aProblem[i].ulPropTag, nullptr, false, false).c_str(),
				problems.aProblem[i].scode,
				ErrorNameFromErrorCode(problems.aProblem[i].scode).c_str());
		}

		return szOut;
	}

	std::wstring MAPIErrToString(ULONG ulFlags, _In_ const MAPIERROR& err)
	{
		auto szOut = strings::formatmessage(
			ulFlags & MAPI_UNICODE ? IDS_MAPIERRUNICODE : IDS_MAPIERRANSI,
			err.ulVersion,
			err.lpszError,
			err.lpszComponent,
			err.ulLowLevelError,
			ErrorNameFromErrorCode(err.ulLowLevelError).c_str(),
			err.ulContext);

		return szOut;
	}

	std::wstring TnefProblemArrayToString(_In_ const STnefProblemArray& error)
	{
		std::wstring szOut;
		for (ULONG iError = 0; iError < error.cProblem; iError++)
		{
			szOut += strings::formatmessage(
				IDS_TNEFPROBARRAY,
				error.aProblem[iError].ulComponent,
				error.aProblem[iError].ulAttribute,
				proptags::TagToString(error.aProblem[iError].ulPropTag, nullptr, false, false).c_str(),
				error.aProblem[iError].scode,
				ErrorNameFromErrorCode(error.aProblem[iError].scode).c_str());
		}

		return szOut;
	}

	template <typename T> void CheckExtendedError(HRESULT hRes, T lpObject)
	{
		if (hRes == MAPI_E_EXTENDED_ERROR)
		{
			LPMAPIERROR lpErr = nullptr;
			hRes = WC_MAPI(lpObject->GetLastError(hRes, fMapiUnicode, &lpErr));
			if (lpErr)
			{
				EC_MAPIERR(fMapiUnicode, lpErr);
				MAPIFreeBuffer(lpErr);
			}
		}
		else
			CHECKHRES(hRes);
	}

	template void CheckExtendedError<LPMAPIFORMCONTAINER>(HRESULT hRes, LPMAPIFORMCONTAINER lpObject);
	template void CheckExtendedError<LPMAPIFORMMGR>(HRESULT hRes, LPMAPIFORMMGR lpObject);
} // namespace error