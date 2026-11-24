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

	template <typename IgnoreRange>
	static HRESULT
	ReportResultImpl(const HRESULT hRes, const Diagnostic& diagnostic, const CallSite& site, const IgnoreRange& ignored)
	{
		if (fIsSet(output::dbgLevel::MAPIFunctions) && (diagnostic.flags & Flags::TraceMapi) == Flags::TraceMapi)
		{
			const auto szFunctionString = strings::formatmessage(IDS_FUNCTION, site.file, site.line, site.expression);

			output::Output(output::dbgLevel::MAPIFunctions, nullptr, true, szFunctionString);
			output::Output(output::dbgLevel::MAPIFunctions, nullptr, false, L"\n");
		}

		// Check if we have no work to do
		if (hRes == S_OK) return hRes;

		const auto showDialog = (diagnostic.flags & Flags::ShowDialog) == Flags::ShowDialog;
		const auto report = (showDialog && displayError) || fIsSet(output::dbgLevel::HRes);
		const auto accept = (diagnostic.flags & Flags::AcceptIgnored) == Flags::AcceptIgnored;
		if (!report && !accept) return hRes;
		if (std::find(ignored.begin(), ignored.end(), hRes) != ignored.end()) return accept ? S_OK : hRes;
		if (!report) return hRes;

		// Get our error message if we have one
		const auto szErrorMsg = (diagnostic.flags & Flags::SystemMessage) == Flags::SystemMessage
									? strings::formatmessagesys(diagnostic.message)
								: diagnostic.message ? strings::loadstring(diagnostic.message)
													 : L"";

		const auto szErrString = strings::formatmessage(
			FAILED(hRes) ? IDS_ERRFORMATSTRING : IDS_WARNFORMATSTRING,
			szErrorMsg.c_str(),
			ErrorNameFromErrorCode(hRes).c_str(),
			hRes,
			site.expression,
			site.file,
			site.line);

		if (fIsSet(output::dbgLevel::HRes))
		{
			output::Output(output::dbgLevel::HRes, nullptr, true, strings::StripCarriage(szErrString));
			output::Output(output::dbgLevel::HRes, nullptr, false, L"\n");
		}

		if (showDialog && displayError)
		{
			displayError(szErrString);
		}

		return hRes;
	}

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
		ReportResultImpl(
			hRes,
			{(bShowDialog ? Flags::ShowDialog : Flags::None) | (bMAPICall ? Flags::TraceMapi : Flags::None) |
				 (bSystemCall ? Flags::SystemMessage : Flags::None),
			 uidErrorMsg},
			{szFunction, szFile, iLine},
			hrIgnore);
	}

	HRESULT ReportResult(
		const HRESULT hRes,
		const Diagnostic& diagnostic,
		const CallSite& site,
		const std::span<const HRESULT> ignored)
	{
		return ReportResultImpl(hRes, diagnostic, site, ignored);
	}

	HRESULT
	ReportResult(
		const HRESULT hRes,
		const Diagnostic& diagnostic,
		const CallSite& site,
		const std::list<HRESULT>& ignored)
	{
		return ReportResultImpl(hRes, diagnostic, site, ignored);
	}

	void ReportResult_S(
		const HRESULT hRes,
		const Diagnostic& diagnostic,
		const CallSite& site,
		const std::span<const HRESULT> ignored)
	{
		ReportResultImpl(hRes, {diagnostic.flags & ~Flags::AcceptIgnored, diagnostic.message}, site, ignored);
	}

	void ReportResult_S(
		const HRESULT hRes,
		const Diagnostic& diagnostic,
		const CallSite& site,
		const std::list<HRESULT>& ignored)
	{
		ReportResultImpl(hRes, {diagnostic.flags & ~Flags::AcceptIgnored, diagnostic.message}, site, ignored);
	}

	static bool IsCancelled(const HRESULT hRes) noexcept { return hRes == MAPI_E_USER_CANCEL || hRes == MAPI_E_CANCEL; }

	void CheckCancelled_S(const HRESULT hRes, const CallSite& site)
	{
		const auto diagnostic = IsCancelled(hRes) ? Diagnostic{Flags::TraceMapi, IDS_USERCANCELLED}
												  : Diagnostic{Flags::ShowDialog | Flags::TraceMapi};
		ReportResult_S(hRes, diagnostic, site);
	}

	HRESULT CheckCancelled(const HRESULT hRes, const CallSite& site)
	{
		CheckCancelled_S(hRes, site);
		return IsCancelled(hRes) ? S_OK : hRes;
	}

	HRESULT CheckWin32(const DWORD result, const Flags flags, const CallSite& site)
	{
		const auto message = (flags & Flags::SystemMessage) == Flags::SystemMessage ? result : 0;
		return ReportResult(HRESULT_FROM_WIN32(result), {flags, message}, site);
	}

	void LogWin32Failure(const DWORD lastError, const Flags flags, const CallSite& site)
	{
		const auto diagnostic = lastError == ERROR_SUCCESS
									? Diagnostic{flags & ~Flags::SystemMessage, IDS_API_FAILED_NO_ERROR}
									: Diagnostic{flags | Flags::SystemMessage, lastError};
		ReportResult_S(LastErrorToHRESULT(lastError), diagnostic, site);
	}

	_Check_return_ HRESULT
	CheckWin32Error(bool bDisplayDialog, _In_z_ LPCSTR szFile, int iLine, _In_z_ LPCSTR szFunction)
	{
		const auto dwErr = GetLastError();
		const auto hRes = dwErr ? HRESULT_FROM_WIN32(dwErr) : E_FAIL;
		LogFunctionCall(
			hRes,
			{},
			bDisplayDialog,
			false,
			dwErr != ERROR_SUCCESS,
			dwErr ? dwErr : IDS_API_FAILED_NO_ERROR,
			szFunction,
			szFile,
			iLine);
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