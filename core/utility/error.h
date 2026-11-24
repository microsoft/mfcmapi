#pragma once
#include <span>
#include <type_traits>
#include <utility>

// Here's an index of error macros and functions for use throughout MFCMAPI
// HRESULT-producing wrappers return HRESULT; test them with FAILED/SUCCEEDED.
// EC_H - Execute a function, log and return the HRESULT
//        Will display dialog on error
// WC_H - Execute a function, log and return the HRESULT
//        Will not display an error dialog
// EC_H_MSG - Execute a function, log error with uidErrorMessage, and return the HRESULT
//            Will display dialog on error
// WC_H_MSG - Execute a function, log error with uidErrorMessage, and return the HRESULT
//            Will not display dialog on error
// EC_MAPI Variant of EC_H that is only used to wrap MAPI api calls.
// WC_MAPI Variant of WC_H that is only used to wrap MAPI api calls.
// CHECKHRES - checks an hRes and logs and displays a dialog on error
// CHECKHRESMSG - checks an hRes and logs and displays a dialog with a given error string on error
// WARNHRESMSG - checks an hRes and logs a given error string on error

// EC_W32 - does the same as EC_H, wrapping the result of the function call in HRESULT_FROM_WIN32
// WC_W32 - does the same as WC_H, wrapping the result of the function call in HRESULT_FROM_WIN32

// EC_B - Wrap a bool/BOOL function; return S_OK on success or a captured last-error failure
// WC_B - Same return contract as EC_B, without an error dialog
//        Both return HRESULT, not BOOL; use FAILED/SUCCEEDED

// EC_BOOL - Wrap a bool/BOOL without GetLastError; return S_OK on success or E_FAIL
// WC_BOOL - Same return contract as EC_BOOL, without an error dialog

// EC_D - Wrap a zero-failure function; log errors and return the native result as _TYPE
// WC_D - Same return contract as EC_D, without an error dialog

// EC_H_GETPROPS - Accept MAPI_W_ERRORS_RETURNED without diagnostics and return S_OK for it
// WC_H_GETPROPS - Same return contract as EC_H_GETPROPS, without an error dialog
//                Callers must still check individual properties for errors

// EC_H_CANCEL - Log cancellation without a dialog and return S_OK; preserve other HRESULTs
//               WC_H logs without a dialog but does not normalize cancellation to S_OK

// EC_D_DIALOG - Log dialog errors using CommDlgExtendedError and return the native command ID

// _S variants discard the result, not the diagnostics

// EC_PROBLEMARRAY - logs and displays dialog for any errors in a LPSPropProblemArray
// EC_MAPIERR - logs and displays dialog for any errors in a LPMAPIERROR
// EC_TNEFERR - logs and displays dialog for any errors in a LPSTnefProblemArray

namespace std
{
	template <class _container, class _Ty> inline bool contains(_container _Cont, const _Ty& __t)
	{
		return std::find(_Cont.begin(), _Cont.end(), __t) != _Cont.end();
	}
}; // namespace std

namespace error
{
	// Defaults: no dialog/trace, resource messages, and unchanged ignored results.
	enum class Flags : unsigned
	{
		None = 0,
		ShowDialog = 1 << 0,
		TraceMapi = 1 << 1,
		SystemMessage = 1 << 2,
		AcceptIgnored = 1 << 3
	};
	DEFINE_ENUM_FLAG_OPERATORS(Flags)

	struct CallSite
	{
		LPCSTR expression;
		LPCSTR file;
		int line;
	};

	struct Diagnostic
	{
		Flags flags = Flags::None;
		UINT message = 0;
	};

	_Check_return_ HRESULT ReportResult(
		HRESULT hRes,
		const Diagnostic& diagnostic,
		const CallSite& site,
		std::span<const HRESULT> ignored = {});
	_Check_return_ HRESULT
	ReportResult(HRESULT hRes, const Diagnostic& diagnostic, const CallSite& site, const std::list<HRESULT>& ignored);
	void ReportResult_S(
		HRESULT hRes,
		const Diagnostic& diagnostic,
		const CallSite& site,
		std::span<const HRESULT> ignored = {});
	void
	ReportResult_S(HRESULT hRes, const Diagnostic& diagnostic, const CallSite& site, const std::list<HRESULT>& ignored);
	_Check_return_ HRESULT CheckCancelled(HRESULT hRes, const CallSite& site);
	void CheckCancelled_S(HRESULT hRes, const CallSite& site);

	// A Boolean false must not be interpreted as HRESULT success.
	template <typename Result, typename Ignored = std::list<HRESULT>>
		requires(std::is_same_v<Result, bool> || std::is_same_v<Result, BOOL>)
	HRESULT ReportResult(Result, const Diagnostic&, const CallSite&, const Ignored& = {}) = delete;
	template <typename Result, typename Ignored = std::list<HRESULT>>
		requires(std::is_same_v<Result, bool> || std::is_same_v<Result, BOOL>)
	void ReportResult_S(Result, const Diagnostic&, const CallSite&, const Ignored& = {}) = delete;
	template <typename Result>
		requires(std::is_same_v<Result, bool> || std::is_same_v<Result, BOOL>)
	HRESULT CheckCancelled(Result, const CallSite&) = delete;
	template <typename Result>
		requires(std::is_same_v<Result, bool> || std::is_same_v<Result, BOOL>)
	void CheckCancelled_S(Result, const CallSite&) = delete;

	_Check_return_ HRESULT CheckWin32(DWORD result, Flags flags, const CallSite& site);

	constexpr HRESULT LastErrorToHRESULT(const DWORD lastError) noexcept
	{
		return lastError == ERROR_SUCCESS ? E_FAIL : HRESULT_FROM_WIN32(lastError);
	}

	void LogWin32Failure(DWORD lastError, Flags flags, const CallSite& site);

	// These helpers require a documented zero-failure/GetLastError API contract.
	template <typename Call>
		requires(std::is_same_v<std::invoke_result_t<Call>, bool> || std::is_same_v<std::invoke_result_t<Call>, BOOL>)
	_Check_return_ HRESULT CheckBool(Call&& call, const Flags flags, const CallSite& site)
	{
		if (std::forward<Call>(call)()) return S_OK;

		const auto lastError = GetLastError();
		LogWin32Failure(lastError, flags, site);
		return LastErrorToHRESULT(lastError);
	}

	template <typename Call>
		requires(std::is_same_v<std::invoke_result_t<Call>, bool> || std::is_same_v<std::invoke_result_t<Call>, BOOL>)
	void CheckBool_S(Call&& call, const Flags flags, const CallSite& site)
	{
		if (!std::forward<Call>(call)())
		{
			const auto lastError = GetLastError();
			LogWin32Failure(lastError, flags, site);
		}
	}

	extern std::function<void(const std::wstring& errString)> displayError;

	void LogFunctionCall(
		HRESULT hRes,
		std::list<HRESULT> hrIgnore,
		bool bShowDialog,
		bool bMAPICall,
		bool bSystemCall,
		UINT uidErrorMsg,
		_In_opt_z_ LPCSTR szFunction,
		_In_z_ LPCSTR szFile,
		int iLine);

	void __cdecl ErrDialog(_In_z_ LPCSTR szFile, int iLine, UINT uidErrorFmt, ...);

	// Function to convert error codes to their names
	std::wstring ErrorNameFromErrorCode(ULONG hrErr);

	_Check_return_ HRESULT
	CheckWin32Error(bool bDisplayDialog, _In_z_ LPCSTR szFile, int iLine, _In_z_ LPCSTR szFunction);

	// Flag parsing array - used by ErrorNameFromErrorCode
	struct ERROR_ARRAY_ENTRY
	{
		ULONG ulErrorName;
		LPCWSTR lpszName;
	};
	typedef ERROR_ARRAY_ENTRY* LPERROR_ARRAY_ENTRY;

	std::wstring ProblemArrayToString(_In_ const SPropProblemArray& problems);
	std::wstring MAPIErrToString(ULONG ulFlags, _In_ const MAPIERROR& err);
	std::wstring TnefProblemArrayToString(_In_ const STnefProblemArray& error);

	template <typename T> void CheckExtendedError(HRESULT hRes, T lpObject);
} // namespace error

// Macros for debug output
#define CHECKHRES(hRes) (error::LogFunctionCall((hRes), {}, true, false, false, NULL, nullptr, __FILE__, __LINE__))
#define CHECKHRESMSG(hRes, uidErrorMsg) \
	(error::LogFunctionCall((hRes), {}, true, false, false, (uidErrorMsg), nullptr, __FILE__, __LINE__))
#define WARNHRESMSG(hRes, uidErrorMsg) \
	(error::LogFunctionCall((hRes), {}, false, false, false, (uidErrorMsg), nullptr, __FILE__, __LINE__))

// Execute a function, log and return the HRESULT
// Will display dialog on error
#define EC_H(fnx) error::ReportResult((fnx), {error::Flags::ShowDialog}, {#fnx, __FILE__, __LINE__})

// Execute a function, log and swallow the HRESULT
// Will display dialog on error
#define EC_H_S(fnx) error::ReportResult_S((fnx), {error::Flags::ShowDialog}, {#fnx, __FILE__, __LINE__})

// Execute a function, log and return the HRESULT
// Will not display an error dialog
#define WC_H(fnx) error::ReportResult((fnx), {}, {#fnx, __FILE__, __LINE__})

// Execute a function, log and swallow the HRESULT
// Will not display an error dialog
#define WC_H_S(fnx) error::ReportResult_S((fnx), {}, {#fnx, __FILE__, __LINE__})

// Execute a function, log and return the HRESULT
// Logs a MAPI call trace under output::dbgLevel::MAPIFunctions
// Will display dialog on error
#define EC_MAPI(fnx) \
	error::ReportResult((fnx), {error::Flags::ShowDialog | error::Flags::TraceMapi}, {#fnx, __FILE__, __LINE__})

// Execute a function, log and swallow the HRESULT
// Logs a MAPI call trace under output::dbgLevel::MAPIFunctions
// Will display dialog on error
#define EC_MAPI_S(fnx) \
	error::ReportResult_S((fnx), {error::Flags::ShowDialog | error::Flags::TraceMapi}, {#fnx, __FILE__, __LINE__})

// Execute a function, log and return the HRESULT
// Logs a MAPI call trace under output::dbgLevel::MAPIFunctions
// Will not display an error dialog
#define WC_MAPI(fnx) error::ReportResult((fnx), {error::Flags::TraceMapi}, {#fnx, __FILE__, __LINE__})

// Execute a function, log and swallow the HRESULT
// Logs a MAPI call trace under output::dbgLevel::MAPIFunctions
// Will not display an error dialog
#define WC_MAPI_S(fnx) error::ReportResult_S((fnx), {error::Flags::TraceMapi}, {#fnx, __FILE__, __LINE__})

// Execute a function, log error with uidErrorMessage, and return the HRESULT
// Will display dialog on error
#define EC_H_MSG(uidErrorMsg, fnx) \
	error::ReportResult((fnx), {error::Flags::ShowDialog, (uidErrorMsg)}, {#fnx, __FILE__, __LINE__})

// Execute a function, log error with uidErrorMessage, and return the HRESULT
// Will not display an error dialog
#define WC_H_MSG(uidErrorMsg, fnx) \
	error::ReportResult((fnx), {error::Flags::None, (uidErrorMsg)}, {#fnx, __FILE__, __LINE__})

// Execute a function, log error with uidErrorMessage, and swallow the HRESULT
// Will not display an error dialog
#define WC_H_MSG_S(uidErrorMsg, fnx) \
	error::ReportResult_S((fnx), {error::Flags::None, (uidErrorMsg)}, {#fnx, __FILE__, __LINE__})

// Execute a W32 function which returns ERROR_SUCCESS on success, log error, and return HRESULT_FROM_WIN32
// Will display dialog on error
#define EC_W32(fnx) error::CheckWin32((fnx), error::Flags::ShowDialog, {#fnx, __FILE__, __LINE__})

// Execute a W32 function which returns ERROR_SUCCESS on success, log error, and swallow error
// Will display dialog on error
#define EC_W32_S(fnx) \
	error::ReportResult_S(HRESULT_FROM_WIN32(fnx), {error::Flags::ShowDialog}, {#fnx, __FILE__, __LINE__})

// Execute a W32 function which returns ERROR_SUCCESS on success, log error, and return HRESULT_FROM_WIN32
// Will not display an error dialog
#define WC_W32(fnx) error::CheckWin32((fnx), error::Flags::None, {#fnx, __FILE__, __LINE__})

// Execute a W32 function which returns ERROR_SUCCESS on success, log error, and swallow error
// Will not display an error dialog
#define WC_W32_S(fnx) error::ReportResult_S(HRESULT_FROM_WIN32(fnx), {}, {#fnx, __FILE__, __LINE__})

// Execute a bool/BOOL function; return S_OK on success or a captured last-error failure
// Returns HRESULT, not BOOL; use FAILED/SUCCEEDED
// Will display dialog on error
#define EC_B(fnx) error::CheckBool([&]() { return (fnx); }, error::Flags::ShowDialog, {#fnx, __FILE__, __LINE__})

// Execute a bool/BOOL function, log error, and swallow error
// Will display dialog on error
#define EC_B_S(fnx) error::CheckBool_S([&]() { return (fnx); }, error::Flags::ShowDialog, {#fnx, __FILE__, __LINE__})

// Execute a bool/BOOL function; return S_OK on success or a captured last-error failure
// Returns HRESULT, not BOOL; use FAILED/SUCCEEDED
// Will not display an error dialog
#define WC_B(fnx) error::CheckBool([&]() { return (fnx); }, error::Flags::None, {#fnx, __FILE__, __LINE__})

// Execute a bool/BOOL function, log error, and swallow error
// Will not display an error dialog
#define WC_B_S(fnx) error::CheckBool_S([&]() { return (fnx); }, error::Flags::None, {#fnx, __FILE__, __LINE__})

// Execute a bool/BOOL function without last-error support; return S_OK on success or E_FAIL
// Will display dialog on error
#define EC_BOOL(fnx) error::ReportResult((fnx) ? S_OK : E_FAIL, {error::Flags::ShowDialog}, {#fnx, __FILE__, __LINE__})

// Execute a bool/BOOL function without last-error support; log E_FAIL on failure and swallow the result
// Will display dialog on error
#define EC_BOOL_S(fnx) \
	error::ReportResult_S((fnx) ? S_OK : E_FAIL, {error::Flags::ShowDialog}, {#fnx, __FILE__, __LINE__})

// Execute a bool/BOOL function without last-error support; return S_OK on success or E_FAIL
// Will not display an error dialog
#define WC_BOOL(fnx) error::ReportResult((fnx) ? S_OK : E_FAIL, {}, {#fnx, __FILE__, __LINE__})

// Execute a bool/BOOL function without last-error support; log E_FAIL on failure and swallow the result
// Will not display an error dialog
#define WC_BOOL_S(fnx) error::ReportResult_S((fnx) ? S_OK : E_FAIL, {}, {#fnx, __FILE__, __LINE__})

// Execute a function which returns 0 on error, log error, and return the native result as _TYPE
// Will display dialog on error
#define EC_D(_TYPE, fnx) \
	[&]() -> _TYPE { \
		const auto __ret = (fnx); \
		if (!__ret) \
		{ \
			static_cast<void>(error::CheckWin32Error(true, __FILE__, __LINE__, #fnx)); \
		} \
		return __ret; \
	}()

// Execute a function which returns 0 on error, log error, and return the native result as _TYPE
// Will not display an error dialog
#define WC_D(_TYPE, fnx) \
	[&]() -> _TYPE { \
		const auto __ret = (fnx); \
		if (!__ret) \
		{ \
			static_cast<void>(error::CheckWin32Error(false, __FILE__, __LINE__, #fnx)); \
		} \
		return __ret; \
	}()

// Execute a function which returns 0 on error, log error, and swallow error
// Will not display an error dialog
#define WC_D_S(fnx) \
	[&]() -> void { \
		if (!(fnx)) \
		{ \
			static_cast<void>(error::CheckWin32Error(false, __FILE__, __LINE__, #fnx)); \
		} \
	}()

// Execute a function, log and return the HRESULT
// Does not log/display on error if it matches __ignore
// Does not suppress __ignore as return value so caller can check it
// Logs a MAPI call trace under output::dbgLevel::MAPIFunctions
// Will display dialog on error
#define EC_H_IGNORE_RET_MAPI(__ignore, fnx) \
	error::ReportResult( \
		(fnx), {error::Flags::ShowDialog | error::Flags::TraceMapi}, {#fnx, __FILE__, __LINE__}, (__ignore))

// Execute a function, log and return the HRESULT
// Does not log/display on error if it matches __ignore
// Does not suppress __ignore as return value so caller can check it
// Will not display an error dialog
#define WC_H_IGNORE_RET(__ignore, fnx) error::ReportResult((fnx), {}, {#fnx, __FILE__, __LINE__}, (__ignore))

// Execute a function, log and return the HRESULT
// Suppress diagnostics for results in __ignore and normalize them to S_OK
// Logs a MAPI call trace under output::dbgLevel::MAPIFunctions
// Will display dialog on error
#define EC_H_IGNORE(__ignore, fnx) \
	error::ReportResult( \
		(fnx), \
		{error::Flags::ShowDialog | error::Flags::TraceMapi | error::Flags::AcceptIgnored}, \
		{#fnx, __FILE__, __LINE__}, \
		(__ignore))
#define EC_H_GETPROPS(fnx) EC_H_IGNORE(std::list<HRESULT>{MAPI_W_ERRORS_RETURNED}, fnx)

// Execute a function, log and swallow the HRESULT
// Suppress diagnostics for results in __ignore
// Logs a MAPI call trace under output::dbgLevel::MAPIFunctions
// Will display dialog on error
#define EC_H_IGNORE_S(__ignore, fnx) \
	error::ReportResult_S( \
		(fnx), {error::Flags::ShowDialog | error::Flags::TraceMapi}, {#fnx, __FILE__, __LINE__}, (__ignore))
#define EC_H_GETPROPS_S(fnx) EC_H_IGNORE_S(std::list<HRESULT>{MAPI_W_ERRORS_RETURNED}, fnx)

// Execute a function, log and return the HRESULT
// Suppress diagnostics for results in __ignore and normalize them to S_OK
// Logs a MAPI call trace under output::dbgLevel::MAPIFunctions
// Will not display an error dialog
#define WC_H_IGNORE(__ignore, fnx) \
	error::ReportResult( \
		(fnx), {error::Flags::TraceMapi | error::Flags::AcceptIgnored}, {#fnx, __FILE__, __LINE__}, (__ignore))
#define WC_H_GETPROPS(fnx) WC_H_IGNORE(std::list<HRESULT>{MAPI_W_ERRORS_RETURNED}, fnx)

// Execute a function, log and swallow the HRESULT
// Suppress diagnostics for results in __ignore
// Logs a MAPI call trace under output::dbgLevel::MAPIFunctions
// Will not display an error dialog
#define WC_H_IGNORE_S(__ignore, fnx) \
	error::ReportResult_S((fnx), {error::Flags::TraceMapi}, {#fnx, __FILE__, __LINE__}, (__ignore))
#define WC_H_GETPROPS_S(fnx) WC_H_IGNORE_S(std::list<HRESULT>{MAPI_W_ERRORS_RETURNED}, fnx)

// Execute a function, log and return the HRESULT
// Logs a MAPI call trace under output::dbgLevel::MAPIFunctions
// Some MAPI functions allow MAPI_E_CANCEL or MAPI_E_USER_CANCEL.
// I don't consider these to be errors.
// Logs a MAPI call trace under output::dbgLevel::MAPIFunctions
// Will display dialog on error
#define EC_H_CANCEL(fnx) error::CheckCancelled((fnx), {#fnx, __FILE__, __LINE__})

// Execute a function, log and swallow the HRESULT
// Logs a MAPI call trace under output::dbgLevel::MAPIFunctions
// Some MAPI functions allow MAPI_E_CANCEL or MAPI_E_USER_CANCEL.
// I don't consider these to be errors.
// Logs a MAPI call trace under output::dbgLevel::MAPIFunctions
// Will display dialog on error
#define EC_H_CANCEL_S(fnx) error::CheckCancelled_S((fnx), {#fnx, __FILE__, __LINE__})

// Execute a function, log and return the command ID
// Designed to check return values from dialog functions, primarily DoModal
// These functions use CommDlgExtendedError to get error information
// Will display dialog on error
#define EC_D_DIALOG(fnx) \
	[&]() -> INT_PTR { \
		const auto __iDlgRet = (fnx); \
		if (IDCANCEL == __iDlgRet) \
		{ \
			const auto __err = CommDlgExtendedError(); \
			if (__err) \
			{ \
				error::ErrDialog(__FILE__, __LINE__, IDS_EDCOMMONDLG, #fnx, __err); \
			} \
		} \
		return __iDlgRet; \
	}()

#define EC_PROBLEMARRAY(problemarray) \
	{ \
		if (problemarray) \
		{ \
			const std::wstring szProbArray = error::ProblemArrayToString(*(problemarray)); \
			error::ErrDialog(__FILE__, __LINE__, IDS_EDPROBLEMARRAY, szProbArray.c_str()); \
			output::DebugPrint(output::dbgLevel::Generic, L"Problem array:\n%ws\n", szProbArray.c_str()); \
		} \
	}

#define WC_PROBLEMARRAY(problemarray) \
	{ \
		if (problemarray) \
		{ \
			const std::wstring szProbArray = error::ProblemArrayToString(*(problemarray)); \
			output::DebugPrint(output::dbgLevel::Generic, L"Problem array:\n%ws\n", szProbArray.c_str()); \
		} \
	}

#define EC_MAPIERR(__ulflags, __lperr) \
	{ \
		if (__lperr) \
		{ \
			const std::wstring szErr = error::MAPIErrToString((__ulflags), *(__lperr)); \
			error::ErrDialog(__FILE__, __LINE__, IDS_EDMAPIERROR, szErr.c_str()); \
			output::DebugPrint(output::dbgLevel::Generic, L"LPMAPIERROR:\n%ws\n", szErr.c_str()); \
		} \
	}

#define EC_TNEFERR(problemarray) \
	{ \
		if (problemarray) \
		{ \
			const std::wstring szProbArray = error::TnefProblemArrayToString(*(problemarray)); \
			error::ErrDialog(__FILE__, __LINE__, IDS_EDTNEFPROBLEMARRAY, szProbArray.c_str()); \
			output::DebugPrint(output::dbgLevel::Generic, L"TNEF Problem array:\n%ws\n", szProbArray.c_str()); \
		} \
	}