#pragma once

// Here's an index of error macros and functions for use throughout MFCMAPI
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

// EC_B - Similar to EC_H, but for Boolean functions, using CheckWin32Error to set failure in hRes
// WC_B - Similar to WC_H, but for Boolean functions, using CheckWin32Error to set failure in hRes

// EC_D - Similar to EC_B, but preserves the result in _ret so it can be used later
// WC_D - Similar to WC_B, but preserves the result in _ret so it can be used later

// EC_H_GETPROPS - EC_H tuned for GetProps. Tosses all MAPI_W_ERRORS_RETURNED warnings.
// WC_H_GETPROPS - WC_H tuned for GetProps. Tosses all MAPI_W_ERRORS_RETURNED warnings.

// EC_H_CANCEL - EC_H tuned for functions which may be cancelled. Logs but does not display cancel errors.
//               No WC* macro needed here as WC_H already has the desired behavior.

// EC_D_DIALOG - EC_H tuned for dialog functions that support CommDlgExtendedError

// EC_PROBLEMARRAY - logs and displays dialog for any errors in a LPSPropProblemArray
// EC_MAPIERR - logs and displays dialog for any errors in a LPMAPIERROR
// EC_TNEFERR - logs and displays dialog for any errors in a LPSTnefProblemArray

namespace error
{
	extern std::function<void(const std::wstring& errString)> displayError;

	void LogFunctionCall(
		HRESULT hRes,
		HRESULT hrIgnore,
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

	inline _Check_return_ constexpr HRESULT CheckMe(const HRESULT hRes) noexcept { return hRes; }

	std::wstring ProblemArrayToString(_In_ const SPropProblemArray& problems);
	std::wstring MAPIErrToString(ULONG ulFlags, _In_ const MAPIERROR& err);
	std::wstring TnefProblemArrayToString(_In_ const STnefProblemArray& error);
} // namespace error

// Macros for debug output
#define CHECKHRES(hRes) (error::LogFunctionCall((hRes), NULL, true, false, false, NULL, nullptr, __FILE__, __LINE__))
#define CHECKHRESMSG(hRes, uidErrorMsg) \
	(error::LogFunctionCall((hRes), NULL, true, false, false, (uidErrorMsg), nullptr, __FILE__, __LINE__))
#define WARNHRESMSG(hRes, uidErrorMsg) \
	(error::LogFunctionCall((hRes), NULL, false, false, false, (uidErrorMsg), nullptr, __FILE__, __LINE__))

// Execute a function, log and return the HRESULT
// Will display dialog on error
#define EC_H(fnx) \
	error::CheckMe([&]() -> HRESULT { \
		const auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, true, false, false, NULL, #fnx, __FILE__, __LINE__); \
		return __hRes; \
	}())

// Execute a function, log and swallow the HRESULT
// Will display dialog on error
#define EC_H_S(fnx) \
	[&]() -> void { \
		const auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, true, false, false, NULL, #fnx, __FILE__, __LINE__); \
	}()

// Execute a function, log and return the HRESULT
// Will not display an error dialog
#define WC_H(fnx) \
	error::CheckMe([&]() -> HRESULT { \
		const auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, false, false, false, NULL, #fnx, __FILE__, __LINE__); \
		return __hRes; \
	}())

// Execute a function, log and swallow the HRESULT
// Will not display an error dialog
#define WC_H_S(fnx) \
	[&]() -> void { \
		const auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, false, false, false, NULL, #fnx, __FILE__, __LINE__); \
	}()

// Execute a function, log and return the HRESULT
// Logs a MAPI call trace under DBGMAPIFunctions
// Will display dialog on error
#define EC_MAPI(fnx) \
	error::CheckMe([&]() -> HRESULT { \
		const auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, true, true, false, NULL, #fnx, __FILE__, __LINE__); \
		return __hRes; \
	}())

// Execute a function, log and swallow the HRESULT
// Logs a MAPI call trace under DBGMAPIFunctions
// Will display dialog on error
#define EC_MAPI_S(fnx) \
	[&]() -> void { \
		const auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, true, true, false, NULL, #fnx, __FILE__, __LINE__); \
	}()

// Execute a function, log and return the HRESULT
// Logs a MAPI call trace under DBGMAPIFunctions
// Will not display an error dialog
#define WC_MAPI(fnx) \
	error::CheckMe([&]() -> HRESULT { \
		const auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, false, true, false, NULL, #fnx, __FILE__, __LINE__); \
		return __hRes; \
	}())

// Execute a function, log and swallow the HRESULT
// Logs a MAPI call trace under DBGMAPIFunctions
// Will not display an error dialog
#define WC_MAPI_S(fnx) \
	[&]() -> void { \
		const auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, false, true, false, NULL, #fnx, __FILE__, __LINE__); \
	}()

// Execute a function, log error with uidErrorMessage, and return the HRESULT
// Will display dialog on error
#define EC_H_MSG(uidErrorMsg, fnx) \
	error::CheckMe([&]() -> HRESULT { \
		const auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, true, false, false, uidErrorMsg, #fnx, __FILE__, __LINE__); \
		return __hRes; \
	}())

// Execute a function, log error with uidErrorMessage, and return the HRESULT
// Will not display an error dialog
#define WC_H_MSG(uidErrorMsg, fnx) \
	error::CheckMe([&]() -> HRESULT { \
		const auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, false, false, false, uidErrorMsg, #fnx, __FILE__, __LINE__); \
		return __hRes; \
	}())

// Execute a function, log error with uidErrorMessage, and swallow the HRESULT
// Will not display an error dialog
#define WC_H_MSG_S(uidErrorMsg, fnx) \
	[&]() -> void { \
		const auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, false, false, false, uidErrorMsg, #fnx, __FILE__, __LINE__); \
	}()

// Execute a W32 function which returns ERROR_SUCCESS on success, log error, and return HRESULT_FROM_WIN32
// Will display dialog on error
#define EC_W32(fnx) \
	error::CheckMe([&]() -> HRESULT { \
		const auto __hRes = HRESULT_FROM_WIN32(fnx); \
		error::LogFunctionCall(__hRes, NULL, true, false, false, NULL, #fnx, __FILE__, __LINE__); \
		return __hRes; \
	}())

// Execute a W32 function which returns ERROR_SUCCESS on success, log error, and swallow error
// Will display dialog on error
#define EC_W32_S(fnx) \
	[&]() -> void { \
		const auto __hRes = HRESULT_FROM_WIN32(fnx); \
		error::LogFunctionCall(__hRes, NULL, true, false, false, NULL, #fnx, __FILE__, __LINE__); \
	}()

// Execute a W32 function which returns ERROR_SUCCESS on success, log error, and return HRESULT_FROM_WIN32
// Will not display an error dialog
#define WC_W32(fnx) \
	error::CheckMe([&]() -> HRESULT { \
		const auto __hRes = HRESULT_FROM_WIN32(fnx); \
		error::LogFunctionCall(__hRes, NULL, false, false, false, NULL, #fnx, __FILE__, __LINE__); \
		return __hRes; \
	}())

// Execute a W32 function which returns ERROR_SUCCESS on success, log error, and swallow error
// Will not display an error dialog
#define WC_W32_S(fnx) \
	[&]() -> void { \
		const auto __hRes = HRESULT_FROM_WIN32(fnx); \
		error::LogFunctionCall(__hRes, NULL, false, false, false, NULL, #fnx, __FILE__, __LINE__); \
	}()

// Execute a bool/BOOL function, log error, and return CheckWin32Error(HRESULT)
// Will display dialog on error
#define EC_B(fnx) \
	error::CheckMe( \
		[&]() -> HRESULT { return !(fnx) ? error::CheckWin32Error(true, __FILE__, __LINE__, #fnx) : S_OK; }())

// Execute a bool/BOOL function, log error, and swallow error
// Will display dialog on error
#define EC_B_S(fnx) \
	[&]() -> void { \
		if (!(fnx)) \
		{ \
			static_cast<void>(error::CheckWin32Error(true, __FILE__, __LINE__, #fnx)); \
		} \
	}()

// Execute a bool/BOOL function, log error, and return CheckWin32Error(HRESULT)
// Will not display an error dialog
#define WC_B(fnx) \
	error::CheckMe( \
		[&]() -> HRESULT { return !(fnx) ? error::CheckWin32Error(false, __FILE__, __LINE__, #fnx) : S_OK; }())

// Execute a bool/BOOL function, log error, and swallow error
// Will not display an error dialog
#define WC_B_S(fnx) \
	[&]() -> void { \
		if (!(fnx)) \
		{ \
			static_cast<void>(error::CheckWin32Error(false, __FILE__, __LINE__, #fnx)); \
		} \
	}()

// Execute a function which returns 0 on error, log error, and return result
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

// Execute a function which returns 0 on error, log error, and return CheckWin32Error(HRESULT)
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
// Will display dialog on error
#define EC_H_IGNORE_RET(__ignore, fnx) \
	error::CheckMe([&]() -> HRESULT { \
		const auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, __ignore, true, true, false, NULL, #fnx, __FILE__, __LINE__); \
		return __hRes; \
	}())

// Execute a function, log and return the HRESULT
// MAPI's GetProps call will return MAPI_W_ERRORS_RETURNED if even one prop fails
// This is annoying, so this macro tosses those warnings.
// We have to check each prop before we use it anyway, so we don't lose anything here.
// Using this macro, all we have to check is that we got a props array back
// Will display dialog on error
#define EC_H_IGNORE(__ignore, fnx) \
	error::CheckMe([&]() -> HRESULT { \
		const auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, __ignore, true, true, false, NULL, #fnx, __FILE__, __LINE__); \
		return __hRes == (__ignore) ? S_OK : __hRes; \
	}())
#define EC_H_GETPROPS(fnx) EC_H_IGNORE(MAPI_W_ERRORS_RETURNED, fnx)

// Execute a function, log and swallow the HRESULT
// MAPI's GetProps call will return MAPI_W_ERRORS_RETURNED if even one prop fails
// This is annoying, so this macro tosses those warnings.
// We have to check each prop before we use it anyway, so we don't lose anything here.
// Using this macro, all we have to check is that we got a props array back
// Will display dialog on error
#define EC_H_IGNORE_S(__ignore, fnx) \
	[&]() -> void { \
		const auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, __ignore, true, true, false, NULL, #fnx, __FILE__, __LINE__); \
	}()
#define EC_H_GETPROPS_S(fnx) EC_H_IGNORE_S(MAPI_W_ERRORS_RETURNED, fnx)

// Execute a function, log and return the HRESULT
// MAPI's GetProps call will return MAPI_W_ERRORS_RETURNED if even one prop fails
// This is annoying, so this macro tosses those warnings.
// We have to check each prop before we use it anyway, so we don't lose anything here.
// Using this macro, all we have to check is that we got a props array back
// Will not display an error dialog
#define WC_H_IGNORE(__ignore, fnx) \
	error::CheckMe([&]() -> HRESULT { \
		const auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, __ignore, false, true, false, NULL, #fnx, __FILE__, __LINE__); \
		return __hRes == (__ignore) ? S_OK : __hRes; \
	}())
#define WC_H_GETPROPS(fnx) WC_H_IGNORE(MAPI_W_ERRORS_RETURNED, fnx)

// Execute a function, log and swallow the HRESULT
// MAPI's GetProps call will return MAPI_W_ERRORS_RETURNED if even one prop fails
// This is annoying, so this macro tosses those warnings.
// We have to check each prop before we use it anyway, so we don't lose anything here.
// Using this macro, all we have to check is that we got a props array back
// Will not display an error dialog
#define WC_H_IGNORE_S(__ignore, fnx) \
	[&]() -> void { \
		const auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, __ignore, false, true, false, NULL, #fnx, __FILE__, __LINE__); \
	}()
#define WC_H_GETPROPS_S(fnx) WC_H_IGNORE_S(MAPI_W_ERRORS_RETURNED, fnx)

// Execute a function, log and return the HRESULT
// Logs a MAPI call trace under DBGMAPIFunctions
// Some MAPI functions allow MAPI_E_CANCEL or MAPI_E_USER_CANCEL.
// I don't consider these to be errors.
// Will display dialog on error
#define EC_H_CANCEL(fnx) \
	error::CheckMe([&]() -> HRESULT { \
		const auto __hRes = (fnx); \
		if (__hRes == MAPI_E_USER_CANCEL || __hRes == MAPI_E_CANCEL) \
		{ \
			error::LogFunctionCall(__hRes, NULL, false, true, false, IDS_USERCANCELLED, #fnx, __FILE__, __LINE__); \
			return S_OK; \
		} \
		else \
			error::LogFunctionCall(__hRes, NULL, true, true, false, NULL, #fnx, __FILE__, __LINE__); \
		return __hRes; \
	}())

// Execute a function, log and swallow the HRESULT
// Logs a MAPI call trace under DBGMAPIFunctions
// Some MAPI functions allow MAPI_E_CANCEL or MAPI_E_USER_CANCEL.
// I don't consider these to be errors.
// Will display dialog on error
#define EC_H_CANCEL_S(fnx) \
	[&]() -> void { \
		const auto __hRes = (fnx); \
		if (__hRes == MAPI_E_USER_CANCEL || __hRes == MAPI_E_CANCEL) \
			error::LogFunctionCall(__hRes, NULL, false, true, false, IDS_USERCANCELLED, #fnx, __FILE__, __LINE__); \
		else \
			error::LogFunctionCall(__hRes, NULL, true, true, false, NULL, #fnx, __FILE__, __LINE__); \
	}()

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