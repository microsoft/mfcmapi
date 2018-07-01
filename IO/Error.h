#pragma once

// Here's an index of error macros and functions for use throughout MFCMAPI
// EC_H - When wrapping a function, checks that hRes is SUCCEEDED, then makes the call
//        If the call fails, logs and displays a dialog.
//        Prints a skip note in debug builds if the call is not made.
// WC_H - When wrapping a function, checks that hRes is SUCCEEDED, then makes the call
//        If the call fails, logs. It does not display a dialog.
//        Prints a skip note in debug builds if the call is not made.
// EC_H_MSG - When wrapping a function, checks that hRes is SUCCEEDED, then makes the call
//        If the call fails, logs and displays a dialog with a given error string.
//        Prints a skip note in debug builds if the call is not made.
// WC_H_MSG - When wrapping a function, checks that hRes is SUCCEEDED, then makes the call
//        If the call fails, logs and a given error string.
//        Prints a skip note in debug builds if the call is not made.
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

	// We'll only output this information in debug builds.
#ifdef _DEBUG
	void PrintSkipNote(HRESULT hRes, _In_z_ LPCSTR szFunc);
#else
	inline void PrintSkipNote(HRESULT, _In_z_ LPCSTR) {}
#endif

	// Flag parsing array - used by GetPropFlags
	struct ERROR_ARRAY_ENTRY
	{
		ULONG ulErrorName;
		LPCWSTR lpszName;
	};
	typedef ERROR_ARRAY_ENTRY* LPERROR_ARRAY_ENTRY;
}

#define CheckHResFn(hRes, hrIgnore, bDisplayDialog, szFunction, uidErrorMsg, szFile, iLine) \
	error::LogFunctionCall(hRes, hrIgnore, bDisplayDialog, false, false, uidErrorMsg, szFunction, szFile, iLine)

#define CheckMAPICall(hRes, hrIgnore, bDisplayDialog, szFunction, uidErrorMsg, szFile, iLine) \
	error::LogFunctionCall(hRes, hrIgnore, bDisplayDialog, true, false, uidErrorMsg, szFunction, szFile, iLine)

// Macros for debug output
#define CHECKHRES(hRes) (CheckHResFn(hRes, NULL, true, "", NULL, __FILE__, __LINE__))
#define CHECKHRESMSG(hRes, uidErrorMsg) (CheckHResFn(hRes, NULL, true, nullptr, uidErrorMsg, __FILE__, __LINE__))
#define WARNHRESMSG(hRes, uidErrorMsg) (CheckHResFn(hRes, NULL, false, nullptr, uidErrorMsg, __FILE__, __LINE__))

// Execute a function, log and return the HRESULT
// Does not modify or reference existing hRes
// Will display dialog on error
#define EC_H2(fnx) \
	[&]() -> HRESULT { \
		auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, true, false, false, NULL, #fnx, __FILE__, __LINE__); \
		return __hRes; \
	}()

// Execute a function, log and swallow the HRESULT
// Does not modify or reference existing hRes
// Will display dialog on error
#define EC_H2S(fnx) \
	[&]() -> void { \
		auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, true, false, false, NULL, #fnx, __FILE__, __LINE__); \
	}()

#define EC_H(fnx) \
	{ \
		if (SUCCEEDED(hRes)) \
		{ \
			hRes = (fnx); \
			CheckHResFn(hRes, NULL, true, #fnx, NULL, __FILE__, __LINE__); \
		} \
		else \
		{ \
			error::PrintSkipNote(hRes, #fnx); \
		} \
	}

// Execute a function, log and return the HRESULT
// Does not modify or reference existing hRes
// Will not display an error dialog
#define WC_H2(fnx) \
	[&]() -> HRESULT { \
		auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, false, false, false, NULL, #fnx, __FILE__, __LINE__); \
		return __hRes; \
	}()

// Execute a function, log and swallow the HRESULT
// Does not modify or reference existing hRes
// Will not display an error dialog
#define WC_H2S(fnx) \
	[&]() -> void { \
		auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, false, false, false, NULL, #fnx, __FILE__, __LINE__); \
	}()

#define WC_H(fnx) \
	{ \
		if (SUCCEEDED(hRes)) \
		{ \
			hRes = (fnx); \
			CheckHResFn(hRes, NULL, false, #fnx, NULL, __FILE__, __LINE__); \
		} \
		else \
		{ \
			error::PrintSkipNote(hRes, #fnx); \
		} \
	}

// Execute a function, log and return the HRESULT
// Logs a MAPI call trace under DBGMAPIFunctions
// Does not modify or reference existing hRes
#define EC_MAPI2(fnx) \
	[&]() -> HRESULT { \
		auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, true, true, false, NULL, #fnx, __FILE__, __LINE__); \
		return __hRes; \
	}()

// Execute a function, log and swallow the HRESULT
// Logs a MAPI call trace under DBGMAPIFunctions
// Does not modify or reference existing hRes
#define EC_MAPI2S(fnx) \
	[&]() -> void { \
		auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, true, true, false, NULL, #fnx, __FILE__, __LINE__); \
	}()

#define EC_MAPI(fnx) \
	{ \
		if (SUCCEEDED(hRes)) \
		{ \
			hRes = (fnx); \
			CheckMAPICall(hRes, NULL, true, #fnx, NULL, __FILE__, __LINE__); \
		} \
		else \
		{ \
			error::PrintSkipNote(hRes, #fnx); \
		} \
	}

// Execute a function, log and return the HRESULT
// Logs a MAPI call trace under DBGMAPIFunctions
// Does not modify or reference existing hRes
// Will not display an error dialog
#define WC_MAPI2(fnx) \
	[&]() -> HRESULT { \
		auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, false, true, false, NULL, #fnx, __FILE__, __LINE__); \
		return __hRes; \
	}()

// Execute a function, log and swallow the HRESULT
// Logs a MAPI call trace under DBGMAPIFunctions
// Does not modify or reference existing hRes
// Will not display an error dialog
#define WC_MAPI2S(fnx) \
	[&]() -> void { \
		auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, false, true, false, NULL, #fnx, __FILE__, __LINE__); \
	}()

#define WC_MAPI(fnx) \
	{ \
		if (SUCCEEDED(hRes)) \
		{ \
			hRes = (fnx); \
			CheckMAPICall(hRes, NULL, false, #fnx, NULL, __FILE__, __LINE__); \
		} \
		else \
		{ \
			error::PrintSkipNote(hRes, #fnx); \
		} \
	}

// Execute a function, log error with uidErrorMessage, and return the HRESULT
// Does not modify or reference existing hRes
#define EC_H_MSG(uidErrorMsg, fnx) \
	[&]() -> HRESULT { \
		auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, true, false, false, uidErrorMsg, #fnx, __FILE__, __LINE__); \
		return __hRes; \
	}()

// Execute a function, log error with uidErrorMessage, and return the HRESULT
// Does not modify or reference existing hRes
// Will not display an error dialog
#define WC_H_MSG(uidErrorMsg, fnx) \
	[&]() -> HRESULT { \
		auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, false, false, false, uidErrorMsg, #fnx, __FILE__, __LINE__); \
		return __hRes; \
	}()

// Execute a function, log error with uidErrorMessage, and swallow the HRESULT
// Does not modify or reference existing hRes
// Will not display an error dialog
#define WC_H_MSGS(uidErrorMsg, fnx) \
	[&]() -> void { \
		auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, NULL, false, false, false, uidErrorMsg, #fnx, __FILE__, __LINE__); \
	}()

// Execute a W32 function which returns ERROR_SUCCESS on success, log error, and return HRESULT_FROM_WIN32
// Does not modify or reference existing hRes
#define EC_W32(fnx) \
	[&]() -> HRESULT { \
		auto __hRes = HRESULT_FROM_WIN32(fnx); \
		error::LogFunctionCall(__hRes, NULL, true, false, false, NULL, #fnx, __FILE__, __LINE__); \
		return __hRes; \
	}()

// Execute a W32 function which returns ERROR_SUCCESS on success, log error, and swallow error
// Does not modify or reference existing hRes
#define EC_W32S(fnx) \
	[&]() -> void { \
		auto __hRes = HRESULT_FROM_WIN32(fnx); \
		error::LogFunctionCall(__hRes, NULL, true, false, false, NULL, #fnx, __FILE__, __LINE__); \
	}()

// Execute a W32 function which returns ERROR_SUCCESS on success, log error, and return HRESULT_FROM_WIN32
// Does not modify or reference existing hRes
// Will not display an error dialog
#define WC_W32(fnx) \
	[&]() -> HRESULT { \
		auto __hRes = HRESULT_FROM_WIN32(fnx); \
		error::LogFunctionCall(__hRes, NULL, false, false, false, NULL, #fnx, __FILE__, __LINE__); \
		return __hRes; \
	}()

// Execute a W32 function which returns ERROR_SUCCESS on success, log error, and swallow error
// Does not modify or reference existing hRes
// Will not display an error dialog
#define WC_W32S(fnx) \
	[&]() -> void { \
		auto __hRes = HRESULT_FROM_WIN32(fnx); \
		error::LogFunctionCall(__hRes, NULL, false, false, false, NULL, #fnx, __FILE__, __LINE__); \
	}()

// Execute a bool/BOOL function, log error, and return CheckWin32Error(HRESULT)
// Does not modify or reference existing hRes
#define EC_B(fnx) [&]() -> HRESULT { return !(fnx) ? error::CheckWin32Error(true, __FILE__, __LINE__, #fnx) : S_OK; }()

// Execute a bool/BOOL function, log error, and swallow error
// Does not modify or reference existing hRes
#define EC_BS(fnx) \
	[&]() -> void { \
		if (!(fnx)) \
		{ \
			error::CheckWin32Error(true, __FILE__, __LINE__, #fnx); \
		} \
	}()

// Execute a bool/BOOL function, log error, and return CheckWin32Error(HRESULT)
// Does not modify or reference existing hRes
// Will not display an error dialog
#define WC_B(fnx) [&]() -> HRESULT { return !(fnx) ? error::CheckWin32Error(false, __FILE__, __LINE__, #fnx) : S_OK; }()

// Execute a bool/BOOL function, log error, and swallow error
// Does not modify or reference existing hRes
// Will not display an error dialog
#define WC_BS(fnx) \
	[&]() -> void { \
		if (!(fnx)) \
		{ \
			error::CheckWin32Error(false, __FILE__, __LINE__, #fnx); \
		} \
	}()

// Execute a function which returns 0 on error, log error, and return result
#define EC_D(_TYPE, fnx) \
	[&]() -> _TYPE { \
		auto __ret = (fnx); \
		if (!__ret) \
		{ \
			error::CheckWin32Error(true, __FILE__, __LINE__, #fnx); \
		} \
		return __ret; \
	}()

// Execute a function which returns 0 on error, log error, and return CheckWin32Error(HRESULT)
// Will not display an error dialog
#define WC_D(_TYPE, fnx) \
	[&]() -> _TYPE { \
		auto __ret = (fnx); \
		if (!__ret) \
		{ \
			error::CheckWin32Error(false, __FILE__, __LINE__, #fnx); \
		} \
		return __ret; \
	}()

// Execute a function which returns 0 on error, log error, and swallow error
// Will not display an error dialog
#define WC_DS(fnx) \
	[&]() -> void { \
		if (!(fnx)) \
		{ \
			error::CheckWin32Error(false, __FILE__, __LINE__, #fnx); \
		} \
	}()

// Execute a function, log and return the HRESULT
// MAPI's GetProps call will return MAPI_W_ERRORS_RETURNED if even one prop fails
// This is annoying, so this macro tosses those warnings.
// We have to check each prop before we use it anyway, so we don't lose anything here.
// Using this macro, all we have to check is that we got a props array back
#define EC_H_GETPROPS(fnx) \
	[&]() -> HRESULT { \
		auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, MAPI_W_ERRORS_RETURNED, true, true, false, NULL, #fnx, __FILE__, __LINE__); \
		return __hRes == MAPI_W_ERRORS_RETURNED ? S_OK : __hRes; \
	}()

// Execute a function, log and swallow the HRESULT
// MAPI's GetProps call will return MAPI_W_ERRORS_RETURNED if even one prop fails
// This is annoying, so this macro tosses those warnings.
// We have to check each prop before we use it anyway, so we don't lose anything here.
// Using this macro, all we have to check is that we got a props array back
#define EC_H_GETPROPS_S(fnx) \
	[&]() -> void { \
		auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, MAPI_W_ERRORS_RETURNED, true, true, false, NULL, #fnx, __FILE__, __LINE__); \
	}()

// Execute a function, log and return the HRESULT
// MAPI's GetProps call will return MAPI_W_ERRORS_RETURNED if even one prop fails
// This is annoying, so this macro tosses those warnings.
// We have to check each prop before we use it anyway, so we don't lose anything here.
// Using this macro, all we have to check is that we got a props array back
// Will not display an error dialog
#define WC_H_GETPROPS(fnx) \
	[&]() -> HRESULT { \
		auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, MAPI_W_ERRORS_RETURNED, false, true, false, NULL, #fnx, __FILE__, __LINE__); \
		return __hRes == MAPI_W_ERRORS_RETURNED ? S_OK : __hRes; \
	}()

// Execute a function, log and swallow the HRESULT
// MAPI's GetProps call will return MAPI_W_ERRORS_RETURNED if even one prop fails
// This is annoying, so this macro tosses those warnings.
// We have to check each prop before we use it anyway, so we don't lose anything here.
// Using this macro, all we have to check is that we got a props array back
// Will not display an error dialog
#define WC_H_GETPROPS_S(fnx) \
	[&]() -> void { \
		auto __hRes = (fnx); \
		error::LogFunctionCall(__hRes, MAPI_W_ERRORS_RETURNED, false, true, false, NULL, #fnx, __FILE__, __LINE__); \
	}()

// Execute a function, log and return the HRESULT
// Logs a MAPI call trace under DBGMAPIFunctions
// Some MAPI functions allow MAPI_E_CANCEL or MAPI_E_USER_CANCEL.
// I don't consider these to be errors.
// Does not modify or reference existing hRes
#define EC_H_CANCEL(fnx) \
	[&]() -> HRESULT { \
		auto __hRes = (fnx); \
		if (MAPI_E_USER_CANCEL == __hRes || MAPI_E_CANCEL == __hRes) \
		{ \
			error::LogFunctionCall(__hRes, NULL, true, true, false, IDS_USERCANCELLED, #fnx, __FILE__, __LINE__); \
			return S_OK; \
		} \
		else \
			error::LogFunctionCall(__hRes, NULL, true, true, false, NULL, #fnx, __FILE__, __LINE__); \
		return __hRes; \
	}()

// Execute a function, log and swallow the HRESULT
// Logs a MAPI call trace under DBGMAPIFunctions
// Some MAPI functions allow MAPI_E_CANCEL or MAPI_E_USER_CANCEL.
// I don't consider these to be errors.
// Does not modify or reference existing hRes
#define EC_H_CANCEL_S(fnx) \
	[&]() -> void { \
		auto __hRes = (fnx); \
		if (MAPI_E_USER_CANCEL == __hRes || MAPI_E_CANCEL == __hRes) \
			error::LogFunctionCall(__hRes, NULL, true, true, false, IDS_USERCANCELLED, #fnx, __FILE__, __LINE__); \
		else \
			error::LogFunctionCall(__hRes, NULL, true, true, false, NULL, #fnx, __FILE__, __LINE__); \
	}()

// Execute a function, log and return the command ID
// Designed to check return values from dialog functions, primarily DoModal
// These functions use CommDlgExtendedError to get error information
#define EC_D_DIALOG(fnx) \
	[&]() -> INT_PTR { \
		auto __iDlgRet = (fnx); \
		if (IDCANCEL == __iDlgRet) \
		{ \
			auto __err = CommDlgExtendedError(); \
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
			std::wstring szProbArray = interpretprop::ProblemArrayToString(*(problemarray)); \
			error::ErrDialog(__FILE__, __LINE__, IDS_EDPROBLEMARRAY, szProbArray.c_str()); \
			output::DebugPrint(DBGGeneric, L"Problem array:\n%ws\n", szProbArray.c_str()); \
		} \
	}

#define WC_PROBLEMARRAY(problemarray) \
	{ \
		if (problemarray) \
		{ \
			std::wstring szProbArray = interpretprop::ProblemArrayToString(*(problemarray)); \
			output::DebugPrint(DBGGeneric, L"Problem array:\n%ws\n", szProbArray.c_str()); \
		} \
	}

#define EC_MAPIERR(__ulflags, __lperr) \
	{ \
		if (__lperr) \
		{ \
			std::wstring szErr = interpretprop::MAPIErrToString((__ulflags), *(__lperr)); \
			error::ErrDialog(__FILE__, __LINE__, IDS_EDMAPIERROR, szErr.c_str()); \
			output::DebugPrint(DBGGeneric, L"LPMAPIERROR:\n%ws\n", szErr.c_str()); \
		} \
	}

#define EC_TNEFERR(problemarray) \
	{ \
		if (problemarray) \
		{ \
			std::wstring szProbArray = interpretprop::TnefProblemArrayToString(*(problemarray)); \
			error::ErrDialog(__FILE__, __LINE__, IDS_EDTNEFPROBLEMARRAY, szProbArray.c_str()); \
			output::DebugPrint(DBGGeneric, L"TNEF Problem array:\n%ws\n", szProbArray.c_str()); \
		} \
	}