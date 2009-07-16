#pragma once
// Error.h : header file
//

// Here's an index of error macros and functions for use throughout MFCMAPI
// EC_H - When wrapping a function, checks that hRes is SUCCEEDED, then makes the call
//        If the call fails, logs and displays a dialog.
//        Prints a skip note if the call is not made.
// WC_H - When wrapping a function, checks that hRes is SUCCEEDED, then makes the call
//        If the call fails, logs. It does not display a dialog.
//        Prints a skip note if the call is not made.
// EC_H_MSG - When wrapping a function, checks that hRes is SUCCEEDED, then makes the call
//        If the call fails, logs and displays a dialog with a given error string.
//        Prints a skip note if the call is not made.
// WC_H_MSG - When wrapping a function, checks that hRes is SUCCEEDED, then makes the call
//        If the call fails, logs and a given error string.
//        Prints a skip note if the call is not made.
// CHECKHRES - checks an hRes and logs and displays a dialog on error
// CHECKHRESMSG - checks an hRes and logs and displays a dialog with a given error string on error
// WARNHRESMSG - checks an hRes and logs a given error string on error

// EC_W32 - does the same as EC_H, wrapping the result of the function call in HRESULT_FROM_WIN32
// WC_W32 - does the same as WC_H, wrapping the result of the function call in HRESULT_FROM_WIN32

// EC_B - Similar to EC_H, but for Boolean functions, using DialogOnWin32Error to set failure in hRes
// WC_B - Similar to WC_H, but for Boolean functions, using DialogOnWin32Error to set failure in hRes

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

BOOL CheckHResFn(HRESULT hRes,LPCSTR szFunction,UINT uidErrorMsg,LPCSTR szFile,int iLine);
BOOL WarnHResFn(HRESULT hRes,LPCSTR szFunction,UINT uidErrorMsg,LPCSTR szFile,int iLine);

void __cdecl ErrDialog(LPCSTR szFile,int iLine, UINT uidErrorFmt,...);

// Function to convert error codes to their names
LPTSTR	ErrorNameFromErrorCode(HRESULT hrErr);

// Macros for debug output
#define CHECKHRES(hRes) (CheckHResFn(hRes,"",NULL,__FILE__,__LINE__))
#define CHECKHRESMSG(hRes,szErrorMsg) (CheckHResFn(hRes,NULL,szErrorMsg,__FILE__,__LINE__))
#define WARNHRESMSG(hRes,szErrorMsg) (WarnHResFn(hRes,NULL,szErrorMsg,__FILE__,__LINE__))

HRESULT WarnOnWin32Error(LPCSTR szFile,int iLine, LPCSTR szFunction);
HRESULT DialogOnWin32Error(LPCSTR szFile,int iLine, LPCSTR szFunction);

// We'll only output this information in debug builds.
#ifdef _DEBUG
void PrintSkipNote(HRESULT hRes,LPCSTR szFunc);
#else
#define PrintSkipNote
#endif

#define EC_H(fnx)	\
if (SUCCEEDED(hRes))	\
{	\
	hRes = (fnx);	\
	CheckHResFn(hRes,#fnx,NULL,__FILE__,__LINE__);	\
}	\
else	\
{	\
	PrintSkipNote(hRes,#fnx);\
}

#define WC_H(fnx)	\
if (SUCCEEDED(hRes))	\
{	\
	hRes = (fnx);	\
	WarnHResFn(hRes,#fnx,NULL,__FILE__,__LINE__);	\
}	\
else	\
{	\
	PrintSkipNote(hRes,#fnx);\
}

#define EC_H_MSG(fnx,uidErrorMsg)	\
if (SUCCEEDED(hRes))	\
{	\
	hRes = (fnx);	\
	CheckHResFn(hRes,#fnx,uidErrorMsg,__FILE__,__LINE__);	\
}	\
else	\
{	\
	PrintSkipNote(hRes,#fnx);\
}

#define WC_H_MSG(fnx,uidErrorMsg)	\
if (SUCCEEDED(hRes))	\
{	\
	hRes = (fnx);	\
	WarnHResFn(hRes,#fnx,uidErrorMsg,__FILE__,__LINE__);	\
}	\
else	\
{	\
	PrintSkipNote(hRes,#fnx);\
}

#define EC_W32(fnx)	\
if (SUCCEEDED(hRes))	\
{	\
	hRes = (fnx);	\
	CheckHResFn(HRESULT_FROM_WIN32(hRes),#fnx,NULL,__FILE__,__LINE__);	\
}	\
else	\
{	\
	PrintSkipNote(hRes,#fnx);\
}

#define WC_W32(fnx)	\
if (SUCCEEDED(hRes))	\
{	\
	hRes = (fnx);	\
	WarnHResFn(HRESULT_FROM_WIN32(hRes),#fnx,NULL,__FILE__,__LINE__);	\
}	\
else	\
{	\
	PrintSkipNote(hRes,#fnx);\
}

#define EC_B(fnx)	\
if (SUCCEEDED(hRes))	\
{	\
	if (!(fnx))	\
	{	\
		hRes = DialogOnWin32Error(__FILE__,__LINE__,#fnx);	\
	}	\
}	\
else	\
{	\
	PrintSkipNote(hRes,#fnx);\
}

#define WC_B(fnx)	\
if (SUCCEEDED(hRes))	\
{	\
	if (!(fnx))	\
	{	\
		hRes = WarnOnWin32Error(__FILE__,__LINE__,#fnx);\
	}	\
}	\
else	\
{	\
	PrintSkipNote(hRes,#fnx);\
}

// Used for functions which return 0 on error
// dwRet will contain the return value - assign to a local if needed for other calls.
#define EC_D(_ret,fnx)	\
	\
if (SUCCEEDED(hRes))	\
{	\
	_ret = (fnx);	\
	if (!_ret)	\
	{	\
		hRes = DialogOnWin32Error(__FILE__,__LINE__,#fnx);	\
	}	\
}	\
else	\
{	\
	PrintSkipNote(hRes,#fnx);\
	_ret = NULL;	\
}

// whatever's passed to _ret will contain the return value
#define WC_D(_ret,fnx)	\
if (SUCCEEDED(hRes))	\
{	\
	_ret = (fnx);	\
	if (!_ret)	\
	{	\
		hRes = WarnOnWin32Error(__FILE__,__LINE__,#fnx);\
	}	\
}	\
else	\
{	\
	PrintSkipNote(hRes,#fnx);\
	_ret = NULL;	\
}

// MAPI's GetProps call will return MAPI_W_ERRORS_RETURNED if even one prop fails
// This is annoying, so this macro tosses those warnings.
// We have to check each prop before we use it anyway, so we don't lose anything here.
// Using this macro, all we have to check is that we got a props array back
#define EC_H_GETPROPS(fnx)	\
	\
if (SUCCEEDED(hRes))	\
{	\
	hRes = (fnx);	\
	if (MAPI_W_ERRORS_RETURNED != hRes) CheckHResFn(hRes,#fnx,NULL,__FILE__,__LINE__);	\
}	\
else	\
{	\
	PrintSkipNote(hRes,#fnx);\
}

#define WC_H_GETPROPS(fnx)	\
	\
if (SUCCEEDED(hRes))	\
{	\
	hRes = (fnx);	\
	if (MAPI_W_ERRORS_RETURNED != hRes) WarnHResFn(hRes,#fnx,NULL,__FILE__,__LINE__);	\
}	\
else	\
{	\
	PrintSkipNote(hRes,#fnx);\
}

// some MAPI functions allow MAPI_E_CANCEL or MAPI_E_USER_CANCEL. I don't consider these to be errors.
#define EC_H_CANCEL(fnx)	\
	\
if (SUCCEEDED(hRes))	\
{	\
	hRes = (fnx);	\
	if (MAPI_E_USER_CANCEL == hRes || MAPI_E_CANCEL == hRes) \
	{ \
		WarnHResFn(hRes,#fnx,IDS_USERCANCELLED,__FILE__,__LINE__); \
		hRes = S_OK; \
	} \
	else CheckHResFn(hRes,#fnx,NULL,__FILE__,__LINE__);	\
}	\
else	\
{	\
	PrintSkipNote(hRes,#fnx);\
}

// Designed to check return values from dialog functions, primarily DoModal
// These functions use CommDlgExtendedError to get error information
#define EC_D_DIALOG(fnx)	\
{	\
	iDlgRet = (fnx);	\
	if (IDCANCEL == iDlgRet)	\
	{	\
		DWORD err = CommDlgExtendedError();	\
		if (err) \
		{ \
			ErrDialog(__FILE__,__LINE__,IDS_EDCOMMONDLG,#fnx,err);	\
			hRes = MAPI_E_CALL_FAILED;	\
		} \
		else hRes = S_OK; \
	}	\
}

#define EC_PROBLEMARRAY(problemarray)	\
{	\
	if (problemarray)	\
	{	\
		CString szProbArray = ProblemArrayToString((problemarray));	\
		ErrDialog(__FILE__,__LINE__,IDS_EDPROBLEMARRAY,(LPCTSTR) szProbArray);	\
		DebugPrint(DBGGeneric,_T("Problem array:\n%s\n"),(LPCTSTR) szProbArray);	\
	}	\
}

#define EC_MAPIERR(__ulflags,__lperr)	\
{	\
	if (__lperr)	\
	{	\
		CString szErr = MAPIErrToString((__ulflags),(__lperr));	\
		ErrDialog(__FILE__,__LINE__,IDS_EDMAPIERROR,(LPCTSTR) szErr);	\
		DebugPrint(DBGGeneric,_T("LPMAPIERROR:\n%s\n"),(LPCTSTR) szErr);	\
	}	\
}

#define EC_TNEFERR(problemarray)	\
{	\
	if (problemarray)	\
	{	\
		CString szProbArray = TnefProblemArrayToString((problemarray));	\
		ErrDialog(__FILE__,__LINE__,IDS_EDTNEFPROBLEMARRAY,(LPCTSTR) szProbArray);	\
		DebugPrint(DBGGeneric,_T("TNEF Problem array:\n%s\n"),(LPCTSTR) szProbArray);	\
	}	\
}