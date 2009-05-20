/******************************************************************
*                                                                 *
*  strsafe.h -- This module defines safer C library string        *
*               routine replacements. These are meant to make C   *
*               a bit more safe in reference to security and      *
*               robustness                                        *
*                                                                 *
*  Copyright (c) Microsoft Corp.  All rights reserved.            *
*                                                                 *
******************************************************************/
// OFFICEDEV: VC8 says "Need to include strsafe.h after tchar.h"
#include <tchar.h>

#ifndef _STRSAFE_H_INCLUDED_
#define _STRSAFE_H_INCLUDED_
#if (_MSC_VER > 1000)
#pragma once
#endif

#if !(defined( MIDL_PASS ) || defined(__midl) || defined(RC_INVOKED))
#ifndef ALLOW_LEGACY_STRSAFE_USE
#error "Please don't use strsafe anymore -- use secure CRT functions instead."
#else
#pragma message("Please don't use strsafe anymore -- use secure CRT functions instead")
#endif // ALLOW_LEGACY_STRSAFE_USE
#endif // MIDL

#include <stdio.h>          // for _vsnprintf, _vsnwprintf, getc, getwc
#include <string.h>         // for memset
#include <stdarg.h>         // for va_start, etc.
#include <stdlib.h>         // for _countof
#include <specstrings.h>    // for __in, etc.

#if !defined(_W64)
#if !defined(__midl) && (defined(_X86_) || defined(_M_IX86)) && (_MSC_VER >= 1300)
#define _W64 __w64
#else
#define _W64
#endif
#endif

#if defined(_M_MRX000) || defined(_M_ALPHA) || defined(_M_PPC) || defined(_M_IA64) || defined(_M_AMD64)
#define ALIGNMENT_MACHINE
#define UNALIGNED __unaligned
#if defined(_WIN64)
#define UNALIGNED64 __unaligned
#else
#define UNALIGNED64
#endif
#else
#undef ALIGNMENT_MACHINE
#define UNALIGNED
#define UNALIGNED64
#endif

// typedefs
#ifdef  _WIN64
typedef unsigned __int64    size_t;
#else
typedef _W64 unsigned int   size_t;
#endif

#ifndef _HRESULT_DEFINED
#define _HRESULT_DEFINED
typedef __success(return >= 0) long HRESULT;
#endif

typedef unsigned long DWORD;


// macros
#define SUCCEEDED(hr)   (((HRESULT)(hr)) >= 0)
#define FAILED(hr)      (((HRESULT)(hr)) < 0)

#define S_OK            ((HRESULT)0L)

#ifndef SORTPP_PASS
// compiletime asserts (failure results in error C2118: negative subscript)
#define C_ASSERT(e) typedef char __C_ASSERT__[(e)?1:-1]
#else
#define C_ASSERT(e)
#endif

#ifdef __cplusplus
#define EXTERN_C    extern "C"
#else
#define EXTERN_C    extern
#endif

// OFFICEDEV: BEGIN
// utilities specifically helpful when dealing with strings and sizes

// count of elements in an array, generating compile-time error if passed a
// pointer instead of a real array
#define cElements(x) _countof(x)

// simplify argument lists in function calls with RgC:
//    StringCchCopy(szDst, cElements(szDst), ...)
// becomes
//    StringCchCopy(RgC(szDst), ...)
#ifndef RgC
#define RgC(ary) (ary), _countof(ary)
#endif

// On by default
#define _CRT_SECURE_FORCE_DEPRECATE_CORE
// OFFICEDEV: END

// use the new secure crt functions if available
#ifndef STRSAFE_USE_SECURE_CRT
#if defined(__GOT_SECURE_LIB__) && (__GOT_SECURE_LIB__ >= 200402L)
#define STRSAFE_USE_SECURE_CRT 0
#else
#define STRSAFE_USE_SECURE_CRT 0
#endif
#endif  // !STRSAFE_USE_SECURE_CRT

#ifdef _M_CEE_PURE
#define STRSAFEAPI      __inline HRESULT __clrcall
#else
#define STRSAFEAPI      __inline HRESULT __stdcall
#endif

#if defined(STRSAFE_LIB_IMPL) || defined(STRSAFE_LIB)
#define STRSAFEWORKERAPI    EXTERN_C HRESULT __stdcall
#else
#define STRSAFEWORKERAPI    static STRSAFEAPI
#endif

#ifdef STRSAFE_LOCALE_FUNCTIONS
#if defined(STRSAFE_LOCALE_LIB_IMPL) || defined(STRSAFE_LIB)
#define STRSAFELOCALEWORKERAPI  EXTERN_C HRESULT __stdcall
#else
#define STRSAFELOCALEWORKERAPI  static STRSAFEAPI
#endif
#endif // STRSAFE_LOCALE_FUNCTIONS

#if defined(STRSAFE_LIB)
#pragma comment(lib, "strsafe.lib")
#endif

// The user can request no "Cb" or no "Cch" fuctions, but not both
#if defined(STRSAFE_NO_CB_FUNCTIONS) && defined(STRSAFE_NO_CCH_FUNCTIONS)
#error cannot specify both STRSAFE_NO_CB_FUNCTIONS and STRSAFE_NO_CCH_FUNCTIONS !!
#endif

// The user may override STRSAFE_MAX_CCH, but it must always be less than INT_MAX
#ifndef STRSAFE_MAX_CCH
#define STRSAFE_MAX_CCH     2147483647  // max buffer size, in characters, that we support (same as INT_MAX)
#endif
C_ASSERT(STRSAFE_MAX_CCH <= 2147483647);
C_ASSERT(STRSAFE_MAX_CCH > 1);

#define STRSAFE_MAX_LENGTH  (STRSAFE_MAX_CCH - 1)   // max buffer length, in characters, that we support


// Flags for controling the Ex functions
//
//      STRSAFE_FILL_BYTE(0xFF)                         0x000000FF  // bottom byte specifies fill pattern
#define STRSAFE_IGNORE_NULLS                            0x00000100  // treat null string pointers as TEXT("") -- don't fault on NULL buffers
#define STRSAFE_FILL_BEHIND_NULL                        0x00000200  // on success, fill in extra space behind the null terminator with fill pattern
#define STRSAFE_FILL_ON_FAILURE                         0x00000400  // on failure, overwrite pszDest with fill pattern and null terminate it
#define STRSAFE_NULL_ON_FAILURE                         0x00000800  // on failure, set *pszDest = TEXT('\0')
#define STRSAFE_NO_TRUNCATION                           0x00001000  // instead of returning a truncated result, copy/append nothing to pszDest and null terminate it

#define STRSAFE_VALID_FLAGS                     (0x000000FF | STRSAFE_IGNORE_NULLS | STRSAFE_FILL_BEHIND_NULL | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE | STRSAFE_NO_TRUNCATION)

// helper macro to set the fill character and specify buffer filling
#define STRSAFE_FILL_BYTE(x)                    ((DWORD)((x & 0x000000FF) | STRSAFE_FILL_BEHIND_NULL))
#define STRSAFE_FAILURE_BYTE(x)                 ((DWORD)((x & 0x000000FF) | STRSAFE_FILL_ON_FAILURE))

#define STRSAFE_GET_FILL_PATTERN(dwFlags)       ((int)(dwFlags & 0x000000FF))


// error return codes
#define STRSAFE_E_INSUFFICIENT_BUFFER           ((HRESULT)0x8007007AL)  // 0x7A = 122L = ERROR_INSUFFICIENT_BUFFER
#define STRSAFE_E_INVALID_PARAMETER             ((HRESULT)0x80070057L)  // 0x57 =  87L = ERROR_INVALID_PARAMETER
#define STRSAFE_E_END_OF_FILE                   ((HRESULT)0x80070026L)  // 0x26 =  38L = ERROR_HANDLE_EOF

//
// These typedefs are used in places where the string is guaranteed to
// be null terminated.
//
typedef __nullterminated char* STRSAFE_LPSTR;
typedef __nullterminated const char* STRSAFE_LPCSTR;
typedef __nullterminated wchar_t* STRSAFE_LPWSTR;
typedef __nullterminated const wchar_t* STRSAFE_LPCWSTR;
typedef __nullterminated const wchar_t UNALIGNED* STRSAFE_LPCUWSTR;


// prototypes for the worker functions

STRSAFEWORKERAPI
StringLengthWorkerA(
    __in STRSAFE_LPCSTR psz,
    __in size_t cchMax,
    __out_opt size_t* pcchLength);

STRSAFEWORKERAPI
StringLengthWorkerW(
    __in STRSAFE_LPCWSTR psz,
    __in size_t cchMax,
    __out_opt size_t* pcchLength);

#ifdef ALIGNMENT_MACHINE
STRSAFEWORKERAPI
UnalignedStringLengthWorkerW(
    __in STRSAFE_LPCUWSTR psz,
    __in size_t cchMax,
    __out_opt size_t* pcchLength);
#endif  // ALIGNMENT_MACHINE

STRSAFEWORKERAPI
StringExValidateSrcA(
    __deref_inout_opt STRSAFE_LPCSTR* ppszSrc,
    __inout_opt size_t* pcchToRead,
    __in size_t cchMax,
    __in DWORD dwFlags);

STRSAFEWORKERAPI
StringExValidateSrcW(
    __deref_inout_opt STRSAFE_LPCWSTR* ppszSrc,
    __inout_opt size_t* pcchToRead,
    __in size_t cchMax,
    __in DWORD dwFlags);

STRSAFEWORKERAPI
StringValidateDestA(
    __in_ecount_opt(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __out_opt size_t* pcchDestLength,
    __in size_t cchMax);

STRSAFEWORKERAPI
StringValidateDestW(
    __in_ecount_opt(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __out_opt size_t* pcchDestLength,
    __in size_t cchMax);

STRSAFEWORKERAPI
StringExValidateDestA(
    __deref_inout_opt STRSAFE_LPSTR* ppszDest,
    __inout size_t* pcchDest,
    __out_opt size_t* pcchDestLength,
    __in size_t cchMax,
    __in DWORD dwFlags);

STRSAFEWORKERAPI
StringExValidateDestW(
    __deref_inout_opt STRSAFE_LPWSTR* ppszDest,
    __inout size_t* pcchDest,
    __out_opt size_t* pcchDestLength,
    __in size_t cchMax,
    __in DWORD dwFlags);

STRSAFEWORKERAPI
StringCopyWorkerA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __out_opt size_t* pcchNewDestLength,
    __in STRSAFE_LPCSTR pszSrc,
    __in size_t cchToCopy);

STRSAFEWORKERAPI
StringCopyWorkerW(
    __out_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __out_opt size_t* pcchNewDestLength,
    __in STRSAFE_LPCWSTR pszSrc,
    __in size_t cchToCopy);

STRSAFEWORKERAPI
StringVPrintfWorkerA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __out_opt size_t* pcchNewDestLength,
    __in __format_string STRSAFE_LPCSTR pszFormat,
    __in va_list argList);

#ifdef STRSAFE_LOCALE_FUNCTIONS
STRSAFELOCALEWORKERAPI
StringVPrintf_lWorkerA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __out_opt size_t* pcchNewDestLength,
    __in __format_string STRSAFE_LPCSTR pszFormat,
    __in _locale_t locale,
    __in va_list argList);
#endif  // STRSAFE_LOCALE_FUNCTIONS

STRSAFEWORKERAPI
StringVPrintfWorkerW(
    __out_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __out_opt size_t* pcchNewDestLength,
    __in __format_string STRSAFE_LPCWSTR pszFormat,
    __in va_list argList);

#ifdef STRSAFE_LOCALE_FUNCTIONS
STRSAFELOCALEWORKERAPI
StringVPrintf_lWorkerW(
    __out_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __out_opt size_t* pcchNewDestLength,
    __in __format_string STRSAFE_LPCWSTR pszFormat,
    __in _locale_t locale,
    __in va_list argList);
#endif  // STRSAFE_LOCALE_FUNCTIONS

#ifndef STRSAFE_LIB_IMPL
// always run these functions inline always since they use stdin

STRSAFEAPI
StringGetsWorkerA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __out_opt size_t* pcchNewDestLength);

STRSAFEAPI
StringGetsWorkerW(
    __out_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __out_opt size_t* pcchNewDestLength);

#endif  // !STRSAFE_LIB_IMPL

STRSAFEWORKERAPI
StringExHandleFillBehindNullA(
    __out_bcount(cbRemaining) STRSAFE_LPSTR pszDestEnd,
    __in size_t cbRemaining,
    __in DWORD dwFlags);

STRSAFEWORKERAPI
StringExHandleFillBehindNullW(
    __out_bcount(cbRemaining) STRSAFE_LPWSTR pszDestEnd,
    __in size_t cbRemaining,
    __in DWORD dwFlags);

STRSAFEWORKERAPI
StringExHandleOtherFlagsA(
    __out_bcount(cbDest) STRSAFE_LPSTR pszDest,
    __in size_t cbDest,
    __in size_t cchOriginalDestLength,
    __deref_out_ecount(*pcchRemaining) STRSAFE_LPSTR* ppszDestEnd,
    __out size_t* pcchRemaining,
    __in DWORD dwFlags);

STRSAFEWORKERAPI
StringExHandleOtherFlagsW(
    __out_bcount(cbDest) STRSAFE_LPWSTR pszDest,
    __in size_t cbDest,
    __in size_t cchOriginalDestLength,
    __deref_out_ecount(*pcchRemaining)STRSAFE_LPWSTR* ppszDestEnd,
    __out size_t* pcchRemaining,
    __in DWORD dwFlags);

#pragma warning(push)
#pragma warning(disable : 4996) // 'function': was declared deprecated
#pragma warning(disable : 4995) // name was marked as #pragma deprecated


#ifndef STRSAFE_LIB_IMPL

#ifndef STRSAFE_NO_CCH_FUNCTIONS
/*++

STDAPI
StringCchCopy(
    __out_ecount(cchDest) LPTSTR  pszDest,
    __in  size_t  cchDest,
    __in  LPCTSTR pszSrc
    );

Routine Description:

    This routine is a safer version of the C built-in function 'strcpy'.
    The size of the destination buffer (in characters) is a parameter and
    this function will not write past the end of this buffer and it will
    ALWAYS null terminate the destination buffer (unless it is zero length).

    This routine is not a replacement for strncpy.  That function will pad the
    destination string with extra null termination characters if the count is
    greater than the length of the source string, and it will fail to null
    terminate the destination string if the source string length is greater
    than or equal to the count. You can not blindly use this instead of strncpy:
    it is common for code to use it to "patch" strings and you would introduce
    errors if the code started null terminating in the middle of the string.

    This function returns a hresult, and not a pointer.  It returns
    S_OK if the string was copied without truncation and null terminated,
    otherwise it will return a failure code. In failure cases as much of
    pszSrc will be copied to pszDest as possible, and pszDest will be null
    terminated.

Arguments:

    pszDest        -   destination string

    cchDest        -   size of destination buffer in characters.
                       length must be = (_tcslen(src) + 1) to hold all of the
                       source including the null terminator

    pszSrc         -   source string which must be null terminated

Notes:
    Behavior is undefined if source and destination strings overlap.

    pszDest and pszSrc should not be NULL. See StringCchCopyEx if you require
    the handling of NULL values.

Return Value:

    S_OK           -   if there was source data and it was all copied and the
                       resultant dest string was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the copy
                       operation failed due to insufficient space. When this
                       error occurs, the destination buffer is modified to
                       contain a truncated version of the ideal result and is
                       null terminated. This is useful for situations where
                       truncation is ok

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function.

--*/
#ifdef UNICODE
#define StringCchCopy  StringCchCopyW
#else
#define StringCchCopy  StringCchCopyA
#endif // !UNICODE

STRSAFEAPI
StringCchCopyA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __in STRSAFE_LPCSTR pszSrc)
{
    HRESULT hr;

    hr = StringValidateDestA(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        hr = StringCopyWorkerA(pszDest,
                               cchDest,
                               NULL,
                               pszSrc,
                               STRSAFE_MAX_LENGTH);
    }

    return hr;
}

STRSAFEAPI
StringCchCopyW(
    __out_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __in STRSAFE_LPCWSTR pszSrc)
{
    HRESULT hr;

    hr = StringValidateDestW(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        hr = StringCopyWorkerW(pszDest,
                               cchDest,
                               NULL,
                               pszSrc,
                               STRSAFE_MAX_LENGTH);
    }

    return hr;
}
#endif  // !STRSAFE_NO_CCH_FUNCTIONS


#ifndef STRSAFE_NO_CB_FUNCTIONS
/*++

STDAPI
StringCbCopy(
    __out_bcount(cbDest) LPTSTR pszDest,
    __in  size_t cbDest,
    __in  LPCTSTR pszSrc
    );

Routine Description:

    This routine is a safer version of the C built-in function 'strcpy'.
    The size of the destination buffer (in bytes) is a parameter and this
    function will not write past the end of this buffer and it will ALWAYS
    null terminate the destination buffer (unless it is zero length).

    This routine is not a replacement for strncpy.  That function will pad the
    destination string with extra null termination characters if the count is
    greater than the length of the source string, and it will fail to null
    terminate the destination string if the source string length is greater
    than or equal to the count. You can not blindly use this instead of strncpy:
    it is common for code to use it to "patch" strings and you would introduce
    errors if the code started null terminating in the middle of the string.

    This function returns a hresult, and not a pointer.  It returns
    S_OK if the string was copied without truncation and null terminated,
    otherwise it will return a failure code. In failure cases as much of pszSrc
    will be copied to pszDest as possible, and pszDest will be null terminated.

Arguments:

    pszDest        -   destination string

    cbDest         -   size of destination buffer in bytes.
                       length must be = ((_tcslen(src) + 1) * sizeof(TCHAR)) to
                       hold all of the source including the null terminator

    pszSrc         -   source string which must be null terminated

Notes:
    Behavior is undefined if source and destination strings overlap.

    pszDest and pszSrc should not be NULL.  See StringCbCopyEx if you require
    the handling of NULL values.

Return Value:

    S_OK           -   if there was source data and it was all copied and the
                       resultant dest string was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the copy
                       operation failed due to insufficient space. When this
                       error occurs, the destination buffer is modified to
                       contain a truncated version of the ideal result and is
                       null terminated. This is useful for situations where
                       truncation is ok

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function.

--*/
#ifdef UNICODE
#define StringCbCopy  StringCbCopyW
#else
#define StringCbCopy  StringCbCopyA
#endif // !UNICODE

STRSAFEAPI
StringCbCopyA(
    __out_bcount(cbDest) STRSAFE_LPSTR pszDest,
    __in size_t cbDest,
    __in STRSAFE_LPCSTR pszSrc)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(char);

    hr = StringValidateDestA(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        hr = StringCopyWorkerA(pszDest,
                               cchDest,
                               NULL,
                               pszSrc,
                               STRSAFE_MAX_LENGTH);
    }

    return hr;
}

STRSAFEAPI
StringCbCopyW(
    __out_bcount(cbDest) STRSAFE_LPWSTR pszDest,
    __in size_t cbDest,
    __in STRSAFE_LPCWSTR pszSrc)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(wchar_t);

    hr = StringValidateDestW(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        hr = StringCopyWorkerW(pszDest,
                               cchDest,
                               NULL,
                               pszSrc,
                               STRSAFE_MAX_LENGTH);
    }

    return hr;
}
#endif  // !STRSAFE_NO_CB_FUNCTIONS


#ifndef STRSAFE_NO_CCH_FUNCTIONS
/*++

STDAPI
StringCchCopyEx(
    __out_ecount(cchDest) LPTSTR  pszDest         OPTIONAL,
    __in  size_t  cchDest,
    __in  LPCTSTR pszSrc          OPTIONAL,
    __deref_opt_out_ecount(*pcchRemaining) LPTSTR* ppszDestEnd     OPTIONAL,
    __out_opt size_t* pcchRemaining   OPTIONAL,
    __in  DWORD   dwFlags
    );

Routine Description:

    This routine is a safer version of the C built-in function 'strcpy' with
    some additional parameters.  In addition to functionality provided by
    StringCchCopy, this routine also returns a pointer to the end of the
    destination string and the number of characters left in the destination string
    including the null terminator. The flags parameter allows additional controls.

Arguments:

    pszDest         -   destination string

    cchDest         -   size of destination buffer in characters.
                        length must be = (_tcslen(pszSrc) + 1) to hold all of
                        the source including the null terminator

    pszSrc          -   source string which must be null terminated

    ppszDestEnd     -   if ppszDestEnd is non-null, the function will return a
                        pointer to the end of the destination string.  If the
                        function copied any data, the result will point to the
                        null termination character

    pcchRemaining   -   if pcchRemaining is non-null, the function will return the
                        number of characters left in the destination string,
                        including the null terminator

    dwFlags         -   controls some details of the string copy:

        STRSAFE_FILL_BEHIND_NULL
                    if the function succeeds, the low byte of dwFlags will be
                    used to fill the uninitialize part of destination buffer
                    behind the null terminator

        STRSAFE_IGNORE_NULLS
                    treat NULL string pointers like empty strings (TEXT("")).
                    this flag is useful for emulating functions like lstrcpy

        STRSAFE_FILL_ON_FAILURE
                    if the function fails, the low byte of dwFlags will be
                    used to fill all of the destination buffer, and it will
                    be null terminated. This will overwrite any truncated
                    string returned when the failure is
                    STRSAFE_E_INSUFFICIENT_BUFFER

        STRSAFE_NO_TRUNCATION /
        STRSAFE_NULL_ON_FAILURE
                    if the function fails, the destination buffer will be set
                    to the empty string. This will overwrite any truncated string
                    returned when the failure is STRSAFE_E_INSUFFICIENT_BUFFER.

Notes:
    Behavior is undefined if source and destination strings overlap.

    pszDest and pszSrc should not be NULL unless the STRSAFE_IGNORE_NULLS flag
    is specified.  If STRSAFE_IGNORE_NULLS is passed, both pszDest and pszSrc
    may be NULL.  An error may still be returned even though NULLS are ignored
    due to insufficient space.

Return Value:

    S_OK           -   if there was source data and it was all copied and the
                       resultant dest string was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the copy
                       operation failed due to insufficient space. When this
                       error occurs, the destination buffer is modified to
                       contain a truncated version of the ideal result and is
                       null terminated. This is useful for situations where
                       truncation is ok.

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function

--*/
#ifdef UNICODE
#define StringCchCopyEx  StringCchCopyExW
#else
#define StringCchCopyEx  StringCchCopyExA
#endif // !UNICODE

STRSAFEAPI
StringCchCopyExA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __in STRSAFE_LPCSTR pszSrc,
    __deref_opt_out_ecount(*pcchRemaining) STRSAFE_LPSTR* ppszDestEnd,
    __out_opt size_t* pcchRemaining,
    __in DWORD dwFlags)
{
    HRESULT hr;

    hr = StringExValidateDestA(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcA(&pszSrc, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = '\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually src data to copy
                if (*pszSrc != '\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchCopied = 0;

                hr = StringCopyWorkerA(pszDest,
                                       cchDest,
                                       &cchCopied,
                                       pszSrc,
                                       STRSAFE_MAX_LENGTH);

                pszDestEnd = pszDest + cchCopied;
                cchRemaining = cchDest - cchCopied;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                    cbRemaining = cchRemaining * sizeof(char);

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullA(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = '\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cchDest != 0))
        {
            size_t cbDest;

            // safe to multiply cchDest * sizeof(char) since cchDest < STRSAFE_MAX_CCH and sizeof(char) is 1
            cbDest = cchDest * sizeof(char);

            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsA(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcchRemaining)
            {
                *pcchRemaining = cchRemaining;
            }
        }
    }

    return hr;
}

STRSAFEAPI
StringCchCopyExW(
    __out_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __in STRSAFE_LPCWSTR pszSrc,
    __deref_opt_out_ecount(*pcchRemaining) STRSAFE_LPWSTR* ppszDestEnd,
    __out_opt size_t* pcchRemaining,
    __in DWORD dwFlags)
{
    HRESULT hr;

    hr = StringExValidateDestW(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPWSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcW(&pszSrc, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = L'\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually src data to copy
                if (*pszSrc != L'\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchCopied = 0;

                hr = StringCopyWorkerW(pszDest,
                                       cchDest,
                                       &cchCopied,
                                       pszSrc,
                                       STRSAFE_MAX_LENGTH);

                pszDestEnd = pszDest + cchCopied;
                cchRemaining = cchDest - cchCopied;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                    cbRemaining = cchRemaining * sizeof(wchar_t);

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullW(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = L'\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cchDest != 0))
        {
            size_t cbDest;

            // safe to multiply cchDest * sizeof(wchar_t) since cchDest < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
            cbDest = cchDest * sizeof(wchar_t);

            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsW(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcchRemaining)
            {
                *pcchRemaining = cchRemaining;
            }
        }
    }

    return hr;
}
#endif  // !STRSAFE_NO_CCH_FUNCTIONS


#ifndef STRSAFE_NO_CB_FUNCTIONS
/*++

STDAPI
StringCbCopyEx(
    __out_bcount(cbDest) LPTSTR  pszDest         OPTIONAL,
    __in  size_t  cbDest,
    __in  LPCTSTR pszSrc          OPTIONAL,
    __deref_opt_out_bcount(*pcbRemaining) LPTSTR* ppszDestEnd     OPTIONAL,
    __out_opt size_t* pcbRemaining    OPTIONAL,
    __in  DWORD   dwFlags
    );

Routine Description:

    This routine is a safer version of the C built-in function 'strcpy' with
    some additional parameters.  In addition to functionality provided by
    StringCbCopy, this routine also returns a pointer to the end of the
    destination string and the number of bytes left in the destination string
    including the null terminator. The flags parameter allows additional controls.

Arguments:

    pszDest         -   destination string

    cbDest          -   size of destination buffer in bytes.
                        length must be ((_tcslen(pszSrc) + 1) * sizeof(TCHAR)) to
                        hold all of the source including the null terminator

    pszSrc          -   source string which must be null terminated

    ppszDestEnd     -   if ppszDestEnd is non-null, the function will return a
                        pointer to the end of the destination string.  If the
                        function copied any data, the result will point to the
                        null termination character

    pcbRemaining    -   pcbRemaining is non-null,the function will return the
                        number of bytes left in the destination string,
                        including the null terminator

    dwFlags         -   controls some details of the string copy:

        STRSAFE_FILL_BEHIND_NULL
                    if the function succeeds, the low byte of dwFlags will be
                    used to fill the uninitialize part of destination buffer
                    behind the null terminator

        STRSAFE_IGNORE_NULLS
                    treat NULL string pointers like empty strings (TEXT("")).
                    this flag is useful for emulating functions like lstrcpy

        STRSAFE_FILL_ON_FAILURE
                    if the function fails, the low byte of dwFlags will be
                    used to fill all of the destination buffer, and it will
                    be null terminated. This will overwrite any truncated
                    string returned when the failure is
                    STRSAFE_E_INSUFFICIENT_BUFFER

        STRSAFE_NO_TRUNCATION /
        STRSAFE_NULL_ON_FAILURE
                    if the function fails, the destination buffer will be set
                    to the empty string. This will overwrite any truncated string
                    returned when the failure is STRSAFE_E_INSUFFICIENT_BUFFER.

Notes:
    Behavior is undefined if source and destination strings overlap.

    pszDest and pszSrc should not be NULL unless the STRSAFE_IGNORE_NULLS flag
    is specified.  If STRSAFE_IGNORE_NULLS is passed, both pszDest and pszSrc
    may be NULL.  An error may still be returned even though NULLS are ignored
    due to insufficient space.

Return Value:

    S_OK           -   if there was source data and it was all copied and the
                       resultant dest string was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the copy
                       operation failed due to insufficient space. When this
                       error occurs, the destination buffer is modified to
                       contain a truncated version of the ideal result and is
                       null terminated. This is useful for situations where
                       truncation is ok.

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function

--*/
#ifdef UNICODE
#define StringCbCopyEx  StringCbCopyExW
#else
#define StringCbCopyEx  StringCbCopyExA
#endif // !UNICODE

STRSAFEAPI
StringCbCopyExA(
    __out_bcount(cbDest) STRSAFE_LPSTR pszDest,
    __in size_t cbDest,
    __in STRSAFE_LPCSTR pszSrc,
    __deref_opt_out_bcount(*pcbRemaining) STRSAFE_LPSTR* ppszDestEnd,
    __out_opt size_t* pcbRemaining,
    __in DWORD dwFlags)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(char);

    hr = StringExValidateDestA(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcA(&pszSrc, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = '\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually src data to copy
                if (*pszSrc != '\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchCopied = 0;

                hr = StringCopyWorkerA(pszDest,
                                       cchDest,
                                       &cchCopied,
                                       pszSrc,
                                       STRSAFE_MAX_LENGTH);

                pszDestEnd = pszDest + cchCopied;
                cchRemaining = cchDest - cchCopied;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                    cbRemaining = (cchRemaining * sizeof(char)) + (cbDest % sizeof(char));

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullA(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = '\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cbDest != 0))
        {
            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsA(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcbRemaining)
            {
                // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                *pcbRemaining = (cchRemaining * sizeof(char)) + (cbDest % sizeof(char));
            }
        }
    }

    return hr;
}

STRSAFEAPI
StringCbCopyExW(
    __out_bcount(cbDest) STRSAFE_LPWSTR pszDest,
    __in size_t cbDest,
    __in STRSAFE_LPCWSTR pszSrc,
    __deref_opt_out_bcount(*pcbRemaining) STRSAFE_LPWSTR* ppszDestEnd,
    __out_opt size_t* pcbRemaining,
    __in DWORD dwFlags)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(wchar_t);

    hr = StringExValidateDestW(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPWSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcW(&pszSrc, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = L'\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually src data to copy
                if (*pszSrc != L'\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchCopied = 0;

                hr = StringCopyWorkerW(pszDest,
                                       cchDest,
                                       &cchCopied,
                                       pszSrc,
                                       STRSAFE_MAX_LENGTH);

                pszDestEnd = pszDest + cchCopied;
                cchRemaining = cchDest - cchCopied;

                if (SUCCEEDED(hr) && (dwFlags & STRSAFE_FILL_BEHIND_NULL))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                    cbRemaining = (cchRemaining * sizeof(wchar_t)) + (cbDest % sizeof(wchar_t));

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullW(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = L'\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cbDest != 0))
        {
            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsW(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcbRemaining)
            {
                // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                *pcbRemaining = (cchRemaining * sizeof(wchar_t)) + (cbDest % sizeof(wchar_t));
            }
        }
    }

    return hr;
}
#endif  // !STRSAFE_NO_CB_FUNCTIONS


#ifndef STRSAFE_NO_CCH_FUNCTIONS
/*++

STDAPI
StringCchCopyN(
    __out_ecount(cchDest) LPTSTR  pszDest,
    __in  size_t  cchDest,
    __in  LPCTSTR pszSrc,
    __in  size_t  cchToCopy
    );


Routine Description:

    This routine is a safer version of the C built-in function 'strncpy'.
    The size of the destination buffer (in characters) is a parameter and
    this function will not write past the end of this buffer and it will
    ALWAYS null terminate the destination buffer (unless it is zero length).

    This routine is meant as a replacement for strncpy, but it does behave
    differently. This function will not pad the destination buffer with extra
    null termination characters if cchToCopy is greater than the length of pszSrc.

    This function returns a hresult, and not a pointer.  It returns
    S_OK if the entire string or the first cchToCopy characters were copied
    without truncation and the resultant destination string was null terminated,
    otherwise it will return a failure code. In failure cases as much of pszSrc
    will be copied to pszDest as possible, and pszDest will be null terminated.

Arguments:

    pszDest        -   destination string

    cchDest        -   size of destination buffer in characters.
                       length must be = (_tcslen(src) + 1) to hold all of the
                       source including the null terminator

    pszSrc         -   source string

    cchToCopy      -   maximum number of characters to copy from source string,
                       not including the null terminator.

Notes:
    Behavior is undefined if source and destination strings overlap.

    pszDest and pszSrc should not be NULL. See StringCchCopyNEx if you require
    the handling of NULL values.

Return Value:

    S_OK           -   if there was source data and it was all copied and the
                       resultant dest string was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the copy
                       operation failed due to insufficient space. When this
                       error occurs, the destination buffer is modified to
                       contain a truncated version of the ideal result and is
                       null terminated. This is useful for situations where
                       truncation is ok

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function.

--*/
#ifdef UNICODE
#define StringCchCopyN  StringCchCopyNW
#else
#define StringCchCopyN  StringCchCopyNA
#endif // !UNICODE

STRSAFEAPI
StringCchCopyNA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __in STRSAFE_LPCSTR pszSrc,
    __in size_t cchToCopy)
{
    HRESULT hr;

    hr = StringValidateDestA(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        if (cchToCopy > STRSAFE_MAX_LENGTH)
        {
            hr = STRSAFE_E_INVALID_PARAMETER;

            *pszDest = '\0';
        }
        else
        {
            hr = StringCopyWorkerA(pszDest,
                                   cchDest,
                                   NULL,
                                   pszSrc,
                                   cchToCopy);
        }
    }

    return hr;
}

STRSAFEAPI
StringCchCopyNW(
    __out_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __in STRSAFE_LPCWSTR pszSrc,
    __in size_t cchToCopy)
{
    HRESULT hr;

    hr = StringValidateDestW(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        if (cchToCopy > STRSAFE_MAX_LENGTH)
        {
            hr = STRSAFE_E_INVALID_PARAMETER;

            *pszDest = L'\0';
        }
        else
        {
            hr = StringCopyWorkerW(pszDest,
                                   cchDest,
                                   NULL,
                                   pszSrc,
                                   cchToCopy);
        }
    }

    return hr;
}
#endif  // !STRSAFE_NO_CCH_FUNCTIONS


#ifndef STRSAFE_NO_CB_FUNCTIONS
/*++

STDAPI
StringCbCopyN(
    __out_bcount(cbDest) LPTSTR  pszDest,
    __in  size_t  cbDest,
    __in  LPCTSTR pszSrc,
    __in  size_t  cbToCopy
    );

Routine Description:

    This routine is a safer version of the C built-in function 'strncpy'.
    The size of the destination buffer (in bytes) is a parameter and this
    function will not write past the end of this buffer and it will ALWAYS
    null terminate the destination buffer (unless it is zero length).

    This routine is meant as a replacement for strncpy, but it does behave
    differently. This function will not pad the destination buffer with extra
    null termination characters if cbToCopy is greater than the size of pszSrc.

    This function returns a hresult, and not a pointer.  It returns
    S_OK if the entire string or the first cbToCopy characters were
    copied without truncation and the resultant destination string was null
    terminated, otherwise it will return a failure code. In failure cases as
    much of pszSrc will be copied to pszDest as possible, and pszDest will be
    null terminated.

Arguments:

    pszDest        -   destination string

    cbDest         -   size of destination buffer in bytes.
                       length must be = ((_tcslen(src) + 1) * sizeof(TCHAR)) to
                       hold all of the source including the null terminator

    pszSrc         -   source string

    cbToCopy       -   maximum number of bytes to copy from source string,
                       not including the null terminator.

Notes:
    Behavior is undefined if source and destination strings overlap.

    pszDest and pszSrc should not be NULL.  See StringCbCopyEx if you require
    the handling of NULL values.

Return Value:

    S_OK           -   if there was source data and it was all copied and the
                       resultant dest string was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the copy
                       operation failed due to insufficient space. When this
                       error occurs, the destination buffer is modified to
                       contain a truncated version of the ideal result and is
                       null terminated. This is useful for situations where
                       truncation is ok

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function.

--*/
#ifdef UNICODE
#define StringCbCopyN  StringCbCopyNW
#else
#define StringCbCopyN  StringCbCopyNA
#endif // !UNICODE

STRSAFEAPI
StringCbCopyNA(
    __out_bcount(cbDest) STRSAFE_LPSTR pszDest,
    __in size_t cbDest,
    __in STRSAFE_LPCSTR pszSrc,
    __in size_t cbToCopy)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(char);

    hr = StringValidateDestA(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        size_t cchToCopy = cbToCopy / sizeof(char);

        if (cchToCopy > STRSAFE_MAX_LENGTH)
        {
            hr = STRSAFE_E_INVALID_PARAMETER;

            *pszDest = '\0';
        }
        else
        {
            hr = StringCopyWorkerA(pszDest,
                                   cchDest,
                                   NULL,
                                   pszSrc,
                                   cchToCopy);
        }
    }

    return hr;
}

STRSAFEAPI
StringCbCopyNW(
    __out_bcount(cbDest) STRSAFE_LPWSTR pszDest,
    __in size_t cbDest,
    __in STRSAFE_LPCWSTR pszSrc,
    __in size_t cbToCopy)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(wchar_t);

    hr = StringValidateDestW(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        size_t cchToCopy = cbToCopy / sizeof(wchar_t);

        if (cchToCopy > STRSAFE_MAX_LENGTH)
        {
            hr = STRSAFE_E_INVALID_PARAMETER;

            *pszDest = L'\0';
        }
        else
        {
            hr = StringCopyWorkerW(pszDest,
                                   cchDest,
                                   NULL,
                                   pszSrc,
                                   cchToCopy);
        }
    }

    return hr;
}
#endif  // !STRSAFE_NO_CB_FUNCTIONS


#ifndef STRSAFE_NO_CCH_FUNCTIONS
/*++

STDAPI
StringCchCopyNEx(
    __out_ecount(cchDest) LPTSTR  pszDest         OPTIONAL,
    __in  size_t  cchDest,
    __in  LPCTSTR pszSrc          OPTIONAL,
    __in  size_t  cchToCopy,
    __deref_opt_out_ecount(*pcchRemaining) LPTSTR* ppszDestEnd OPTIONAL,
    __out_opt size_t* pcchRemaining OPTIONAL,
    __in  DWORD   dwFlags
    );

Routine Description:

    This routine is a safer version of the C built-in function 'strncpy' with
    some additional parameters.  In addition to functionality provided by
    StringCchCopyN, this routine also returns a pointer to the end of the
    destination string and the number of characters left in the destination
    string including the null terminator. The flags parameter allows
    additional controls.

    This routine is meant as a replacement for strncpy, but it does behave
    differently. This function will not pad the destination buffer with extra
    null termination characters if cchToCopy is greater than the length of pszSrc.

Arguments:

    pszDest         -   destination string

    cchDest         -   size of destination buffer in characters.
                        length must be = (_tcslen(pszSrc) + 1) to hold all of
                        the source including the null terminator

    pszSrc          -   source string

    cchToCopy       -   maximum number of characters to copy from the source
                        string

    ppszDestEnd     -   if ppszDestEnd is non-null, the function will return a
                        pointer to the end of the destination string.  If the
                        function copied any data, the result will point to the
                        null termination character

    pcchRemaining   -   if pcchRemaining is non-null, the function will return the
                        number of characters left in the destination string,
                        including the null terminator

    dwFlags         -   controls some details of the string copy:

        STRSAFE_FILL_BEHIND_NULL
                    if the function succeeds, the low byte of dwFlags will be
                    used to fill the uninitialize part of destination buffer
                    behind the null terminator

        STRSAFE_IGNORE_NULLS
                    treat NULL string pointers like empty strings (TEXT("")).
                    this flag is useful for emulating functions like lstrcpy

        STRSAFE_FILL_ON_FAILURE
                    if the function fails, the low byte of dwFlags will be
                    used to fill all of the destination buffer, and it will
                    be null terminated. This will overwrite any truncated
                    string returned when the failure is
                    STRSAFE_E_INSUFFICIENT_BUFFER

        STRSAFE_NO_TRUNCATION /
        STRSAFE_NULL_ON_FAILURE
                    if the function fails, the destination buffer will be set
                    to the empty string. This will overwrite any truncated string
                    returned when the failure is STRSAFE_E_INSUFFICIENT_BUFFER.

Notes:
    Behavior is undefined if source and destination strings overlap.

    pszDest and pszSrc should not be NULL unless the STRSAFE_IGNORE_NULLS flag
    is specified. If STRSAFE_IGNORE_NULLS is passed, both pszDest and pszSrc
    may be NULL. An error may still be returned even though NULLS are ignored
    due to insufficient space.

Return Value:

    S_OK           -   if there was source data and it was all copied and the
                       resultant dest string was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the copy
                       operation failed due to insufficient space. When this
                       error occurs, the destination buffer is modified to
                       contain a truncated version of the ideal result and is
                       null terminated. This is useful for situations where
                       truncation is ok.

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function

--*/
#ifdef UNICODE
#define StringCchCopyNEx  StringCchCopyNExW
#else
#define StringCchCopyNEx  StringCchCopyNExA
#endif // !UNICODE

STRSAFEAPI
StringCchCopyNExA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __in STRSAFE_LPCSTR pszSrc,
    __in size_t cchToCopy,
    __deref_opt_out_ecount(*pcchRemaining) STRSAFE_LPSTR* ppszDestEnd,
    __out_opt size_t* pcchRemaining,
    __in DWORD dwFlags)
{
    HRESULT hr;

    hr = StringExValidateDestA(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcA(&pszSrc, &cchToCopy, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = '\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually src data to copy
                if ((cchToCopy != 0) && (*pszSrc != '\0'))
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchCopied = 0;

                hr = StringCopyWorkerA(pszDest,
                                       cchDest,
                                       &cchCopied,
                                       pszSrc,
                                       cchToCopy);

                pszDestEnd = pszDest + cchCopied;
                cchRemaining = cchDest - cchCopied;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                    cbRemaining = cchRemaining * sizeof(char);

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullA(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = '\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cchDest != 0))
        {
            size_t cbDest;

            // safe to multiply cchDest * sizeof(char) since cchDest < STRSAFE_MAX_CCH and sizeof(char) is 1
            cbDest = cchDest * sizeof(char);

            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsA(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcchRemaining)
            {
                *pcchRemaining = cchRemaining;
            }
        }
    }

    return hr;
}

STRSAFEAPI
StringCchCopyNExW(
    __out_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __in STRSAFE_LPCWSTR pszSrc,
    __in size_t cchToCopy,
    __deref_opt_out_ecount(*pcchRemaining) STRSAFE_LPWSTR* ppszDestEnd,
    __out_opt size_t* pcchRemaining,
    __in DWORD dwFlags)
{
    HRESULT hr;

    hr = StringExValidateDestW(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPWSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcW(&pszSrc, &cchToCopy, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = L'\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually src data to copy
                if ((cchToCopy != 0) && (*pszSrc != L'\0'))
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchCopied = 0;

                hr = StringCopyWorkerW(pszDest,
                                       cchDest,
                                       &cchCopied,
                                       pszSrc,
                                       cchToCopy);

                pszDestEnd = pszDest + cchCopied;
                cchRemaining = cchDest - cchCopied;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                    cbRemaining = cchRemaining * sizeof(wchar_t);

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullW(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = L'\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cchDest != 0))
        {
            size_t cbDest;

            // safe to multiply cchDest * sizeof(wchar_t) since cchDest < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
            cbDest = cchDest * sizeof(wchar_t);

            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsW(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcchRemaining)
            {
                *pcchRemaining = cchRemaining;
            }
        }
    }

    return hr;
}
#endif  // !STRSAFE_NO_CCH_FUNCTIONS


#ifndef STRSAFE_NO_CB_FUNCTIONS
/*++

STDAPI
StringCbCopyNEx(
    __out_bcount(cbDest) LPTSTR  pszDest         OPTIONAL,
    __in  size_t  cbDest,
    __in  LPCTSTR pszSrc          OPTIONAL,
    __in  size_t  cbToCopy,
    __deref_opt_out_bcount(*pcbRemaining) LPTSTR* ppszDestEnd     OPTIONAL,
    __out_opt size_t* pcbRemaining    OPTIONAL,
    __in  DWORD   dwFlags
    );

Routine Description:

    This routine is a safer version of the C built-in function 'strncpy' with
    some additional parameters.  In addition to functionality provided by
    StringCbCopyN, this routine also returns a pointer to the end of the
    destination string and the number of bytes left in the destination string
    including the null terminator. The flags parameter allows additional controls.

    This routine is meant as a replacement for strncpy, but it does behave
    differently. This function will not pad the destination buffer with extra
    null termination characters if cbToCopy is greater than the size of pszSrc.

Arguments:

    pszDest         -   destination string

    cbDest          -   size of destination buffer in bytes.
                        length must be ((_tcslen(pszSrc) + 1) * sizeof(TCHAR)) to
                        hold all of the source including the null terminator

    pszSrc          -   source string

    cbToCopy        -   maximum number of bytes to copy from source string

    ppszDestEnd     -   if ppszDestEnd is non-null, the function will return a
                        pointer to the end of the destination string.  If the
                        function copied any data, the result will point to the
                        null termination character

    pcbRemaining    -   pcbRemaining is non-null,the function will return the
                        number of bytes left in the destination string,
                        including the null terminator

    dwFlags         -   controls some details of the string copy:

        STRSAFE_FILL_BEHIND_NULL
                    if the function succeeds, the low byte of dwFlags will be
                    used to fill the uninitialize part of destination buffer
                    behind the null terminator

        STRSAFE_IGNORE_NULLS
                    treat NULL string pointers like empty strings (TEXT("")).
                    this flag is useful for emulating functions like lstrcpy

        STRSAFE_FILL_ON_FAILURE
                    if the function fails, the low byte of dwFlags will be
                    used to fill all of the destination buffer, and it will
                    be null terminated. This will overwrite any truncated
                    string returned when the failure is
                    STRSAFE_E_INSUFFICIENT_BUFFER

        STRSAFE_NO_TRUNCATION /
        STRSAFE_NULL_ON_FAILURE
                    if the function fails, the destination buffer will be set
                    to the empty string. This will overwrite any truncated string
                    returned when the failure is STRSAFE_E_INSUFFICIENT_BUFFER.

Notes:
    Behavior is undefined if source and destination strings overlap.

    pszDest and pszSrc should not be NULL unless the STRSAFE_IGNORE_NULLS flag
    is specified.  If STRSAFE_IGNORE_NULLS is passed, both pszDest and pszSrc
    may be NULL.  An error may still be returned even though NULLS are ignored
    due to insufficient space.

Return Value:

    S_OK           -   if there was source data and it was all copied and the
                       resultant dest string was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the copy
                       operation failed due to insufficient space. When this
                       error occurs, the destination buffer is modified to
                       contain a truncated version of the ideal result and is
                       null terminated. This is useful for situations where
                       truncation is ok.

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function

--*/
#ifdef UNICODE
#define StringCbCopyNEx  StringCbCopyNExW
#else
#define StringCbCopyNEx  StringCbCopyNExA
#endif // !UNICODE

STRSAFEAPI
StringCbCopyNExA(
    __out_bcount(cbDest) STRSAFE_LPSTR pszDest,
    __in size_t cbDest,
    __in STRSAFE_LPCSTR pszSrc,
    __in size_t cbToCopy,
    __deref_opt_out_bcount(*pcbRemaining) STRSAFE_LPSTR* ppszDestEnd,
    __out_opt size_t* pcbRemaining,
    __in DWORD dwFlags)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(char);

    hr = StringExValidateDestA(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;
        size_t cchToCopy = cbToCopy / sizeof(char);

        hr = StringExValidateSrcA(&pszSrc, &cchToCopy, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = '\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually src data to copy
                if ((cchToCopy != 0) && (*pszSrc != '\0'))
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchCopied = 0;

                hr = StringCopyWorkerA(pszDest,
                                       cchDest,
                                       &cchCopied,
                                       pszSrc,
                                       cchToCopy);

                pszDestEnd = pszDest + cchCopied;
                cchRemaining = cchDest - cchCopied;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                    cbRemaining = (cchRemaining * sizeof(char)) + (cbDest % sizeof(char));

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullA(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = '\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cbDest != 0))
        {
            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsA(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcbRemaining)
            {
                // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                *pcbRemaining = (cchRemaining * sizeof(char)) + (cbDest % sizeof(char));
            }
        }
    }

    return hr;
}

STRSAFEAPI
StringCbCopyNExW(
    __out_bcount(cbDest) STRSAFE_LPWSTR pszDest,
    __in size_t cbDest,
    __in STRSAFE_LPCWSTR pszSrc,
    __in size_t cbToCopy,
    __deref_opt_out_bcount(*pcbRemaining) STRSAFE_LPWSTR* ppszDestEnd,
    __out_opt size_t* pcbRemaining,
    __in DWORD dwFlags)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(wchar_t);

    hr = StringExValidateDestW(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPWSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;
        size_t cchToCopy = cbToCopy / sizeof(wchar_t);

        hr = StringExValidateSrcW(&pszSrc, &cchToCopy, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = L'\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually src data to copy
                if ((cchToCopy != 0) && (*pszSrc != L'\0'))
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchCopied = 0;

                hr = StringCopyWorkerW(pszDest,
                                       cchDest,
                                       &cchCopied,
                                       pszSrc,
                                       cchToCopy);

                pszDestEnd = pszDest + cchCopied;
                cchRemaining = cchDest - cchCopied;

                if (SUCCEEDED(hr) && (dwFlags & STRSAFE_FILL_BEHIND_NULL))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                    cbRemaining = (cchRemaining * sizeof(wchar_t)) + (cbDest % sizeof(wchar_t));

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullW(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = L'\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cbDest != 0))
        {
            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsW(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcbRemaining)
            {
                // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                *pcbRemaining = (cchRemaining * sizeof(wchar_t)) + (cbDest % sizeof(wchar_t));
            }
        }
    }

    return hr;
}
#endif  // !STRSAFE_NO_CB_FUNCTIONS


#ifndef STRSAFE_NO_CCH_FUNCTIONS
/*++

STDAPI
StringCchCat(
    __inout_ecount(cchDest) LPTSTR  pszDest,
    __in     size_t  cchDest,
    __in     LPCTSTR pszSrc
    );

Routine Description:

    This routine is a safer version of the C built-in function 'strcat'.
    The size of the destination buffer (in characters) is a parameter and this
    function will not write past the end of this buffer and it will ALWAYS
    null terminate the destination buffer (unless it is zero length).

    This function returns a hresult, and not a pointer.  It returns
    S_OK if the string was concatenated without truncation and null terminated,
    otherwise it will return a failure code. In failure cases as much of pszSrc
    will be appended to pszDest as possible, and pszDest will be null
    terminated.

Arguments:

    pszDest     -  destination string which must be null terminated

    cchDest     -  size of destination buffer in characters.
                   length must be = (_tcslen(pszDest) + _tcslen(pszSrc) + 1)
                   to hold all of the combine string plus the null
                   terminator

    pszSrc      -  source string which must be null terminated

Notes:
    Behavior is undefined if source and destination strings overlap.

    pszDest and pszSrc should not be NULL.  See StringCchCatEx if you require
    the handling of NULL values.

Return Value:

    S_OK           -   if there was source data and it was all concatenated and
                       the resultant dest string was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the operation
                       failed due to insufficient space. When this error occurs,
                       the destination buffer is modified to contain a truncated
                       version of the ideal result and is null terminated. This
                       is useful for situations where truncation is ok.

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function

--*/
#ifdef UNICODE
#define StringCchCat  StringCchCatW
#else
#define StringCchCat  StringCchCatA
#endif // !UNICODE

STRSAFEAPI
StringCchCatA(
    __inout_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __in STRSAFE_LPCSTR pszSrc)
{
    HRESULT hr;
    size_t cchDestLength;

    hr = StringValidateDestA(pszDest, cchDest, &cchDestLength, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        hr = StringCopyWorkerA(pszDest + cchDestLength,
                               cchDest - cchDestLength,
                               NULL,
                               pszSrc,
                               STRSAFE_MAX_CCH);
    }

    return hr;
}

STRSAFEAPI
StringCchCatW(
    __inout_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __in STRSAFE_LPCWSTR pszSrc)
{
    HRESULT hr;
    size_t cchDestLength;

    hr = StringValidateDestW(pszDest, cchDest, &cchDestLength, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        hr = StringCopyWorkerW(pszDest + cchDestLength,
                               cchDest - cchDestLength,
                               NULL,
                               pszSrc,
                               STRSAFE_MAX_CCH);
    }

    return hr;
}
#endif  // !STRSAFE_NO_CCH_FUNCTIONS


#ifndef STRSAFE_NO_CB_FUNCTIONS
/*++

STDAPI
StringCbCat(
    __inout_bcount(cbDest) LPTSTR  pszDest,
    __in     size_t  cbDest,
    __in     LPCTSTR pszSrc
    );

Routine Description:

    This routine is a safer version of the C built-in function 'strcat'.
    The size of the destination buffer (in bytes) is a parameter and this
    function will not write past the end of this buffer and it will ALWAYS
    null terminate the destination buffer (unless it is zero length).

    This function returns a hresult, and not a pointer.  It returns
    S_OK if the string was concatenated without truncation and null terminated,
    otherwise it will return a failure code. In failure cases as much of pszSrc
    will be appended to pszDest as possible, and pszDest will be null
    terminated.

Arguments:

    pszDest     -  destination string which must be null terminated

    cbDest      -  size of destination buffer in bytes.
                   length must be = ((_tcslen(pszDest) + _tcslen(pszSrc) + 1) * sizeof(TCHAR)
                   to hold all of the combine string plus the null
                   terminator

    pszSrc      -  source string which must be null terminated

Notes:
    Behavior is undefined if source and destination strings overlap.

    pszDest and pszSrc should not be NULL.  See StringCbCatEx if you require
    the handling of NULL values.

Return Value:

    S_OK           -   if there was source data and it was all concatenated and
                       the resultant dest string was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the operation
                       failed due to insufficient space. When this error occurs,
                       the destination buffer is modified to contain a truncated
                       version of the ideal result and is null terminated. This
                       is useful for situations where truncation is ok.

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function

--*/
#ifdef UNICODE
#define StringCbCat  StringCbCatW
#else
#define StringCbCat  StringCbCatA
#endif // !UNICODE

STRSAFEAPI
StringCbCatA(
    __inout_bcount(cbDest) STRSAFE_LPSTR pszDest,
    __in size_t cbDest,
    __in STRSAFE_LPCSTR pszSrc)
{
    HRESULT hr;
    size_t cchDestLength;
    size_t cchDest = cbDest / sizeof(char);

    hr = StringValidateDestA(pszDest, cchDest, &cchDestLength, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        hr = StringCopyWorkerA(pszDest + cchDestLength,
                               cchDest - cchDestLength,
                               NULL,
                               pszSrc,
                               STRSAFE_MAX_CCH);
    }

    return hr;
}

STRSAFEAPI
StringCbCatW(
    __inout_bcount(cbDest) STRSAFE_LPWSTR pszDest,
    __in size_t cbDest,
    __in STRSAFE_LPCWSTR pszSrc)
{
    HRESULT hr;
    size_t cchDestLength;
    size_t cchDest = cbDest / sizeof(wchar_t);

    hr = StringValidateDestW(pszDest, cchDest, &cchDestLength, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        hr = StringCopyWorkerW(pszDest + cchDestLength,
                               cchDest - cchDestLength,
                               NULL,
                               pszSrc,
                               STRSAFE_MAX_CCH);
    }

    return hr;
}
#endif  // !STRSAFE_NO_CB_FUNCTIONS


#ifndef STRSAFE_NO_CCH_FUNCTIONS
/*++

STDAPI
StringCchCatEx(
    __inout_ecount(cchDest) LPTSTR  pszDest         OPTIONAL,
    __in     size_t  cchDest,
    __in     LPCTSTR pszSrc          OPTIONAL,
    __deref_opt_out_ecount(*pcchRemaining)    LPTSTR* ppszDestEnd     OPTIONAL,
    __out_opt    size_t* pcchRemaining   OPTIONAL,
    __in     DWORD   dwFlags
    );

Routine Description:

    This routine is a safer version of the C built-in function 'strcat' with
    some additional parameters.  In addition to functionality provided by
    StringCchCat, this routine also returns a pointer to the end of the
    destination string and the number of characters left in the destination string
    including the null terminator. The flags parameter allows additional controls.

Arguments:

    pszDest         -   destination string which must be null terminated

    cchDest         -   size of destination buffer in characters
                        length must be (_tcslen(pszDest) + _tcslen(pszSrc) + 1)
                        to hold all of the combine string plus the null
                        terminator.

    pszSrc          -   source string which must be null terminated

    ppszDestEnd     -   if ppszDestEnd is non-null, the function will return a
                        pointer to the end of the destination string.  If the
                        function appended any data, the result will point to the
                        null termination character

    pcchRemaining   -   if pcchRemaining is non-null, the function will return the
                        number of characters left in the destination string,
                        including the null terminator

    dwFlags         -   controls some details of the string copy:

        STRSAFE_FILL_BEHIND_NULL
                    if the function succeeds, the low byte of dwFlags will be
                    used to fill the uninitialize part of destination buffer
                    behind the null terminator

        STRSAFE_IGNORE_NULLS
                    treat NULL string pointers like empty strings (TEXT("")).
                    this flag is useful for emulating functions like lstrcat

        STRSAFE_FILL_ON_FAILURE
                    if the function fails, the low byte of dwFlags will be
                    used to fill all of the destination buffer, and it will
                    be null terminated. This will overwrite any pre-existing
                    or truncated string

        STRSAFE_NULL_ON_FAILURE
                    if the function fails, the destination buffer will be set
                    to the empty string. This will overwrite any pre-existing or
                    truncated string

        STRSAFE_NO_TRUNCATION
                    if the function returns STRSAFE_E_INSUFFICIENT_BUFFER, pszDest
                    will not contain a truncated string, it will remain unchanged.

Notes:
    Behavior is undefined if source and destination strings overlap.

    pszDest and pszSrc should not be NULL unless the STRSAFE_IGNORE_NULLS flag
    is specified.  If STRSAFE_IGNORE_NULLS is passed, both pszDest and pszSrc
    may be NULL.  An error may still be returned even though NULLS are ignored
    due to insufficient space.

Return Value:

    S_OK           -   if there was source data and it was all concatenated and
                       the resultant dest string was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the operation
                       failed due to insufficient space. When this error
                       occurs, the destination buffer is modified to contain
                       a truncated version of the ideal result and is null
                       terminated. This is useful for situations where
                       truncation is ok.

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function

--*/
#ifdef UNICODE
#define StringCchCatEx  StringCchCatExW
#else
#define StringCchCatEx  StringCchCatExA
#endif // !UNICODE

STRSAFEAPI
StringCchCatExA(
    __inout_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __in STRSAFE_LPCSTR pszSrc,
    __deref_opt_out_ecount(*pcchRemaining) STRSAFE_LPSTR* ppszDestEnd,
    __out_opt size_t* pcchRemaining,
    __in DWORD dwFlags)
{
    HRESULT hr;
    size_t cchDestLength;

    hr = StringExValidateDestA(&pszDest,
                               &cchDest,
                               &cchDestLength,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPSTR pszDestEnd = pszDest + cchDestLength;
        size_t cchRemaining = cchDest - cchDestLength;

        hr = StringExValidateSrcA(&pszSrc, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;
            }
            else if (cchRemaining <= 1)
            {
                // only fail if there was actually src data to append
                if (*pszSrc != '\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchCopied = 0;

                hr = StringCopyWorkerA(pszDestEnd,
                                       cchRemaining,
                                       &cchCopied,
                                       pszSrc,
                                       STRSAFE_MAX_LENGTH);

                pszDestEnd = pszDestEnd + cchCopied;
                cchRemaining = cchRemaining - cchCopied;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                    cbRemaining = cchRemaining * sizeof(char);

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullA(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cchDest != 0))
        {
            size_t cbDest;

            // safe to multiply cchDest * sizeof(char) since cchDest < STRSAFE_MAX_CCH and sizeof(char) is 1
            cbDest = cchDest * sizeof(char);

            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsA(pszDest,
                                      cbDest,
                                      cchDestLength,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcchRemaining)
            {
                *pcchRemaining = cchRemaining;
            }
        }
    }

    return hr;
}

STRSAFEAPI
StringCchCatExW(
    __inout_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __in STRSAFE_LPCWSTR pszSrc,
    __deref_opt_out_ecount(*pcchRemaining) STRSAFE_LPWSTR* ppszDestEnd,
    __out_opt size_t* pcchRemaining,
    __in DWORD dwFlags)
{
    HRESULT hr;
    size_t cchDestLength;

    hr = StringExValidateDestW(&pszDest,
                               &cchDest,
                               &cchDestLength,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPWSTR pszDestEnd = pszDest + cchDestLength;
        size_t cchRemaining = cchDest - cchDestLength;

        hr = StringExValidateSrcW(&pszSrc, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;
            }
            else if (cchRemaining <= 1)
            {
                // only fail if there was actually src data to append
                if (*pszSrc != L'\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchCopied = 0;

                hr = StringCopyWorkerW(pszDestEnd,
                                       cchRemaining,
                                       &cchCopied,
                                       pszSrc,
                                       STRSAFE_MAX_LENGTH);

                pszDestEnd = pszDestEnd + cchCopied;
                cchRemaining = cchRemaining - cchCopied;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                    cbRemaining = cchRemaining * sizeof(wchar_t);

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullW(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cchDest != 0))
        {
            size_t cbDest;

            // safe to multiply cchDest * sizeof(char) since cchDest < STRSAFE_MAX_CCH and sizeof(char) is 1
            cbDest = cchDest * sizeof(wchar_t);

            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsW(pszDest,
                                      cbDest,
                                      cchDestLength,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcchRemaining)
            {
                *pcchRemaining = cchRemaining;
            }
        }
    }

    return hr;
}
#endif  // !STRSAFE_NO_CCH_FUNCTIONS


#ifndef STRSAFE_NO_CB_FUNCTIONS
/*++

STDAPI
StringCbCatEx(
    __inout_bcount(cbDest) LPTSTR  pszDest         OPTIONAL,
    __in     size_t  cbDest,
    __in     LPCTSTR pszSrc          OPTIONAL,
    __deref_opt_out_bcount(*pcbRemaining)    LPTSTR* ppszDestEnd     OPTIONAL,
    __out_opt    size_t* pcbRemaining    OPTIONAL,
    __in     DWORD   dwFlags
    );

Routine Description:

    This routine is a safer version of the C built-in function 'strcat' with
    some additional parameters.  In addition to functionality provided by
    StringCbCat, this routine also returns a pointer to the end of the
    destination string and the number of bytes left in the destination string
    including the null terminator. The flags parameter allows additional controls.

Arguments:

    pszDest         -   destination string which must be null terminated

    cbDest          -   size of destination buffer in bytes.
                        length must be ((_tcslen(pszDest) + _tcslen(pszSrc) + 1) * sizeof(TCHAR)
                        to hold all of the combine string plus the null
                        terminator.

    pszSrc          -   source string which must be null terminated

    ppszDestEnd     -   if ppszDestEnd is non-null, the function will return a
                        pointer to the end of the destination string.  If the
                        function appended any data, the result will point to the
                        null termination character

    pcbRemaining    -   if pcbRemaining is non-null, the function will return
                        the number of bytes left in the destination string,
                        including the null terminator

    dwFlags         -   controls some details of the string copy:

        STRSAFE_FILL_BEHIND_NULL
                    if the function succeeds, the low byte of dwFlags will be
                    used to fill the uninitialize part of destination buffer
                    behind the null terminator

        STRSAFE_IGNORE_NULLS
                    treat NULL string pointers like empty strings (TEXT("")).
                    this flag is useful for emulating functions like lstrcat

        STRSAFE_FILL_ON_FAILURE
                    if the function fails, the low byte of dwFlags will be
                    used to fill all of the destination buffer, and it will
                    be null terminated. This will overwrite any pre-existing
                    or truncated string

        STRSAFE_NULL_ON_FAILURE
                    if the function fails, the destination buffer will be set
                    to the empty string. This will overwrite any pre-existing or
                    truncated string

        STRSAFE_NO_TRUNCATION
                    if the function returns STRSAFE_E_INSUFFICIENT_BUFFER, pszDest
                    will not contain a truncated string, it will remain unchanged.

Notes:
    Behavior is undefined if source and destination strings overlap.

    pszDest and pszSrc should not be NULL unless the STRSAFE_IGNORE_NULLS flag
    is specified.  If STRSAFE_IGNORE_NULLS is passed, both pszDest and pszSrc
    may be NULL.  An error may still be returned even though NULLS are ignored
    due to insufficient space.

Return Value:

    S_OK           -   if there was source data and it was all concatenated
                       and the resultant dest string was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the operation
                       failed due to insufficient space. When this error
                       occurs, the destination buffer is modified to contain
                       a truncated version of the ideal result and is null
                       terminated. This is useful for situations where
                       truncation is ok.

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function

--*/
#ifdef UNICODE
#define StringCbCatEx  StringCbCatExW
#else
#define StringCbCatEx  StringCbCatExA
#endif // !UNICODE

STRSAFEAPI
StringCbCatExA(
    __inout_bcount(cbDest) STRSAFE_LPSTR pszDest,
    __in size_t cbDest,
    __in STRSAFE_LPCSTR pszSrc,
    __deref_opt_out_bcount(*pcbRemaining) STRSAFE_LPSTR* ppszDestEnd,
    __out_opt size_t* pcbRemaining,
    __in DWORD dwFlags)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(char);
    size_t cchDestLength;

    hr = StringExValidateDestA(&pszDest,
                               &cchDest,
                               &cchDestLength,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPSTR pszDestEnd = pszDest + cchDestLength;
        size_t cchRemaining = cchDest - cchDestLength;

        hr = StringExValidateSrcA(&pszSrc, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;
            }
            else if (cchRemaining <= 1)
            {
                // only fail if there was actually src data to append
                if (*pszSrc != '\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchCopied = 0;

                hr = StringCopyWorkerA(pszDestEnd,
                                       cchRemaining,
                                       &cchCopied,
                                       pszSrc,
                                       STRSAFE_MAX_LENGTH);

                pszDestEnd = pszDestEnd + cchCopied;
                cchRemaining = cchRemaining - cchCopied;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                    cbRemaining = (cchRemaining * sizeof(char)) + (cbDest % sizeof(char));

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullA(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cbDest != 0))
        {
            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsA(pszDest,
                                      cbDest,
                                      cchDestLength,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcbRemaining)
            {
                // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                *pcbRemaining = (cchRemaining * sizeof(char)) + (cbDest % sizeof(char));
            }
        }
    }

    return hr;
}

STRSAFEAPI
StringCbCatExW(
    __inout_bcount(cbDest) STRSAFE_LPWSTR pszDest,
    __in size_t cbDest,
    __in STRSAFE_LPCWSTR pszSrc,
    __deref_opt_out_bcount(*pcbRemaining) STRSAFE_LPWSTR* ppszDestEnd,
    __out_opt size_t* pcbRemaining,
    __in DWORD dwFlags)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(wchar_t);
    size_t cchDestLength;

    hr = StringExValidateDestW(&pszDest,
                               &cchDest,
                               &cchDestLength,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPWSTR pszDestEnd = pszDest + cchDestLength;
        size_t cchRemaining = cchDest - cchDestLength;

        hr = StringExValidateSrcW(&pszSrc, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;
            }
            else if (cchRemaining <= 1)
            {
                // only fail if there was actually src data to append
                if (*pszSrc != L'\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchCopied = 0;

                hr = StringCopyWorkerW(pszDestEnd,
                                       cchRemaining,
                                       &cchCopied,
                                       pszSrc,
                                       STRSAFE_MAX_LENGTH);

                pszDestEnd = pszDestEnd + cchCopied;
                cchRemaining = cchRemaining - cchCopied;

                if (SUCCEEDED(hr) && (dwFlags & STRSAFE_FILL_BEHIND_NULL))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                    cbRemaining = (cchRemaining * sizeof(wchar_t)) + (cbDest % sizeof(wchar_t));

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullW(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cbDest != 0))
        {
            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsW(pszDest,
                                      cbDest,
                                      cchDestLength,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcbRemaining)
            {
                // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                *pcbRemaining = (cchRemaining * sizeof(wchar_t)) + (cbDest % sizeof(wchar_t));
            }
        }
    }

    return hr;
}
#endif  // !STRSAFE_NO_CB_FUNCTIONS


#ifndef STRSAFE_NO_CCH_FUNCTIONS
/*++

STDAPI
StringCchCatN(
    __inout_ecount(cchDest) LPTSTR  pszDest,
    __in     size_t  cchDest,
    __in     LPCTSTR pszSrc,
    __in     size_t  cchToAppend
    );

Routine Description:

    This routine is a safer version of the C built-in function 'strncat'.
    The size of the destination buffer (in characters) is a parameter as well as
    the maximum number of characters to append, excluding the null terminator.
    This function will not write past the end of the destination buffer and it will
    ALWAYS null terminate pszDest (unless it is zero length).

    This function returns a hresult, and not a pointer.  It returns
    S_OK if all of pszSrc or the first cchToAppend characters were appended
    to the destination string and it was null terminated, otherwise it will
    return a failure code. In failure cases as much of pszSrc will be appended
    to pszDest as possible, and pszDest will be null terminated.

Arguments:

    pszDest         -   destination string which must be null terminated

    cchDest         -   size of destination buffer in characters.
                        length must be (_tcslen(pszDest) + min(cchToAppend, _tcslen(pszSrc)) + 1)
                        to hold all of the combine string plus the null
                        terminator.

    pszSrc          -   source string

    cchToAppend     -   maximum number of characters to append

Notes:
    Behavior is undefined if source and destination strings overlap.

    pszDest and pszSrc should not be NULL. See StringCchCatNEx if you require
    the handling of NULL values.

Return Value:

    S_OK           -   if all of pszSrc or the first cchToAppend characters
                       were concatenated to pszDest and the resultant dest
                       string was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the operation
                       failed due to insufficient space. When this error
                       occurs, the destination buffer is modified to contain
                       a truncated version of the ideal result and is null
                       terminated. This is useful for situations where
                       truncation is ok.

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function

--*/
#ifdef UNICODE
#define StringCchCatN  StringCchCatNW
#else
#define StringCchCatN  StringCchCatNA
#endif // !UNICODE

STRSAFEAPI
StringCchCatNA(
    __inout_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __in STRSAFE_LPCSTR pszSrc,
    __in size_t cchToAppend)
{
    HRESULT hr;
    size_t cchDestLength;

    hr = StringValidateDestA(pszDest, cchDest, &cchDestLength, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        if (cchToAppend > STRSAFE_MAX_LENGTH)
        {
            hr = STRSAFE_E_INVALID_PARAMETER;
        }
        else
        {
            hr = StringCopyWorkerA(pszDest + cchDestLength,
                                   cchDest - cchDestLength,
                                   NULL,
                                   pszSrc,
                                   cchToAppend);
        }
    }

    return hr;
}

STRSAFEAPI
StringCchCatNW(
    __inout_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __in STRSAFE_LPCWSTR pszSrc,
    __in size_t cchToAppend)
{
    HRESULT hr;
    size_t cchDestLength;

    hr = StringValidateDestW(pszDest, cchDest, &cchDestLength, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        if (cchToAppend > STRSAFE_MAX_LENGTH)
        {
            hr = STRSAFE_E_INVALID_PARAMETER;
        }
        else
        {
            hr = StringCopyWorkerW(pszDest + cchDestLength,
                                   cchDest - cchDestLength,
                                   NULL,
                                   pszSrc,
                                   cchToAppend);
        }
    }

    return hr;
}
#endif  // !STRSAFE_NO_CCH_FUNCTIONS


#ifndef STRSAFE_NO_CB_FUNCTIONS
/*++

STDAPI
StringCbCatN(
    __inout_bcount(cbDest) LPTSTR  pszDest,
    __in     size_t  cbDest,
    __in     LPCTSTR pszSrc,
    __in     size_t  cbToAppend
    );

Routine Description:

    This routine is a safer version of the C built-in function 'strncat'.
    The size of the destination buffer (in bytes) is a parameter as well as
    the maximum number of bytes to append, excluding the null terminator.
    This function will not write past the end of the destination buffer and it will
    ALWAYS null terminate pszDest (unless it is zero length).

    This function returns a hresult, and not a pointer.  It returns
    S_OK if all of pszSrc or the first cbToAppend bytes were appended
    to the destination string and it was null terminated, otherwise it will
    return a failure code. In failure cases as much of pszSrc will be appended
    to pszDest as possible, and pszDest will be null terminated.

Arguments:

    pszDest         -   destination string which must be null terminated

    cbDest          -   size of destination buffer in bytes.
                        length must be ((_tcslen(pszDest) + min(cbToAppend / sizeof(TCHAR), _tcslen(pszSrc)) + 1) * sizeof(TCHAR)
                        to hold all of the combine string plus the null
                        terminator.

    pszSrc          -   source string

    cbToAppend      -   maximum number of bytes to append

Notes:
    Behavior is undefined if source and destination strings overlap.

    pszDest and pszSrc should not be NULL. See StringCbCatNEx if you require
    the handling of NULL values.

Return Value:

    S_OK           -   if all of pszSrc or the first cbToAppend bytes were
                       concatenated to pszDest and the resultant dest string
                       was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the operation
                       failed due to insufficient space. When this error
                       occurs, the destination buffer is modified to contain
                       a truncated version of the ideal result and is null
                       terminated. This is useful for situations where
                       truncation is ok.

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function

--*/
#ifdef UNICODE
#define StringCbCatN  StringCbCatNW
#else
#define StringCbCatN  StringCbCatNA
#endif // !UNICODE

STRSAFEAPI
StringCbCatNA(
    __inout_bcount(cbDest) STRSAFE_LPSTR pszDest,
    __in size_t cbDest,
    __in STRSAFE_LPCSTR pszSrc,
    __in size_t cbToAppend)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(char);
    size_t cchDestLength;

    hr = StringValidateDestA(pszDest, cchDest, &cchDestLength, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        size_t cchToAppend = cbToAppend / sizeof(char);

        if (cchToAppend > STRSAFE_MAX_LENGTH)
        {
            hr = STRSAFE_E_INVALID_PARAMETER;
        }
        else
        {
            hr = StringCopyWorkerA(pszDest + cchDestLength,
                                   cchDest - cchDestLength,
                                   NULL,
                                   pszSrc,
                                   cchToAppend);
        }
    }

    return hr;
}

STRSAFEAPI
StringCbCatNW(
    __inout_bcount(cbDest) STRSAFE_LPWSTR pszDest,
    __in size_t cbDest,
    __in STRSAFE_LPCWSTR pszSrc,
    __in size_t cbToAppend)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(wchar_t);
    size_t cchDestLength;

    hr = StringValidateDestW(pszDest, cchDest, &cchDestLength, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        size_t cchToAppend = cbToAppend / sizeof(wchar_t);

        if (cchToAppend > STRSAFE_MAX_LENGTH)
        {
            hr = STRSAFE_E_INVALID_PARAMETER;
        }
        else
        {
            hr = StringCopyWorkerW(pszDest + cchDestLength,
                                   cchDest - cchDestLength,
                                   NULL,
                                   pszSrc,
                                   cchToAppend);
        }
    }

    return hr;
}
#endif  // !STRSAFE_NO_CB_FUNCTIONS


#ifndef STRSAFE_NO_CCH_FUNCTIONS
/*++

STDAPI
StringCchCatNEx(
    __inout_ecount(cchDest) LPTSTR  pszDest         OPTIONAL,
    __in     size_t  cchDest,
    __in     LPCTSTR pszSrc          OPTIONAL,
    __in     size_t  cchToAppend,
    __deref_opt_out_ecount(*pcchRemaining)    LPTSTR* ppszDestEnd     OPTIONAL,
    __out_opt    size_t* pcchRemaining   OPTIONAL,
    __in     DWORD   dwFlags
    );

Routine Description:

    This routine is a safer version of the C built-in function 'strncat', with
    some additional parameters.  In addition to functionality provided by
    StringCchCatN, this routine also returns a pointer to the end of the
    destination string and the number of characters left in the destination string
    including the null terminator. The flags parameter allows additional controls.

Arguments:

    pszDest         -   destination string which must be null terminated

    cchDest         -   size of destination buffer in characters.
                        length must be (_tcslen(pszDest) + min(cchToAppend, _tcslen(pszSrc)) + 1)
                        to hold all of the combine string plus the null
                        terminator.

    pszSrc          -   source string

    cchToAppend     -   maximum number of characters to append

    ppszDestEnd     -   if ppszDestEnd is non-null, the function will return a
                        pointer to the end of the destination string.  If the
                        function appended any data, the result will point to the
                        null termination character

    pcchRemaining   -   if pcchRemaining is non-null, the function will return the
                        number of characters left in the destination string,
                        including the null terminator

    dwFlags         -   controls some details of the string copy:

        STRSAFE_FILL_BEHIND_NULL
                    if the function succeeds, the low byte of dwFlags will be
                    used to fill the uninitialize part of destination buffer
                    behind the null terminator

        STRSAFE_IGNORE_NULLS
                    treat NULL string pointers like empty strings (TEXT(""))

        STRSAFE_FILL_ON_FAILURE
                    if the function fails, the low byte of dwFlags will be
                    used to fill all of the destination buffer, and it will
                    be null terminated. This will overwrite any pre-existing
                    or truncated string

        STRSAFE_NULL_ON_FAILURE
                    if the function fails, the destination buffer will be set
                    to the empty string. This will overwrite any pre-existing or
                    truncated string

        STRSAFE_NO_TRUNCATION
                    if the function returns STRSAFE_E_INSUFFICIENT_BUFFER, pszDest
                    will not contain a truncated string, it will remain unchanged.

Notes:
    Behavior is undefined if source and destination strings overlap.

    pszDest and pszSrc should not be NULL unless the STRSAFE_IGNORE_NULLS flag
    is specified.  If STRSAFE_IGNORE_NULLS is passed, both pszDest and pszSrc
    may be NULL.  An error may still be returned even though NULLS are ignored
    due to insufficient space.

Return Value:

    S_OK           -   if all of pszSrc or the first cchToAppend characters
                       were concatenated to pszDest and the resultant dest
                       string was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the operation
                       failed due to insufficient space. When this error
                       occurs, the destination buffer is modified to contain
                       a truncated version of the ideal result and is null
                       terminated. This is useful for situations where
                       truncation is ok.

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function

--*/
#ifdef UNICODE
#define StringCchCatNEx  StringCchCatNExW
#else
#define StringCchCatNEx  StringCchCatNExA
#endif // !UNICODE

STRSAFEAPI
StringCchCatNExA(
    __inout_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __in STRSAFE_LPCSTR pszSrc,
    __in size_t cchToAppend,
    __deref_opt_out_ecount(*pcchRemaining) STRSAFE_LPSTR* ppszDestEnd,
    __out_opt size_t* pcchRemaining,
    __in DWORD dwFlags)
{
    HRESULT hr;
    size_t cchDestLength;

    hr = StringExValidateDestA(&pszDest,
                               &cchDest,
                               &cchDestLength,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPSTR pszDestEnd = pszDest + cchDestLength;
        size_t cchRemaining = cchDest - cchDestLength;

        hr = StringExValidateSrcA(&pszSrc, &cchToAppend, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;
            }
            else if (cchRemaining <= 1)
            {
                // only fail if there was actually src data to append
                if ((cchToAppend != 0) && (*pszSrc != '\0'))
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchCopied = 0;

                hr = StringCopyWorkerA(pszDestEnd,
                                       cchRemaining,
                                       &cchCopied,
                                       pszSrc,
                                       cchToAppend);

                pszDestEnd = pszDestEnd + cchCopied;
                cchRemaining = cchRemaining - cchCopied;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                    cbRemaining = cchRemaining * sizeof(char);

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullA(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cchDest != 0))
        {
            size_t cbDest;

            // safe to multiply cchDest * sizeof(char) since cchDest < STRSAFE_MAX_CCH and sizeof(char) is 1
            cbDest = cchDest * sizeof(char);

            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsA(pszDest,
                                      cbDest,
                                      cchDestLength,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcchRemaining)
            {
                *pcchRemaining = cchRemaining;
            }
        }
    }

    return hr;
}

STRSAFEAPI
StringCchCatNExW(
    __inout_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __in STRSAFE_LPCWSTR pszSrc,
    __in size_t cchToAppend,
    __deref_opt_out_ecount(*pcchRemaining) STRSAFE_LPWSTR* ppszDestEnd,
    __out_opt size_t* pcchRemaining,
    __in DWORD dwFlags)
{
    HRESULT hr;
    size_t cchDestLength;

    hr = StringExValidateDestW(&pszDest,
                               &cchDest,
                               &cchDestLength,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPWSTR pszDestEnd = pszDest + cchDestLength;
        size_t cchRemaining = cchDest - cchDestLength;

        hr = StringExValidateSrcW(&pszSrc, &cchToAppend, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;
            }
            else if (cchRemaining <= 1)
            {
                // only fail if there was actually src data to append
                if ((cchToAppend != 0) && (*pszSrc != L'\0'))
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchCopied = 0;

                hr = StringCopyWorkerW(pszDestEnd,
                                       cchRemaining,
                                       &cchCopied,
                                       pszSrc,
                                       cchToAppend);

                pszDestEnd = pszDestEnd + cchCopied;
                cchRemaining = cchRemaining - cchCopied;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                    cbRemaining = cchRemaining * sizeof(wchar_t);

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullW(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cchDest != 0))
        {
            size_t cbDest;

            // safe to multiply cchDest * sizeof(wchar_t) since cchDest < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
            cbDest = cchDest * sizeof(wchar_t);

            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsW(pszDest,
                                      cbDest,
                                      cchDestLength,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcchRemaining)
            {
                *pcchRemaining = cchRemaining;
            }
        }
    }

    return hr;
}
#endif  // !STRSAFE_NO_CCH_FUNCTIONS


#ifndef STRSAFE_NO_CB_FUNCTIONS
/*++

STDAPI
StringCbCatNEx(
    __inout_bcount(cbDest) LPTSTR  pszDest         OPTIONAL,
    __in     size_t  cbDest,
    __in     LPCTSTR pszSrc          OPTIONAL,
    __in     size_t  cbToAppend,
    __deref_opt_out_bcount(*pcbRemaining)    LPTSTR* ppszDestEnd     OPTIONAL,
    __out_opt    size_t* pcchRemaining   OPTIONAL,
    __in     DWORD   dwFlags
    );

Routine Description:

    This routine is a safer version of the C built-in function 'strncat', with
    some additional parameters.  In addition to functionality provided by
    StringCbCatN, this routine also returns a pointer to the end of the
    destination string and the number of bytes left in the destination string
    including the null terminator. The flags parameter allows additional controls.

Arguments:

    pszDest         -   destination string which must be null terminated

    cbDest          -   size of destination buffer in bytes.
                        length must be ((_tcslen(pszDest) + min(cbToAppend / sizeof(TCHAR), _tcslen(pszSrc)) + 1) * sizeof(TCHAR)
                        to hold all of the combine string plus the null
                        terminator.

    pszSrc          -   source string

    cbToAppend      -   maximum number of bytes to append

    ppszDestEnd     -   if ppszDestEnd is non-null, the function will return a
                        pointer to the end of the destination string.  If the
                        function appended any data, the result will point to the
                        null termination character

    pcbRemaining    -   if pcbRemaining is non-null, the function will return the
                        number of bytes left in the destination string,
                        including the null terminator

    dwFlags         -   controls some details of the string copy:

        STRSAFE_FILL_BEHIND_NULL
                    if the function succeeds, the low byte of dwFlags will be
                    used to fill the uninitialize part of destination buffer
                    behind the null terminator

        STRSAFE_IGNORE_NULLS
                    treat NULL string pointers like empty strings (TEXT(""))

        STRSAFE_FILL_ON_FAILURE
                    if the function fails, the low byte of dwFlags will be
                    used to fill all of the destination buffer, and it will
                    be null terminated. This will overwrite any pre-existing
                    or truncated string

        STRSAFE_NULL_ON_FAILURE
                    if the function fails, the destination buffer will be set
                    to the empty string. This will overwrite any pre-existing or
                    truncated string

        STRSAFE_NO_TRUNCATION
                    if the function returns STRSAFE_E_INSUFFICIENT_BUFFER, pszDest
                    will not contain a truncated string, it will remain unchanged.

Notes:
    Behavior is undefined if source and destination strings overlap.

    pszDest and pszSrc should not be NULL unless the STRSAFE_IGNORE_NULLS flag
    is specified.  If STRSAFE_IGNORE_NULLS is passed, both pszDest and pszSrc
    may be NULL.  An error may still be returned even though NULLS are ignored
    due to insufficient space.

Return Value:

    S_OK           -   if all of pszSrc or the first cbToAppend bytes were
                       concatenated to pszDest and the resultant dest string
                       was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the operation
                       failed due to insufficient space. When this error
                       occurs, the destination buffer is modified to contain
                       a truncated version of the ideal result and is null
                       terminated. This is useful for situations where
                       truncation is ok.

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function

--*/
#ifdef UNICODE
#define StringCbCatNEx  StringCbCatNExW
#else
#define StringCbCatNEx  StringCbCatNExA
#endif // !UNICODE

STRSAFEAPI
StringCbCatNExA(
    __inout_bcount(cbDest) STRSAFE_LPSTR pszDest,
    __in size_t cbDest,
    __in STRSAFE_LPCSTR pszSrc,
    __in size_t cbToAppend,
    __deref_opt_out_bcount(*pcbRemaining) STRSAFE_LPSTR* ppszDestEnd,
    __out_opt size_t* pcbRemaining,
    __in DWORD dwFlags)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(char);
    size_t cchDestLength;

    hr = StringExValidateDestA(&pszDest,
                               &cchDest,
                               &cchDestLength,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPSTR pszDestEnd = pszDest + cchDestLength;
        size_t cchRemaining = cchDest - cchDestLength;
        size_t cchToAppend = cbToAppend / sizeof(char);

        hr = StringExValidateSrcA(&pszSrc, &cchToAppend, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;
            }
            else if (cchRemaining <= 1)
            {
                // only fail if there was actually src data to append
                if ((cchToAppend != 0) && (*pszSrc != '\0'))
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchCopied = 0;

                hr = StringCopyWorkerA(pszDestEnd,
                                       cchRemaining,
                                       &cchCopied,
                                       pszSrc,
                                       cchToAppend);

                pszDestEnd = pszDestEnd + cchCopied;
                cchRemaining = cchRemaining - cchCopied;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                    cbRemaining = (cchRemaining * sizeof(char)) + (cbDest % sizeof(char));

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullA(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cbDest != 0))
        {
            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsA(pszDest,
                                      cbDest,
                                      cchDestLength,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcbRemaining)
            {
                // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                *pcbRemaining = (cchRemaining * sizeof(char)) + (cbDest % sizeof(char));
            }
        }
    }

    return hr;
}

STRSAFEAPI
StringCbCatNExW(
    __inout_bcount(cbDest) STRSAFE_LPWSTR pszDest,
    __in size_t cbDest,
    __in STRSAFE_LPCWSTR pszSrc,
    __in size_t cbToAppend,
    __deref_opt_out_bcount(*pcbRemaining) STRSAFE_LPWSTR* ppszDestEnd,
    __out_opt size_t* pcbRemaining,
    __in DWORD dwFlags)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(wchar_t);
    size_t cchDestLength;

    hr = StringExValidateDestW(&pszDest,
                               &cchDest,
                               &cchDestLength,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPWSTR pszDestEnd = pszDest + cchDestLength;
        size_t cchRemaining = cchDest - cchDestLength;
        size_t cchToAppend = cbToAppend / sizeof(wchar_t);

        hr = StringExValidateSrcW(&pszSrc, &cchToAppend, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;
            }
            else if (cchRemaining <= 1)
            {
                // only fail if there was actually src data to append
                if ((cchToAppend != 0) && (*pszSrc != L'\0'))
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchCopied = 0;

                hr = StringCopyWorkerW(pszDestEnd,
                                       cchRemaining,
                                       &cchCopied,
                                       pszSrc,
                                       cchToAppend);

                pszDestEnd = pszDestEnd + cchCopied;
                cchRemaining = cchRemaining - cchCopied;

                if (SUCCEEDED(hr) && (dwFlags & STRSAFE_FILL_BEHIND_NULL))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                    cbRemaining = (cchRemaining * sizeof(wchar_t)) + (cbDest % sizeof(wchar_t));

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullW(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cbDest != 0))
        {
            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsW(pszDest,
                                      cbDest,
                                      cchDestLength,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcbRemaining)
            {
                // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                *pcbRemaining = (cchRemaining * sizeof(wchar_t)) + (cbDest % sizeof(wchar_t));
            }
        }
    }

    return hr;
}
#endif  // !STRSAFE_NO_CB_FUNCTIONS


#ifndef STRSAFE_NO_CCH_FUNCTIONS
/*++

STDAPI
StringCchVPrintf(
    __out_ecount(cchDest) LPTSTR  pszDest,
    __in  size_t  cchDest,
    __in __format_string  LPCTSTR pszFormat,
    __in  va_list argList
    );

Routine Description:

    This routine is a safer version of the C built-in function 'vsprintf'.
    The size of the destination buffer (in characters) is a parameter and
    this function will not write past the end of this buffer and it will
    ALWAYS null terminate the destination buffer (unless it is zero length).

    This function returns a hresult, and not a pointer.  It returns
    S_OK if the string was printed without truncation and null terminated,
    otherwise it will return a failure code. In failure cases it will return
    a truncated version of the ideal result.

Arguments:

    pszDest     -  destination string

    cchDest     -  size of destination buffer in characters
                   length must be sufficient to hold the resulting formatted
                   string, including the null terminator.

    pszFormat   -  format string which must be null terminated

    argList     -  va_list from the variable arguments according to the
                   stdarg.h convention

Notes:
    Behavior is undefined if destination, format strings or any arguments
    strings overlap.

    pszDest and pszFormat should not be NULL.  See StringCchVPrintfEx if you
    require the handling of NULL values.

Return Value:

    S_OK           -   if there was sufficient space in the dest buffer for
                       the resultant string and it was null terminated.

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the print
                       operation failed due to insufficient space. When this
                       error occurs, the destination buffer is modified to
                       contain a truncated version of the ideal result and is
                       null terminated. This is useful for situations where
                       truncation is ok.

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function

--*/
#ifdef UNICODE
#define StringCchVPrintf  StringCchVPrintfW
#else
#define StringCchVPrintf  StringCchVPrintfA
#endif // !UNICODE

STRSAFEAPI
StringCchVPrintfA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __in __format_string STRSAFE_LPCSTR pszFormat,
    __in va_list argList)
{
    HRESULT hr;

    hr = StringValidateDestA(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        hr = StringVPrintfWorkerA(pszDest,
                                  cchDest,
                                  NULL,
                                  pszFormat,
                                  argList);
    }

    return hr;
}

STRSAFEAPI
StringCchVPrintfW(
    __out_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __in __format_string STRSAFE_LPCWSTR pszFormat,
    __in va_list argList)
{
    HRESULT hr;

    hr = StringValidateDestW(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        hr = StringVPrintfWorkerW(pszDest,
                                  cchDest,
                                  NULL,
                                  pszFormat,
                                  argList);
    }

    return hr;
}
#endif  // !STRSAFE_NO_CCH_FUNCTIONS


#if defined(STRSAFE_LOCALE_FUNCTIONS) && !defined(STRSAFE_NO_CCH_FUNCTIONS)
/*++

STDAPI
StringCchVPrintf_l(
    __out_ecount(cchDest) LPTSTR  pszDest,
    __in size_t  cchDest,
    __in __format_string  LPCTSTR pszFormat,
    __in locale_t locale,
    __in va_list argList
    );

Routine Description:

    This routine is a version of StringCchVPrintf that also takes a locale.
    Please see notes for StringCchVPrintf above.

--*/
#ifdef UNICODE
#define StringCchVPrintf_l  StringCchVPrintf_lW
#else
#define StringCchVPrintf_l  StringCchVPrintf_lA
#endif // !UNICODE

STRSAFEAPI
StringCchVPrintf_lA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __in __format_string STRSAFE_LPCSTR pszFormat,
    __in _locale_t locale,
    __in va_list argList)
{
    HRESULT hr;

    hr = StringValidateDestA(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        hr = StringVPrintf_lWorkerA(pszDest,
                                    cchDest,
                                    NULL,
                                    pszFormat,
                                    locale,
                                    argList);
    }

    return hr;
}

STRSAFEAPI
StringCchVPrintf_lW(
    __out_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __in __format_string STRSAFE_LPCWSTR pszFormat,
    __in _locale_t locale,
    __in va_list argList)
{
    HRESULT hr;

    hr = StringValidateDestW(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        hr = StringVPrintf_lWorkerW(pszDest,
                                    cchDest,
                                    NULL,
                                    pszFormat,
                                    locale,
                                    argList);
    }

    return hr;
}
#endif  //  STRSAFE_LOCALE_FUNCTIONS && !STRSAFE_NO_CCH_FUNCTIONS


#ifndef STRSAFE_NO_CB_FUNCTIONS
/*++

STDAPI
StringCbVPrintf(
    __out_bcount(cbDest) LPTSTR  pszDest,
    __in size_t  cbDest,
    __in __format_string LPCTSTR pszFormat,
    __in va_list argList
    );

Routine Description:

    This routine is a safer version of the C built-in function 'vsprintf'.
    The size of the destination buffer (in bytes) is a parameter and
    this function will not write past the end of this buffer and it will
    ALWAYS null terminate the destination buffer (unless it is zero length).

    This function returns a hresult, and not a pointer.  It returns
    S_OK if the string was printed without truncation and null terminated,
    otherwise it will return a failure code. In failure cases it will return
    a truncated version of the ideal result.

Arguments:

    pszDest     -  destination string

    cbDest      -  size of destination buffer in bytes
                   length must be sufficient to hold the resulting formatted
                   string, including the null terminator.

    pszFormat   -  format string which must be null terminated

    argList     -  va_list from the variable arguments according to the
                   stdarg.h convention

Notes:
    Behavior is undefined if destination, format strings or any arguments
    strings overlap.

    pszDest and pszFormat should not be NULL.  See StringCbVPrintfEx if you
    require the handling of NULL values.


Return Value:

    S_OK           -   if there was sufficient space in the dest buffer for
                       the resultant string and it was null terminated.

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the print
                       operation failed due to insufficient space. When this
                       error occurs, the destination buffer is modified to
                       contain a truncated version of the ideal result and is
                       null terminated. This is useful for situations where
                       truncation is ok.

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function

--*/
#ifdef UNICODE
#define StringCbVPrintf  StringCbVPrintfW
#else
#define StringCbVPrintf  StringCbVPrintfA
#endif // !UNICODE

STRSAFEAPI
StringCbVPrintfA(
    __out_bcount(cbDest) STRSAFE_LPSTR pszDest,
    __in size_t cbDest,
    __in __format_string STRSAFE_LPCSTR pszFormat,
    __in va_list argList)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(char);

    hr = StringValidateDestA(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        hr = StringVPrintfWorkerA(pszDest,
                                  cchDest,
                                  NULL,
                                  pszFormat,
                                  argList);
    }

    return hr;
}

STRSAFEAPI
StringCbVPrintfW(
    __out_bcount(cbDest) STRSAFE_LPWSTR pszDest,
    __in size_t cbDest,
    __in __format_string STRSAFE_LPCWSTR pszFormat,
    __in va_list argList)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(wchar_t);

    hr = StringValidateDestW(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        hr = StringVPrintfWorkerW(pszDest,
                                  cchDest,
                                  NULL,
                                  pszFormat,
                                  argList);
    }

    return hr;
}
#endif  // !STRSAFE_NO_CB_FUNCTIONS


#if defined(STRSAFE_LOCALE_FUNCTIONS) && !defined(STRSAFE_NO_CB_FUNCTIONS)
/*++

STDAPI
StringCbVPrintf_l(
    __out_bcount(cbDest) LPTSTR  pszDest,
    __in size_t  cbDest,
    __in __format_string LPCTSTR pszFormat,
    __in local_t locale,
    __in va_list argList
    );

Routine Description:

    This routine is a version of StringCbVPrintf that also takes a locale.
    Please see notes for StringCbVPrintf above.

--*/
#ifdef UNICODE
#define StringCbVPrintf_l   StringCbVPrintf_lW
#else
#define StringCbVPrintf_l   StringCbVPrintf_lA
#endif // !UNICODE

STRSAFEAPI
StringCbVPrintf_lA(
    __out_bcount(cbDest) STRSAFE_LPSTR pszDest,
    __in size_t cbDest,
    __in __format_string STRSAFE_LPCSTR pszFormat,
    __in _locale_t locale,
    __in va_list argList)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(char);

    hr = StringValidateDestA(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        hr = StringVPrintf_lWorkerA(pszDest,
                                    cchDest,
                                    NULL,
                                    pszFormat,
                                    locale,
                                    argList);
    }

    return hr;
}

STRSAFEAPI
StringCbVPrintf_lW(
    __out_bcount(cbDest) STRSAFE_LPWSTR pszDest,
    __in size_t cbDest,
    __in __format_string STRSAFE_LPCWSTR pszFormat,
    __in _locale_t locale,
    __in va_list argList)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(wchar_t);

    hr = StringValidateDestW(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        hr = StringVPrintf_lWorkerW(pszDest,
                                    cchDest,
                                    NULL,
                                    pszFormat,
                                    locale,
                                    argList);
    }

    return hr;
}
#endif  // STRSAFE_LOCALE_FUNCTIONS && !STRSAFE_NO_CB_FUNCTIONS


#ifndef _M_CEE_PURE

#ifndef STRSAFE_NO_CCH_FUNCTIONS
/*++

STDAPI
StringCchPrintf(
    __out_ecount(cchDest) LPTSTR  pszDest,
    __in size_t  cchDest,
    __in __format_string  LPCTSTR pszFormat,
    ...
    );

Routine Description:

    This routine is a safer version of the C built-in function 'sprintf'.
    The size of the destination buffer (in characters) is a parameter and
    this function will not write past the end of this buffer and it will
    ALWAYS null terminate the destination buffer (unless it is zero length).

    This function returns a hresult, and not a pointer.  It returns
    S_OK if the string was printed without truncation and null terminated,
    otherwise it will return a failure code. In failure cases it will return
    a truncated version of the ideal result.

Arguments:

    pszDest     -  destination string

    cchDest     -  size of destination buffer in characters
                   length must be sufficient to hold the resulting formatted
                   string, including the null terminator.

    pszFormat   -  format string which must be null terminated

    ...         -  additional parameters to be formatted according to
                   the format string

Notes:
    Behavior is undefined if destination, format strings or any arguments
    strings overlap.

    pszDest and pszFormat should not be NULL.  See StringCchPrintfEx if you
    require the handling of NULL values.

Return Value:

    S_OK           -   if there was sufficient space in the dest buffer for
                       the resultant string and it was null terminated.

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the print
                       operation failed due to insufficient space. When this
                       error occurs, the destination buffer is modified to
                       contain a truncated version of the ideal result and is
                       null terminated. This is useful for situations where
                       truncation is ok.

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function

--*/
#ifdef UNICODE
#define StringCchPrintf  StringCchPrintfW
#else
#define StringCchPrintf  StringCchPrintfA
#endif // !UNICODE

STRSAFEAPI
StringCchPrintfA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __in __format_string STRSAFE_LPCSTR pszFormat,
    ...)
{
    HRESULT hr;

    hr = StringValidateDestA(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        va_list argList;

        va_start(argList, pszFormat);

        hr = StringVPrintfWorkerA(pszDest,
                                  cchDest,
                                  NULL,
                                  pszFormat,
                                  argList);

        va_end(argList);
    }

    return hr;
}

STRSAFEAPI
StringCchPrintfW(
    __out_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __in __format_string STRSAFE_LPCWSTR pszFormat,
    ...)
{
    HRESULT hr;

    hr = StringValidateDestW(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        va_list argList;

        va_start(argList, pszFormat);

        hr = StringVPrintfWorkerW(pszDest,
                                  cchDest,
                                  NULL,
                                  pszFormat,
                                  argList);

        va_end(argList);
    }

    return hr;
}
#endif  // !STRSAFE_NO_CCH_FUNCTIONS


#if defined(STRSAFE_LOCALE_FUNCTIONS) && !defined(STRSAFE_NO_CCH_FUNCTIONS)
/*++

STDAPI
StringCchPrintf_l(
    __out_ecount(cchDest) LPTSTR  pszDest,
    __in size_t  cchDest,
    __in __format_string  LPCTSTR pszFormat,
    __in  locale_t locale,
    ...
    );

Routine Description:

    This routine is a version of a StringCchPrintf_l that also takes a locale.
    Please see notes for StringCchPrintf above.

--*/
#ifdef UNICODE
#define StringCchPrintf_l   StringCchPrintf_lW
#else
#define StringCchPrintf_l   StringCchPrintf_lA
#endif // !UNICODE

STRSAFEAPI
StringCchPrintf_lA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __in __format_string STRSAFE_LPCSTR pszFormat,
    __in _locale_t locale,
    ...)
{
    HRESULT hr;

    hr = StringValidateDestA(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        va_list argList;

        va_start(argList, locale);

        hr = StringVPrintf_lWorkerA(pszDest,
                                    cchDest,
                                    NULL,
                                    pszFormat,
                                    locale,
                                    argList);

        va_end(argList);
    }

    return hr;
}

STRSAFEAPI
StringCchPrintf_lW(
    __out_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __in __format_string STRSAFE_LPCWSTR pszFormat,
    __in _locale_t locale,
    ...)
{
    HRESULT hr;

    hr = StringValidateDestW(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        va_list argList;

        va_start(argList, locale);

        hr = StringVPrintf_lWorkerW(pszDest,
                                    cchDest,
                                    NULL,
                                    pszFormat,
                                    locale,
                                    argList);

        va_end(argList);
    }

    return hr;
}
#endif  // STRSAFE_LOCALE_FUNCTIONS && !STRSAFE_NO_CCH_FUNCTIONS


#ifndef STRSAFE_NO_CB_FUNCTIONS
/*++

STDAPI
StringCbPrintf(
    __out_bcount(cbDest) LPTSTR  pszDest,
    __in size_t  cbDest,
    __in __format_string LPCTSTR pszFormat,
    ...
    );

Routine Description:

    This routine is a safer version of the C built-in function 'sprintf'.
    The size of the destination buffer (in bytes) is a parameter and
    this function will not write past the end of this buffer and it will
    ALWAYS null terminate the destination buffer (unless it is zero length).

    This function returns a hresult, and not a pointer.  It returns
    S_OK if the string was printed without truncation and null terminated,
    otherwise it will return a failure code. In failure cases it will return
    a truncated version of the ideal result.

Arguments:

    pszDest     -  destination string

    cbDest      -  size of destination buffer in bytes
                   length must be sufficient to hold the resulting formatted
                   string, including the null terminator.

    pszFormat   -  format string which must be null terminated

    ...         -  additional parameters to be formatted according to
                   the format string

Notes:
    Behavior is undefined if destination, format strings or any arguments
    strings overlap.

    pszDest and pszFormat should not be NULL.  See StringCbPrintfEx if you
    require the handling of NULL values.


Return Value:

    S_OK           -   if there was sufficient space in the dest buffer for
                       the resultant string and it was null terminated.

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the print
                       operation failed due to insufficient space. When this
                       error occurs, the destination buffer is modified to
                       contain a truncated version of the ideal result and is
                       null terminated. This is useful for situations where
                       truncation is ok.

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function

--*/
#ifdef UNICODE
#define StringCbPrintf  StringCbPrintfW
#else
#define StringCbPrintf  StringCbPrintfA
#endif // !UNICODE

STRSAFEAPI
StringCbPrintfA(
    __out_bcount(cbDest) STRSAFE_LPSTR pszDest,
    __in size_t cbDest,
    __in __format_string STRSAFE_LPCSTR pszFormat,
    ...)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(char);

    hr = StringValidateDestA(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        va_list argList;

        va_start(argList, pszFormat);

        hr = StringVPrintfWorkerA(pszDest,
                                  cchDest,
                                  NULL,
                                  pszFormat,
                                  argList);

        va_end(argList);
    }

    return hr;
}

STRSAFEAPI
StringCbPrintfW(
    __out_bcount(cbDest) STRSAFE_LPWSTR pszDest,
    __in size_t cbDest,
    __in __format_string STRSAFE_LPCWSTR pszFormat,
    ...)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(wchar_t);

    hr = StringValidateDestW(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        va_list argList;

        va_start(argList, pszFormat);

        hr = StringVPrintfWorkerW(pszDest,
                                  cchDest,
                                  NULL,
                                  pszFormat,
                                  argList);

        va_end(argList);
    }

    return hr;
}
#endif  // !STRSAFE_NO_CB_FUNCTIONS


#if defined(STRSAFE_LOCALE_FUNCTIONS) && !defined(STRSAFE_NO_CB_FUNCTIONS)
/*++

STDAPI
StringCbPrintf_l(
    __out_bcount(cbDest) LPTSTR  pszDest,
    __in size_t  cbDest,
    __in __format_string LPCTSTR pszFormat,
    __in locale_t locale,
    ...
    );

Routine Description:

    This routine is a version of StringCbPrintf that also takes a locale.
    Please see notes for StringCbPrintf above.

--*/
#ifdef UNICODE
#define StringCbPrintf_l    StringCbPrintf_lW
#else
#define StringCbPrintf_l    StringCbPrintf_lA
#endif // !UNICODE

STRSAFEAPI
StringCbPrintf_lA(
    __out_bcount(cbDest) STRSAFE_LPSTR pszDest,
    __in size_t cbDest,
    __in __format_string STRSAFE_LPCSTR pszFormat,
    __in _locale_t locale,
    ...)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(char);

    hr = StringValidateDestA(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        va_list argList;

        va_start(argList, locale);

        hr = StringVPrintf_lWorkerA(pszDest,
                                    cchDest,
                                    NULL,
                                    pszFormat,
                                    locale,
                                    argList);

        va_end(argList);
    }

    return hr;
}

STRSAFEAPI
StringCbPrintf_lW(
    __out_bcount(cbDest) STRSAFE_LPWSTR pszDest,
    __in size_t cbDest,
    __in __format_string STRSAFE_LPCWSTR pszFormat,
    __in _locale_t locale,
    ...)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(wchar_t);

    hr = StringValidateDestW(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        va_list argList;

        va_start(argList, locale);

        hr = StringVPrintf_lWorkerW(pszDest,
                                    cchDest,
                                    NULL,
                                    pszFormat,
                                    locale,
                                    argList);

        va_end(argList);
    }

    return hr;
}
#endif  // STRSAFE_LOCALE_FUNCTIONS && !STRSAFE_NO_CB_FUNCTIONS


#ifndef STRSAFE_NO_CCH_FUNCTIONS
/*++

STDAPI
StringCchPrintfEx(
    __out_ecount(cchDest) LPTSTR  pszDest         OPTIONAL,
    __in  size_t  cchDest,
    __deref_opt_out_ecount(*pcchRemaining) LPTSTR* ppszDestEnd     OPTIONAL,
    __out_opt size_t* pcchRemaining   OPTIONAL,
    __in DWORD   dwFlags,
    __in __format_string LPCTSTR pszFormat       OPTIONAL,
    ...
    );

Routine Description:

    This routine is a safer version of the C built-in function 'sprintf' with
    some additional parameters.  In addition to functionality provided by
    StringCchPrintf, this routine also returns a pointer to the end of the
    destination string and the number of characters left in the destination string
    including the null terminator. The flags parameter allows additional controls.

Arguments:

    pszDest         -   destination string

    cchDest         -   size of destination buffer in characters.
                        length must be sufficient to contain the resulting
                        formatted string plus the null terminator.

    ppszDestEnd     -   if ppszDestEnd is non-null, the function will return a
                        pointer to the end of the destination string.  If the
                        function printed any data, the result will point to the
                        null termination character

    pcchRemaining   -   if pcchRemaining is non-null, the function will return
                        the number of characters left in the destination string,
                        including the null terminator

    dwFlags         -   controls some details of the string copy:

        STRSAFE_FILL_BEHIND_NULL
                    if the function succeeds, the low byte of dwFlags will be
                    used to fill the uninitialize part of destination buffer
                    behind the null terminator

        STRSAFE_IGNORE_NULLS
                    treat NULL string pointers like empty strings (TEXT(""))

        STRSAFE_FILL_ON_FAILURE
                    if the function fails, the low byte of dwFlags will be
                    used to fill all of the destination buffer, and it will
                    be null terminated. This will overwrite any truncated
                    string returned when the failure is
                    STRSAFE_E_INSUFFICIENT_BUFFER

        STRSAFE_NO_TRUNCATION /
        STRSAFE_NULL_ON_FAILURE
                    if the function fails, the destination buffer will be set
                    to the empty string. This will overwrite any truncated string
                    returned when the failure is STRSAFE_E_INSUFFICIENT_BUFFER.

    pszFormat       -   format string which must be null terminated

    ...             -   additional parameters to be formatted according to
                        the format string

Notes:
    Behavior is undefined if destination, format strings or any arguments
    strings overlap.

    pszDest and pszFormat should not be NULL unless the STRSAFE_IGNORE_NULLS
    flag is specified.  If STRSAFE_IGNORE_NULLS is passed, both pszDest and
    pszFormat may be NULL.  An error may still be returned even though NULLS
    are ignored due to insufficient space.

Return Value:

    S_OK           -   if there was sufficient space in the dest buffer for
                       the resultant string and it was null terminated.

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the print
                       operation failed due to insufficient space. When this
                       error occurs, the destination buffer is modified to
                       contain a truncated version of the ideal result and is
                       null terminated. This is useful for situations where
                       truncation is ok.

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function

--*/
#ifdef UNICODE
#define StringCchPrintfEx  StringCchPrintfExW
#else
#define StringCchPrintfEx  StringCchPrintfExA
#endif // !UNICODE

STRSAFEAPI
StringCchPrintfExA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __deref_opt_out_ecount(*pcchRemaining) STRSAFE_LPSTR* ppszDestEnd,
    __out_opt size_t* pcchRemaining,
    __in DWORD dwFlags,
    __in __format_string STRSAFE_LPCSTR pszFormat,
    ...)
{
    HRESULT hr;

    hr = StringExValidateDestA(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcA(&pszFormat, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = '\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually a non-empty format string
                if (*pszFormat != '\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchNewDestLength = 0;
                va_list argList;

                va_start(argList, pszFormat);

                hr = StringVPrintfWorkerA(pszDest,
                                          cchDest,
                                          &cchNewDestLength,
                                          pszFormat,
                                          argList);

                va_end(argList);

                pszDestEnd = pszDest + cchNewDestLength;
                cchRemaining = cchDest - cchNewDestLength;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                    cbRemaining = cchRemaining * sizeof(char);

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullA(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = '\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cchDest != 0))
        {
            size_t cbDest;

             // safe to multiply cchDest * sizeof(char) since cchDest < STRSAFE_MAX_CCH and sizeof(char) is 1
            cbDest = cchDest * sizeof(char);

            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsA(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcchRemaining)
            {
                *pcchRemaining = cchRemaining;
            }
        }
    }

    return hr;
}

STRSAFEAPI
StringCchPrintfExW(
    __out_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __deref_opt_out_ecount(*pcchRemaining) STRSAFE_LPWSTR* ppszDestEnd,
    __out_opt size_t* pcchRemaining,
    __in DWORD dwFlags,
    __in __format_string STRSAFE_LPCWSTR pszFormat,
    ...)
{
    HRESULT hr;

    hr = StringExValidateDestW(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPWSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcW(&pszFormat, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = L'\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually a non-empty format string
                if (*pszFormat != L'\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchNewDestLength = 0;
                va_list argList;

                va_start(argList, pszFormat);

                hr = StringVPrintfWorkerW(pszDest,
                                          cchDest,
                                          &cchNewDestLength,
                                          pszFormat,
                                          argList);

                va_end(argList);

                pszDestEnd = pszDest + cchNewDestLength;
                cchRemaining = cchDest - cchNewDestLength;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                    cbRemaining = cchRemaining * sizeof(wchar_t);

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullW(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = L'\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cchDest != 0))
        {
            size_t cbDest;

            // safe to multiply cchDest * sizeof(wchar_t) since cchDest < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
            cbDest = cchDest * sizeof(wchar_t);

            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsW(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcchRemaining)
            {
                *pcchRemaining = cchRemaining;
            }
        }
    }

    return hr;
}
#endif  // !STRSAFE_NO_CCH_FUNCTIONS


#if defined(STRSAFE_LOCALE_FUNCTIONS) && !defined(STRSAFE_NO_CCH_FUNCTIONS)
/*++

STDAPI
StringCchPrintf_lEx(
    __out_ecount(cchDest) LPTSTR  pszDest         OPTIONAL,
    __in size_t  cchDest,
    __deref_opt_out_ecount(*pcchRemaining) LPTSTR* ppszDestEnd     OPTIONAL,
    __out_opt size_t* pcchRemaining   OPTIONAL,
    __in DWORD   dwFlags,
    __in __format_string LPCTSTR pszFormat       OPTIONAL,
    __in locale_t locale,
    ...
    );

Routine Description:

    This routine is a version of StringCchPrintfEx that also takes a locale.
    Please see notes for StringCchPrintfEx above.

--*/
#ifdef UNICODE
#define StringCchPrintf_lEx StringCchPrintf_lExW
#else
#define StringCchPrintf_lEx StringCchPrintf_lExA
#endif // !UNICODE

STRSAFEAPI
StringCchPrintf_lExA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __deref_opt_out_ecount(*pcchRemaining) STRSAFE_LPSTR* ppszDestEnd,
    __out_opt size_t* pcchRemaining,
    __in DWORD dwFlags,
    __in __format_string STRSAFE_LPCSTR pszFormat,
    __in _locale_t locale,
    ...)
{
    HRESULT hr;

    hr = StringExValidateDestA(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcA(&pszFormat, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = '\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually a non-empty format string
                if (*pszFormat != '\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchNewDestLength = 0;
                va_list argList;

                va_start(argList, locale);

                hr = StringVPrintf_lWorkerA(pszDest,
                                            cchDest,
                                            &cchNewDestLength,
                                            pszFormat,
                                            locale,
                                            argList);

                va_end(argList);

                pszDestEnd = pszDest + cchNewDestLength;
                cchRemaining = cchDest - cchNewDestLength;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                    cbRemaining = cchRemaining * sizeof(char);

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullA(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = '\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cchDest != 0))
        {
            size_t cbDest;

             // safe to multiply cchDest * sizeof(char) since cchDest < STRSAFE_MAX_CCH and sizeof(char) is 1
            cbDest = cchDest * sizeof(char);

            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsA(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcchRemaining)
            {
                *pcchRemaining = cchRemaining;
            }
        }
    }

    return hr;
}

STRSAFEAPI
StringCchPrintf_lExW(
    __out_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __deref_opt_out_ecount(*pcchRemaining) STRSAFE_LPWSTR* ppszDestEnd,
    __out_opt size_t* pcchRemaining,
    __in DWORD dwFlags,
    __in __format_string STRSAFE_LPCWSTR pszFormat,
    __in _locale_t locale,
    ...)
{
    HRESULT hr;

    hr = StringExValidateDestW(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPWSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcW(&pszFormat, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = L'\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually a non-empty format string
                if (*pszFormat != L'\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchNewDestLength = 0;
                va_list argList;

                va_start(argList, locale);

                hr = StringVPrintf_lWorkerW(pszDest,
                                            cchDest,
                                            &cchNewDestLength,
                                            pszFormat,
                                            locale,
                                            argList);

                va_end(argList);

                pszDestEnd = pszDest + cchNewDestLength;
                cchRemaining = cchDest - cchNewDestLength;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                    cbRemaining = cchRemaining * sizeof(wchar_t);

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullW(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = L'\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cchDest != 0))
        {
            size_t cbDest;

            // safe to multiply cchDest * sizeof(wchar_t) since cchDest < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
            cbDest = cchDest * sizeof(wchar_t);

            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsW(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcchRemaining)
            {
                *pcchRemaining = cchRemaining;
            }
        }
    }

    return hr;
}
#endif  // STRSAFE_LOCALE_FUNCTIONS && !STRSAFE_NO_CCH_FUNCTIONS


#ifndef STRSAFE_NO_CB_FUNCTIONS
/*++

STDAPI
StringCbPrintfEx(
    __out_bcount(cbDest) LPTSTR  pszDest         OPTIONAL,
    __in size_t  cbDest,
    __deref_opt_out_bcount(*pcbRemaining) LPTSTR* ppszDestEnd     OPTIONAL,
    __out_opt size_t* pcbRemaining    OPTIONAL,
    __in DWORD   dwFlags,
    __in __format_string LPCTSTR pszFormat       OPTIONAL,
    ...
    );

Routine Description:

    This routine is a safer version of the C built-in function 'sprintf' with
    some additional parameters.  In addition to functionality provided by
    StringCbPrintf, this routine also returns a pointer to the end of the
    destination string and the number of bytes left in the destination string
    including the null terminator. The flags parameter allows additional controls.

Arguments:

    pszDest         -   destination string

    cbDest          -   size of destination buffer in bytes.
                        length must be sufficient to contain the resulting
                        formatted string plus the null terminator.

    ppszDestEnd     -   if ppszDestEnd is non-null, the function will return a
                        pointer to the end of the destination string.  If the
                        function printed any data, the result will point to the
                        null termination character

    pcbRemaining    -   if pcbRemaining is non-null, the function will return
                        the number of bytes left in the destination string,
                        including the null terminator

    dwFlags         -   controls some details of the string copy:

        STRSAFE_FILL_BEHIND_NULL
                    if the function succeeds, the low byte of dwFlags will be
                    used to fill the uninitialize part of destination buffer
                    behind the null terminator

        STRSAFE_IGNORE_NULLS
                    treat NULL string pointers like empty strings (TEXT(""))

        STRSAFE_FILL_ON_FAILURE
                    if the function fails, the low byte of dwFlags will be
                    used to fill all of the destination buffer, and it will
                    be null terminated. This will overwrite any truncated
                    string returned when the failure is
                    STRSAFE_E_INSUFFICIENT_BUFFER

        STRSAFE_NO_TRUNCATION /
        STRSAFE_NULL_ON_FAILURE
                    if the function fails, the destination buffer will be set
                    to the empty string. This will overwrite any truncated string
                    returned when the failure is STRSAFE_E_INSUFFICIENT_BUFFER.

    pszFormat       -   format string which must be null terminated

    ...             -   additional parameters to be formatted according to
                        the format string

Notes:
    Behavior is undefined if destination, format strings or any arguments
    strings overlap.

    pszDest and pszFormat should not be NULL unless the STRSAFE_IGNORE_NULLS
    flag is specified.  If STRSAFE_IGNORE_NULLS is passed, both pszDest and
    pszFormat may be NULL.  An error may still be returned even though NULLS
    are ignored due to insufficient space.

Return Value:

    S_OK           -   if there was source data and it was all concatenated and
                       the resultant dest string was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the print
                       operation failed due to insufficient space. When this
                       error occurs, the destination buffer is modified to
                       contain a truncated version of the ideal result and is
                       null terminated. This is useful for situations where
                       truncation is ok.

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function

--*/
#ifdef UNICODE
#define StringCbPrintfEx  StringCbPrintfExW
#else
#define StringCbPrintfEx  StringCbPrintfExA
#endif // !UNICODE

STRSAFEAPI
StringCbPrintfExA(
    __out_bcount(cbDest) STRSAFE_LPSTR pszDest,
    __in size_t cbDest,
    __deref_opt_out_bcount(*pcbRemaining) STRSAFE_LPSTR* ppszDestEnd,
    __out_opt size_t* pcbRemaining,
    __in DWORD dwFlags,
    __in __format_string STRSAFE_LPCSTR pszFormat,
    ...)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(char);

    hr = StringExValidateDestA(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcA(&pszFormat, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = '\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually a non-empty format string
                if (*pszFormat != '\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchNewDestLength = 0;
                va_list argList;

                va_start(argList, pszFormat);

                hr = StringVPrintfWorkerA(pszDest,
                                          cchDest,
                                          &cchNewDestLength,
                                          pszFormat,
                                          argList);

                va_end(argList);

                pszDestEnd = pszDest + cchNewDestLength;
                cchRemaining = cchDest - cchNewDestLength;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                    cbRemaining = (cchRemaining * sizeof(char)) + (cbDest % sizeof(char));

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullA(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = '\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cbDest != 0))
        {
            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsA(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcbRemaining)
            {
                // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                *pcbRemaining = (cchRemaining * sizeof(char)) + (cbDest % sizeof(char));
            }
        }
    }

    return hr;
}

STRSAFEAPI
StringCbPrintfExW(
    __out_bcount(cbDest) STRSAFE_LPWSTR pszDest,
    __in size_t cbDest,
    __deref_opt_out_bcount(*pcbRemaining) STRSAFE_LPWSTR* ppszDestEnd,
    __out_opt size_t* pcbRemaining,
    __in DWORD dwFlags,
    __in __format_string STRSAFE_LPCWSTR pszFormat,
    ...)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(wchar_t);

    hr = StringExValidateDestW(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPWSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcW(&pszFormat, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = L'\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually a non-empty format string
                if (*pszFormat != L'\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchNewDestLength = 0;
                va_list argList;

                va_start(argList, pszFormat);

                hr = StringVPrintfWorkerW(pszDest,
                                          cchDest,
                                          &cchNewDestLength,
                                          pszFormat,
                                          argList);

                va_end(argList);

                pszDestEnd = pszDest + cchNewDestLength;
                cchRemaining = cchDest - cchNewDestLength;

                if (SUCCEEDED(hr) && (dwFlags & STRSAFE_FILL_BEHIND_NULL))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                    cbRemaining = (cchRemaining * sizeof(wchar_t)) + (cbDest % sizeof(wchar_t));

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullW(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = L'\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cbDest != 0))
        {
            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsW(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcbRemaining)
            {
                // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                *pcbRemaining = (cchRemaining * sizeof(wchar_t)) + (cbDest % sizeof(wchar_t));
            }
        }
    }

    return hr;
}
#endif  // !STRSAFE_NO_CB_FUNCTIONS


#if defined(STRSAFE_LOCALE_FUNCTIONS) && !defined(STRSAFE_NO_CB_FUNCTIONS)
/*++

STDAPI
StringCbPrintf_lEx(
    __out_bcount(cbDest) LPTSTR  pszDest         OPTIONAL,
    __in size_t  cbDest,
    __deref_opt_out_bcount(*pcbRemaining) LPTSTR* ppszDestEnd     OPTIONAL,
    __out_opt size_t* pcbRemaining    OPTIONAL,
    __in DWORD   dwFlags,
    __in __format_string LPCTSTR pszFormat       OPTIONAL,
    __in locale_t locale,
    ...
    );

Routine Description:

    This routine is a version of StringCbPrintfEx that also takes a locale.
    Please seee notes for StringCbPrintfEx above.

--*/
#ifdef UNICODE
#define StringCbPrintf_lEx  StringCbPrintf_lExW
#else
#define StringCbPrintf_lEx  StringCbPrintf_lExA
#endif // !UNICODE

STRSAFEAPI
StringCbPrintf_lExA(
    __out_bcount(cbDest) STRSAFE_LPSTR pszDest,
    __in size_t cbDest,
    __deref_opt_out_bcount(*pcbRemaining) STRSAFE_LPSTR* ppszDestEnd,
    __out_opt size_t* pcbRemaining,
    __in DWORD dwFlags,
    __in __format_string STRSAFE_LPCSTR pszFormat,
    __in _locale_t locale,
    ...)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(char);

    hr = StringExValidateDestA(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcA(&pszFormat, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = '\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually a non-empty format string
                if (*pszFormat != '\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchNewDestLength = 0;
                va_list argList;

                va_start(argList, locale);

                hr = StringVPrintf_lWorkerA(pszDest,
                                            cchDest,
                                            &cchNewDestLength,
                                            pszFormat,
                                            locale,
                                            argList);

                va_end(argList);

                pszDestEnd = pszDest + cchNewDestLength;
                cchRemaining = cchDest - cchNewDestLength;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                    cbRemaining = (cchRemaining * sizeof(char)) + (cbDest % sizeof(char));

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullA(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = '\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cbDest != 0))
        {
            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsA(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcbRemaining)
            {
                // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                *pcbRemaining = (cchRemaining * sizeof(char)) + (cbDest % sizeof(char));
            }
        }
    }

    return hr;
}

STRSAFEAPI
StringCbPrintf_lExW(
    __out_bcount(cbDest) STRSAFE_LPWSTR pszDest,
    __in size_t cbDest,
    __deref_opt_out_bcount(*pcbRemaining) STRSAFE_LPWSTR* ppszDestEnd,
    __out_opt size_t* pcbRemaining,
    __in DWORD dwFlags,
    __in __format_string STRSAFE_LPCWSTR pszFormat,
    __in _locale_t locale,
    ...)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(wchar_t);

    hr = StringExValidateDestW(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPWSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcW(&pszFormat, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = L'\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually a non-empty format string
                if (*pszFormat != L'\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchNewDestLength = 0;
                va_list argList;

                va_start(argList, locale);

                hr = StringVPrintf_lWorkerW(pszDest,
                                            cchDest,
                                            &cchNewDestLength,
                                            pszFormat,
                                            locale,
                                            argList);

                va_end(argList);

                pszDestEnd = pszDest + cchNewDestLength;
                cchRemaining = cchDest - cchNewDestLength;

                if (SUCCEEDED(hr) && (dwFlags & STRSAFE_FILL_BEHIND_NULL))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                    cbRemaining = (cchRemaining * sizeof(wchar_t)) + (cbDest % sizeof(wchar_t));

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullW(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = L'\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cbDest != 0))
        {
            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsW(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcbRemaining)
            {
                // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                *pcbRemaining = (cchRemaining * sizeof(wchar_t)) + (cbDest % sizeof(wchar_t));
            }
        }
    }

    return hr;
}
#endif  // STRSAFE_LOCALE_FUNCTIONS && !STRSAFE_NO_CB_FUNCTIONS

#endif  // !_M_CEE_PURE


#ifndef STRSAFE_NO_CCH_FUNCTIONS
/*++

STDAPI
StringCchVPrintfEx(
    __out_ecount(cchDest) LPTSTR  pszDest         OPTIONAL,
    __in size_t  cchDest,
    __deref_opt_out_ecount(*pcchRemaining) LPTSTR* ppszDestEnd     OPTIONAL,
    __out_opt size_t* pcchRemaining   OPTIONAL,
    __in DWORD   dwFlags,
    __in __format_string LPCTSTR pszFormat       OPTIONAL,
    __in va_list argList
    );


Routine Description:

    This routine is a safer version of the C built-in function 'vsprintf' with
    some additional parameters.  In addition to functionality provided by
    StringCchVPrintf, this routine also returns a pointer to the end of the
    destination string and the number of characters left in the destination string
    including the null terminator. The flags parameter allows additional controls.

Arguments:

    pszDest         -   destination string

    cchDest         -   size of destination buffer in characters.
                        length must be sufficient to contain the resulting
                        formatted string plus the null terminator.

    ppszDestEnd     -   if ppszDestEnd is non-null, the function will return a
                        pointer to the end of the destination string.  If the
                        function printed any data, the result will point to the
                        null termination character

    pcchRemaining   -   if pcchRemaining is non-null, the function will return
                        the number of characters left in the destination string,
                        including the null terminator

    dwFlags         -   controls some details of the string copy:

        STRSAFE_FILL_BEHIND_NULL
                    if the function succeeds, the low byte of dwFlags will be
                    used to fill the uninitialize part of destination buffer
                    behind the null terminator

        STRSAFE_IGNORE_NULLS
                    treat NULL string pointers like empty strings (TEXT(""))

        STRSAFE_FILL_ON_FAILURE
                    if the function fails, the low byte of dwFlags will be
                    used to fill all of the destination buffer, and it will
                    be null terminated. This will overwrite any truncated
                    string returned when the failure is
                    STRSAFE_E_INSUFFICIENT_BUFFER

        STRSAFE_NO_TRUNCATION /
        STRSAFE_NULL_ON_FAILURE
                    if the function fails, the destination buffer will be set
                    to the empty string. This will overwrite any truncated string
                    returned when the failure is STRSAFE_E_INSUFFICIENT_BUFFER.

    pszFormat       -   format string which must be null terminated

    argList         -   va_list from the variable arguments according to the
                        stdarg.h convention

Notes:
    Behavior is undefined if destination, format strings or any arguments
    strings overlap.

    pszDest and pszFormat should not be NULL unless the STRSAFE_IGNORE_NULLS
    flag is specified.  If STRSAFE_IGNORE_NULLS is passed, both pszDest and
    pszFormat may be NULL.  An error may still be returned even though NULLS
    are ignored due to insufficient space.

Return Value:

    S_OK           -   if there was source data and it was all concatenated and
                       the resultant dest string was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the print
                       operation failed due to insufficient space. When this
                       error occurs, the destination buffer is modified to
                       contain a truncated version of the ideal result and is
                       null terminated. This is useful for situations where
                       truncation is ok.

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function

--*/
#ifdef UNICODE
#define StringCchVPrintfEx  StringCchVPrintfExW
#else
#define StringCchVPrintfEx  StringCchVPrintfExA
#endif // !UNICODE

STRSAFEAPI
StringCchVPrintfExA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __deref_opt_out_ecount(*pcchRemaining) STRSAFE_LPSTR* ppszDestEnd,
    __out_opt size_t* pcchRemaining,
    __in DWORD dwFlags,
    __in __format_string STRSAFE_LPCSTR pszFormat,
    __in va_list argList)
{
    HRESULT hr;

    hr = StringExValidateDestA(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcA(&pszFormat, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = '\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually a non-empty format string
                if (*pszFormat != '\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchNewDestLength = 0;

                hr = StringVPrintfWorkerA(pszDest,
                                          cchDest,
                                          &cchNewDestLength,
                                          pszFormat,
                                          argList);

                pszDestEnd = pszDest + cchNewDestLength;
                cchRemaining = cchDest - cchNewDestLength;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                    cbRemaining = cchRemaining * sizeof(char);

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullA(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = '\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cchDest != 0))
        {
            size_t cbDest;

            // safe to multiply cchDest * sizeof(char) since cchDest < STRSAFE_MAX_CCH and sizeof(char) is 1
            cbDest = cchDest * sizeof(char);

            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsA(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcchRemaining)
            {
                *pcchRemaining = cchRemaining;
            }
        }
    }

    return hr;
}

STRSAFEAPI
StringCchVPrintfExW(
    __out_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __deref_opt_out_ecount(*pcchRemaining) STRSAFE_LPWSTR* ppszDestEnd,
    __out_opt size_t* pcchRemaining,
    __in DWORD dwFlags,
    __in __format_string STRSAFE_LPCWSTR pszFormat,
    __in va_list argList)
{
    HRESULT hr;

    hr = StringExValidateDestW(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPWSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcW(&pszFormat, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = L'\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually a non-empty format string
                if (*pszFormat != L'\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchNewDestLength = 0;

                hr = StringVPrintfWorkerW(pszDest,
                                          cchDest,
                                          &cchNewDestLength,
                                          pszFormat,
                                          argList);

                pszDestEnd = pszDest + cchNewDestLength;
                cchRemaining = cchDest - cchNewDestLength;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                    cbRemaining = cchRemaining * sizeof(wchar_t);

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullW(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = L'\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cchDest != 0))
        {
            size_t cbDest;

            // safe to multiply cchDest * sizeof(wchar_t) since cchDest < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
            cbDest = cchDest * sizeof(wchar_t);

            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsW(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcchRemaining)
            {
                *pcchRemaining = cchRemaining;
            }
        }
    }

    return hr;
}
#endif  // !STRSAFE_NO_CCH_FUNCTIONS


#if defined(STRSAFE_LOCALE_FUNCTIONS) && !defined(STRSAFE_NO_CCH_FUNCTIONS)
/*++

STDAPI
StringCchVPrintf_lEx(
    __out_ecount(cchDest) LPTSTR  pszDest         OPTIONAL,
    __in size_t  cchDest,
    __deref_opt_out_ecount(*pcchRemaining) LPTSTR* ppszDestEnd     OPTIONAL,
    __out_opt size_t* pcchRemaining   OPTIONAL,
    __in DWORD   dwFlags,
    __in __format_string LPCTSTR pszFormat       OPTIONAL,
    __in locale_t locale,
    __in va_list argList
    );

Routine Description:

    This routine is a version of StringCchVPrintfEx that also takes a locale.
    Please see notes for StringCchVPrintfEx above.

--*/
#ifdef UNICODE
#define StringCchVPrintf_lEx    StringCchVPrintf_lExW
#else
#define StringCchVPrintf_lEx    StringCchVPrintf_lExA
#endif // !UNICODE

STRSAFEAPI
StringCchVPrintf_lExA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __deref_opt_out_ecount(*pcchRemaining) STRSAFE_LPSTR* ppszDestEnd,
    __out_opt size_t* pcchRemaining,
    __in DWORD dwFlags,
    __in __format_string STRSAFE_LPCSTR pszFormat,
    __in _locale_t locale,
    __in va_list argList)
{
    HRESULT hr;

    hr = StringExValidateDestA(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcA(&pszFormat, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = '\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually a non-empty format string
                if (*pszFormat != '\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchNewDestLength = 0;

                hr = StringVPrintf_lWorkerA(pszDest,
                                            cchDest,
                                            &cchNewDestLength,
                                            pszFormat,
                                            locale,
                                            argList);

                pszDestEnd = pszDest + cchNewDestLength;
                cchRemaining = cchDest - cchNewDestLength;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                    cbRemaining = cchRemaining * sizeof(char);

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullA(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = '\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cchDest != 0))
        {
            size_t cbDest;

            // safe to multiply cchDest * sizeof(char) since cchDest < STRSAFE_MAX_CCH and sizeof(char) is 1
            cbDest = cchDest * sizeof(char);

            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsA(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcchRemaining)
            {
                *pcchRemaining = cchRemaining;
            }
        }
    }

    return hr;
}

STRSAFEAPI
StringCchVPrintf_lExW(
    __out_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __deref_opt_out_ecount(*pcchRemaining) STRSAFE_LPWSTR* ppszDestEnd,
    __out_opt size_t* pcchRemaining,
    __in DWORD dwFlags,
    __in __format_string STRSAFE_LPCWSTR pszFormat,
    __in _locale_t locale,
    __in va_list argList)
{
    HRESULT hr;

    hr = StringExValidateDestW(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPWSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcW(&pszFormat, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = L'\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually a non-empty format string
                if (*pszFormat != L'\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchNewDestLength = 0;

                hr = StringVPrintf_lWorkerW(pszDest,
                                            cchDest,
                                            &cchNewDestLength,
                                            pszFormat,
                                            locale,
                                            argList);

                pszDestEnd = pszDest + cchNewDestLength;
                cchRemaining = cchDest - cchNewDestLength;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                    cbRemaining = cchRemaining * sizeof(wchar_t);

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullW(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = L'\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cchDest != 0))
        {
            size_t cbDest;

            // safe to multiply cchDest * sizeof(wchar_t) since cchDest < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
            cbDest = cchDest * sizeof(wchar_t);

            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsW(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcchRemaining)
            {
                *pcchRemaining = cchRemaining;
            }
        }
    }

    return hr;
}
#endif  // STRSAFE_LOCALE_FUNCTIONS && !STRSAFE_NO_CCH_FUNCTIONS


#ifndef STRSAFE_NO_CB_FUNCTIONS
/*++

STDAPI
StringCbVPrintfEx(
    __out_bcount(cbDest) LPTSTR  pszDest         OPTIONAL,
    __in size_t  cbDest,
    __deref_opt_out_bcount(*pcbRemaining) LPTSTR* ppszDestEnd     OPTIONAL,
    __out_opt size_t* pcbRemaining    OPTIONAL,
    __in DWORD   dwFlags,
    __in __format_string LPCTSTR pszFormat       OPTIONAL,
    __in va_list argList
    );

Routine Description:

    This routine is a safer version of the C built-in function 'vsprintf' with
    some additional parameters.  In addition to functionality provided by
    StringCbVPrintf, this routine also returns a pointer to the end of the
    destination string and the number of characters left in the destination string
    including the null terminator. The flags parameter allows additional controls.

Arguments:

    pszDest         -   destination string

    cbDest          -   size of destination buffer in bytes.
                        length must be sufficient to contain the resulting
                        formatted string plus the null terminator.

    ppszDestEnd     -   if ppszDestEnd is non-null, the function will return
                        a pointer to the end of the destination string.  If the
                        function printed any data, the result will point to the
                        null termination character

    pcbRemaining    -   if pcbRemaining is non-null, the function will return
                        the number of bytes left in the destination string,
                        including the null terminator

    dwFlags         -   controls some details of the string copy:

        STRSAFE_FILL_BEHIND_NULL
                    if the function succeeds, the low byte of dwFlags will be
                    used to fill the uninitialize part of destination buffer
                    behind the null terminator

        STRSAFE_IGNORE_NULLS
                    treat NULL string pointers like empty strings (TEXT(""))

        STRSAFE_FILL_ON_FAILURE
                    if the function fails, the low byte of dwFlags will be
                    used to fill all of the destination buffer, and it will
                    be null terminated. This will overwrite any truncated
                    string returned when the failure is
                    STRSAFE_E_INSUFFICIENT_BUFFER

        STRSAFE_NO_TRUNCATION /
        STRSAFE_NULL_ON_FAILURE
                    if the function fails, the destination buffer will be set
                    to the empty string. This will overwrite any truncated string
                    returned when the failure is STRSAFE_E_INSUFFICIENT_BUFFER.

    pszFormat       -   format string which must be null terminated

    argList         -   va_list from the variable arguments according to the
                        stdarg.h convention

Notes:
    Behavior is undefined if destination, format strings or any arguments
    strings overlap.

    pszDest and pszFormat should not be NULL unless the STRSAFE_IGNORE_NULLS
    flag is specified.  If STRSAFE_IGNORE_NULLS is passed, both pszDest and
    pszFormat may be NULL.  An error may still be returned even though NULLS
    are ignored due to insufficient space.

Return Value:

    S_OK           -   if there was source data and it was all concatenated and
                       the resultant dest string was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that the print
                       operation failed due to insufficient space. When this
                       error occurs, the destination buffer is modified to
                       contain a truncated version of the ideal result and is
                       null terminated. This is useful for situations where
                       truncation is ok.

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function

--*/
#ifdef UNICODE
#define StringCbVPrintfEx  StringCbVPrintfExW
#else
#define StringCbVPrintfEx  StringCbVPrintfExA
#endif // !UNICODE

STRSAFEAPI
StringCbVPrintfExA(
    __out_bcount(cbDest) STRSAFE_LPSTR pszDest,
    __in size_t cbDest,
    __deref_opt_out_bcount(*pcbRemaining) STRSAFE_LPSTR* ppszDestEnd,
    __out_opt size_t* pcbRemaining,
    __in DWORD dwFlags,
    __in __format_string STRSAFE_LPCSTR pszFormat,
    __in va_list argList)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(char);

    hr = StringExValidateDestA(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcA(&pszFormat, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = '\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually a non-empty format string
                if (*pszFormat != '\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchNewDestLength = 0;

                hr = StringVPrintfWorkerA(pszDest,
                                          cchDest,
                                          &cchNewDestLength,
                                          pszFormat,
                                          argList);

                pszDestEnd = pszDest + cchNewDestLength;
                cchRemaining = cchDest - cchNewDestLength;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                    cbRemaining = (cchRemaining * sizeof(char)) + (cbDest % sizeof(char));

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullA(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = '\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cbDest != 0))
        {
            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsA(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcbRemaining)
            {
                // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                *pcbRemaining = (cchRemaining * sizeof(char)) + (cbDest % sizeof(char));
            }
        }
    }

    return hr;
}

STRSAFEAPI
StringCbVPrintfExW(
    __out_bcount(cbDest) STRSAFE_LPWSTR pszDest,
    __in size_t cbDest,
    __deref_opt_out_bcount(*pcbRemaining) STRSAFE_LPWSTR* ppszDestEnd,
    __out_opt size_t* pcbRemaining,
    __in DWORD dwFlags,
    __in __format_string STRSAFE_LPCWSTR pszFormat,
    __in va_list argList)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(wchar_t);

    hr = StringExValidateDestW(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPWSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcW(&pszFormat, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = L'\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually a non-empty format string
                if (*pszFormat != L'\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchNewDestLength = 0;

                hr = StringVPrintfWorkerW(pszDest,
                                          cchDest,
                                          &cchNewDestLength,
                                          pszFormat,
                                          argList);

                pszDestEnd = pszDest + cchNewDestLength;
                cchRemaining = cchDest - cchNewDestLength;

                if (SUCCEEDED(hr) && (dwFlags & STRSAFE_FILL_BEHIND_NULL))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                    cbRemaining = (cchRemaining * sizeof(wchar_t)) + (cbDest % sizeof(wchar_t));

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullW(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = L'\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cbDest != 0))
        {
            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsW(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcbRemaining)
            {
                // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                *pcbRemaining = (cchRemaining * sizeof(wchar_t)) + (cbDest % sizeof(wchar_t));
            }
        }
    }

    return hr;
}
#endif  // !STRSAFE_NO_CB_FUNCTIONS


#if defined(STRSAFE_LOCALE_FUNCTIONS) && !defined(STRSAFE_NO_CB_FUNCTIONS)
/*++

STDAPI
StringCbVPrintf_lEx(
    __out_bcount(cbDest) LPTSTR  pszDest         OPTIONAL,
    __in size_t  cbDest,
    __deref_opt_out_bcount(*pcbRemaining) LPTSTR* ppszDestEnd     OPTIONAL,
    __out_opt size_t* pcbRemaining    OPTIONAL,
    __in DWORD   dwFlags,
    __in __format_string LPCTSTR pszFormat       OPTIONAL,
    __in locale_t locale,
    __in va_list argList
    );

Routine Description:

    This routine is a version of StringCbVPrintfEx that also takes a locale.
    Please see notes for StringCbVPrintfEx above.

--*/
#ifdef UNICODE
#define StringCbVPrintf_lEx StringCbVPrintf_lExW
#else
#define StringCbVPrintf_lEx StringCbVPrintf_lExA
#endif // !UNICODE

STRSAFEAPI
StringCbVPrintf_lExA(
    __out_bcount(cbDest) STRSAFE_LPSTR pszDest,
    __in size_t cbDest,
    __deref_opt_out_bcount(*pcbRemaining) STRSAFE_LPSTR* ppszDestEnd,
    __out_opt size_t* pcbRemaining,
    __in DWORD dwFlags,
    __in __format_string STRSAFE_LPCSTR pszFormat,
    __in _locale_t locale,
    __in va_list argList)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(char);

    hr = StringExValidateDestA(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcA(&pszFormat, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = '\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually a non-empty format string
                if (*pszFormat != '\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchNewDestLength = 0;

                hr = StringVPrintf_lWorkerA(pszDest,
                                            cchDest,
                                            &cchNewDestLength,
                                            pszFormat,
                                            locale,
                                            argList);

                pszDestEnd = pszDest + cchNewDestLength;
                cchRemaining = cchDest - cchNewDestLength;

                if (SUCCEEDED(hr)                           &&
                    (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                    (cchRemaining > 1))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                    cbRemaining = (cchRemaining * sizeof(char)) + (cbDest % sizeof(char));

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullA(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = '\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cbDest != 0))
        {
            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsA(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcbRemaining)
            {
                // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                *pcbRemaining = (cchRemaining * sizeof(char)) + (cbDest % sizeof(char));
            }
        }
    }

    return hr;
}

STRSAFEAPI
StringCbVPrintf_lExW(
    __out_bcount(cbDest) STRSAFE_LPWSTR pszDest,
    __in size_t cbDest,
    __deref_opt_out_bcount(*pcbRemaining) STRSAFE_LPWSTR* ppszDestEnd,
    __out_opt size_t* pcbRemaining,
    __in DWORD dwFlags,
    __in __format_string STRSAFE_LPCWSTR pszFormat,
    __in _locale_t locale,
    __in va_list argList)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(wchar_t);

    hr = StringExValidateDestW(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPWSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        hr = StringExValidateSrcW(&pszFormat, NULL, STRSAFE_MAX_CCH, dwFlags);

        if (SUCCEEDED(hr))
        {
            if (dwFlags & (~STRSAFE_VALID_FLAGS))
            {
                hr = STRSAFE_E_INVALID_PARAMETER;

                if (cchDest != 0)
                {
                    *pszDest = L'\0';
                }
            }
            else if (cchDest == 0)
            {
                // only fail if there was actually a non-empty format string
                if (*pszFormat != L'\0')
                {
                    if (pszDest == NULL)
                    {
                        hr = STRSAFE_E_INVALID_PARAMETER;
                    }
                    else
                    {
                        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
                    }
                }
            }
            else
            {
                size_t cchNewDestLength = 0;

                hr = StringVPrintf_lWorkerW(pszDest,
                                            cchDest,
                                            &cchNewDestLength,
                                            pszFormat,
                                            locale,
                                            argList);

                pszDestEnd = pszDest + cchNewDestLength;
                cchRemaining = cchDest - cchNewDestLength;

                if (SUCCEEDED(hr) && (dwFlags & STRSAFE_FILL_BEHIND_NULL))
                {
                    size_t cbRemaining;

                    // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                    cbRemaining = (cchRemaining * sizeof(wchar_t)) + (cbDest % sizeof(wchar_t));

                    // handle the STRSAFE_FILL_BEHIND_NULL flag
                    StringExHandleFillBehindNullW(pszDestEnd, cbRemaining, dwFlags);
                }
            }
        }
        else
        {
            if (cchDest != 0)
            {
                *pszDest = L'\0';
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cbDest != 0))
        {
            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsW(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr) || (hr == STRSAFE_E_INSUFFICIENT_BUFFER))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcbRemaining)
            {
                // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                *pcbRemaining = (cchRemaining * sizeof(wchar_t)) + (cbDest % sizeof(wchar_t));
            }
        }
    }

    return hr;
}
#endif  // STRSAFE_LOCALE_FUNCTIONS && !STRSAFE_NO_CB_FUNCTIONS


#ifndef STRSAFE_LIB_IMPL

#ifndef STRSAFE_NO_CCH_FUNCTIONS
/*++

STDAPI
StringCchGets(
    __out_ecount(cchDest) LPTSTR  pszDest,
    __in size_t  cchDest
    );

Routine Description:

    This routine is a safer version of the C built-in function 'gets'.
    The size of the destination buffer (in characters) is a parameter and
    this function will not write past the end of this buffer and it will
    ALWAYS null terminate the destination buffer (unless it is zero length).

    This routine is not a replacement for fgets.  That function does not replace
    newline characters with a null terminator.

    This function returns a hresult, and not a pointer.  It returns
    S_OK if any characters were read from stdin and copied to pszDest and
    pszDest was null terminated, otherwise it will return a failure code.

Arguments:

    pszDest     -   destination string

    cchDest     -   size of destination buffer in characters.

Notes:
    pszDest should not be NULL. See StringCchGetsEx if you require the handling
    of NULL values.

    cchDest must be > 1 for this function to succeed.

Return Value:

    S_OK           -   data was read from stdin and copied, and the resultant
                       dest string was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_END_OF_FILE /
      HRESULT_CODE(hr) == ERROR_HANDLE_EOF
                   -   this return value indicates an error or end-of-file
                       condition, use feof or ferror to determine which one has
                       occured.

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that there was
                       insufficient space in the destination buffer to copy any
                       data

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function.

--*/
#ifdef UNICODE
#define StringCchGets  StringCchGetsW
#else
#define StringCchGets  StringCchGetsA
#endif // !UNICODE

STRSAFEAPI
StringCchGetsA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest)
{
    HRESULT hr;

    hr = StringValidateDestA(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        hr = StringGetsWorkerA(pszDest, cchDest, NULL);
    }

    return hr;
}

STRSAFEAPI
StringCchGetsW(
    __out_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest)
{
    HRESULT hr;

    hr = StringValidateDestW(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        hr = StringGetsWorkerW(pszDest, cchDest, NULL);
    }

    return hr;
}
#endif  // !STRSAFE_NO_CCH_FUNCTIONS


#ifndef STRSAFE_NO_CB_FUNCTIONS
/*++

STDAPI
StringCbGets(
    __out_bcount(cbDest) LPTSTR  pszDest,
    __in size_t  cbDest
    );

Routine Description:

    This routine is a safer version of the C built-in function 'gets'.
    The size of the destination buffer (in bytes) is a parameter and
    this function will not write past the end of this buffer and it will
    ALWAYS null terminate the destination buffer (unless it is zero length).

    This routine is not a replacement for fgets.  That function does not replace
    newline characters with a null terminator.

    This function returns a hresult, and not a pointer.  It returns
    S_OK if any characters were read from stdin and copied to pszDest
    and pszDest was null terminated, otherwise it will return a failure code.

Arguments:

    pszDest     -   destination string

    cbDest      -   size of destination buffer in bytes.

Notes:
    pszDest should not be NULL. See StringCbGetsEx if you require the handling
    of NULL values.

    cbDest must be > sizeof(TCHAR) for this function to succeed.

Return Value:

    S_OK           -   data was read from stdin and copied, and the resultant
                       dest string was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_END_OF_FILE /
      HRESULT_CODE(hr) == ERROR_HANDLE_EOF
                   -   this return value indicates an error or end-of-file
                       condition, use feof or ferror to determine which one has
                       occured.

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that there was
                       insufficient space in the destination buffer to copy any
                       data

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function.

--*/
#ifdef UNICODE
#define StringCbGets  StringCbGetsW
#else
#define StringCbGets  StringCbGetsA
#endif // !UNICODE

STRSAFEAPI
StringCbGetsA(
    __out_bcount(cbDest) STRSAFE_LPSTR pszDest,
    __in size_t cbDest)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(char);

    hr = StringValidateDestA(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        hr = StringGetsWorkerA(pszDest, cchDest, NULL);
    }

    return hr;
}

STRSAFEAPI
StringCbGetsW(
    __out_bcount(cbDest) STRSAFE_LPWSTR pszDest,
    __in size_t cbDest)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(wchar_t);

    hr = StringValidateDestW(pszDest, cchDest, NULL, STRSAFE_MAX_CCH);

    if (SUCCEEDED(hr))
    {
        hr = StringGetsWorkerW(pszDest, cchDest, NULL);
    }

    return hr;
}
#endif  // !STRSAFE_NO_CB_FUNCTIONS


#ifndef STRSAFE_NO_CCH_FUNCTIONS
/*++

STDAPI
StringCchGetsEx(
    __out_ecount(cchDest) LPTSTR  pszDest         OPTIONAL,
    __in size_t  cchDest,
    __deref_opt_out_ecount(*pcchRemaining) LPTSTR* ppszDestEnd     OPTIONAL,
    __out_opt size_t* pcchRemaining   OPTIONAL,
    __in DWORD   dwFlags
    );

Routine Description:

    This routine is a safer version of the C built-in function 'gets' with
    some additional parameters. In addition to functionality provided by
    StringCchGets, this routine also returns a pointer to the end of the
    destination string and the number of characters left in the destination string
    including the null terminator. The flags parameter allows additional controls.

Arguments:

    pszDest         -   destination string

    cchDest         -   size of destination buffer in characters.

    ppszDestEnd     -   if ppszDestEnd is non-null, the function will return a
                        pointer to the end of the destination string.  If the
                        function copied any data, the result will point to the
                        null termination character

    pcchRemaining   -   if pcchRemaining is non-null, the function will return the
                        number of characters left in the destination string,
                        including the null terminator

    dwFlags         -   controls some details of the string copy:

        STRSAFE_FILL_BEHIND_NULL
                    if the function succeeds, the low byte of dwFlags will be
                    used to fill the uninitialize part of destination buffer
                    behind the null terminator

        STRSAFE_IGNORE_NULLS
                    treat NULL string pointers like empty strings (TEXT("")).

        STRSAFE_FILL_ON_FAILURE
                    if the function fails, the low byte of dwFlags will be
                    used to fill all of the destination buffer, and it will
                    be null terminated.

        STRSAFE_NO_TRUNCATION /
        STRSAFE_NULL_ON_FAILURE
                    if the function fails, the destination buffer will be set
                    to the empty string.

Notes:
    pszDest should not be NULL unless the STRSAFE_IGNORE_NULLS flag is specified.
    If STRSAFE_IGNORE_NULLS is passed and pszDest is NULL, an error may still be
    returned even though NULLS are ignored

    cchDest must be > 1 for this function to succeed.

Return Value:

    S_OK           -   data was read from stdin and copied, and the resultant
                       dest string was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_END_OF_FILE /
      HRESULT_CODE(hr) == ERROR_HANDLE_EOF
                   -   this return value indicates an error or end-of-file
                       condition, use feof or ferror to determine which one has
                       occured.

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that there was
                       insufficient space in the destination buffer to copy any
                       data

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function.

--*/
#ifdef UNICODE
#define StringCchGetsEx  StringCchGetsExW
#else
#define StringCchGetsEx  StringCchGetsExA
#endif // !UNICODE

STRSAFEAPI
StringCchGetsExA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __deref_opt_out_ecount(*pcchRemaining) STRSAFE_LPSTR* ppszDestEnd,
    __out_opt size_t* pcchRemaining,
    __in DWORD dwFlags)
{
    HRESULT hr;

    hr = StringExValidateDestA(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        if (dwFlags & (~STRSAFE_VALID_FLAGS))
        {
            hr = STRSAFE_E_INVALID_PARAMETER;

            if (cchDest != 0)
            {
                *pszDest = '\0';
            }
        }
        else if (cchDest == 0)
        {
            if (pszDest == NULL)
            {
                hr = STRSAFE_E_INVALID_PARAMETER;
            }
            else
            {
                hr = STRSAFE_E_INSUFFICIENT_BUFFER;
            }
        }
        else
        {
            size_t cchNewDestLength = 0;

            hr = StringGetsWorkerA(pszDest, cchDest, &cchNewDestLength);

            pszDestEnd = pszDest + cchNewDestLength;
            cchRemaining = cchDest - cchNewDestLength;

            if (SUCCEEDED(hr)                           &&
                (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                (cchRemaining > 1))
            {
                size_t cbRemaining;

                // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                cbRemaining = cchRemaining * sizeof(char);

                // handle the STRSAFE_FILL_BEHIND_NULL flag
                StringExHandleFillBehindNullA(pszDestEnd, cbRemaining, dwFlags);
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cchDest != 0))
        {
            size_t cbDest;

            // safe to multiply cchDest * sizeof(char) since cchDest < STRSAFE_MAX_CCH and sizeof(char) is 1
            cbDest = cchDest * sizeof(char);

            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsA(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr)                           ||
            (hr == STRSAFE_E_INSUFFICIENT_BUFFER)   ||
            (hr == STRSAFE_E_END_OF_FILE))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcchRemaining)
            {
                *pcchRemaining = cchRemaining;
            }
        }
    }

    return hr;
}

STRSAFEAPI
StringCchGetsExW(
    __out_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __deref_opt_out_ecount(*pcchRemaining) STRSAFE_LPWSTR* ppszDestEnd,
    __out_opt size_t* pcchRemaining,
    __in DWORD dwFlags)
{
    HRESULT hr;

    hr = StringExValidateDestW(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPWSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        if (dwFlags & (~STRSAFE_VALID_FLAGS))
        {
            hr = STRSAFE_E_INVALID_PARAMETER;

            if (cchDest != 0)
            {
                *pszDest = L'\0';
            }
        }
        else if (cchDest == 0)
        {
            if (pszDest == NULL)
            {
                hr = STRSAFE_E_INVALID_PARAMETER;
            }
            else
            {
                hr = STRSAFE_E_INSUFFICIENT_BUFFER;
            }
        }
        else
        {
            size_t cchNewDestLength = 0;

            hr = StringGetsWorkerW(pszDest, cchDest, &cchNewDestLength);

            pszDestEnd = pszDest + cchNewDestLength;
            cchRemaining = cchDest - cchNewDestLength;

            if (SUCCEEDED(hr)                           &&
                (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                (cchRemaining > 1))
            {
                size_t cbRemaining;

                // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                cbRemaining = cchRemaining * sizeof(wchar_t);

                // handle the STRSAFE_FILL_BEHIND_NULL flag
                StringExHandleFillBehindNullW(pszDestEnd, cbRemaining, dwFlags);
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cchDest != 0))
        {
            size_t cbDest;

            // safe to multiply cchDest * sizeof(wchar_t) since cchDest < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
            cbDest = cchDest * sizeof(wchar_t);

            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsW(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr)                           ||
            (hr == STRSAFE_E_INSUFFICIENT_BUFFER)   ||
            (hr == STRSAFE_E_END_OF_FILE))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcchRemaining)
            {
                *pcchRemaining = cchRemaining;
            }
        }
    }

    return hr;
}
#endif  // !STRSAFE_NO_CCH_FUNCTIONS


#ifndef STRSAFE_NO_CB_FUNCTIONS
/*++

STDAPI
StringCbGetsEx(
    __out_bcount(cbDest) LPTSTR  pszDest         OPTIONAL,
    __in size_t  cbDest,
    __deref_opt_out_bcount(*pcbRemaining) LPTSTR* ppszDestEnd     OPTIONAL,
    __out_opt size_t* pcbRemaining    OPTIONAL,
    __in DWORD   dwFlags
    );

Routine Description:

    This routine is a safer version of the C built-in function 'gets' with
    some additional parameters. In addition to functionality provided by
    StringCbGets, this routine also returns a pointer to the end of the
    destination string and the number of characters left in the destination string
    including the null terminator. The flags parameter allows additional controls.

Arguments:

    pszDest         -   destination string

    cbDest          -   size of destination buffer in bytes.

    ppszDestEnd     -   if ppszDestEnd is non-null, the function will return a
                        pointer to the end of the destination string.  If the
                        function copied any data, the result will point to the
                        null termination character

    pcbRemaining    -   if pcbRemaining is non-null, the function will return the
                        number of bytes left in the destination string,
                        including the null terminator

    dwFlags         -   controls some details of the string copy:

        STRSAFE_FILL_BEHIND_NULL
                    if the function succeeds, the low byte of dwFlags will be
                    used to fill the uninitialize part of destination buffer
                    behind the null terminator

        STRSAFE_IGNORE_NULLS
                    treat NULL string pointers like empty strings (TEXT("")).

        STRSAFE_FILL_ON_FAILURE
                    if the function fails, the low byte of dwFlags will be
                    used to fill all of the destination buffer, and it will
                    be null terminated.

        STRSAFE_NO_TRUNCATION /
        STRSAFE_NULL_ON_FAILURE
                    if the function fails, the destination buffer will be set
                    to the empty string.

Notes:
    pszDest should not be NULL unless the STRSAFE_IGNORE_NULLS flag is specified.
    If STRSAFE_IGNORE_NULLS is passed and pszDest is NULL, an error may still be
    returned even though NULLS are ignored

    cbDest must be > sizeof(TCHAR) for this function to succeed

Return Value:

    S_OK           -   data was read from stdin and copied, and the resultant
                       dest string was null terminated

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

      STRSAFE_E_END_OF_FILE /
      HRESULT_CODE(hr) == ERROR_HANDLE_EOF
                   -   this return value indicates an error or end-of-file
                       condition, use feof or ferror to determine which one has
                       occured.

      STRSAFE_E_INSUFFICIENT_BUFFER /
      HRESULT_CODE(hr) == ERROR_INSUFFICIENT_BUFFER
                   -   this return value is an indication that there was
                       insufficient space in the destination buffer to copy any
                       data

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function.

--*/
#ifdef UNICODE
#define StringCbGetsEx  StringCbGetsExW
#else
#define StringCbGetsEx  StringCbGetsExA
#endif // !UNICODE

STRSAFEAPI
StringCbGetsExA(
    __out_bcount(cbDest) STRSAFE_LPSTR pszDest,
    __in size_t cbDest,
    __deref_opt_out_bcount(*pcbRemaining) STRSAFE_LPSTR* ppszDestEnd,
    __out_opt size_t* pcbRemaining,
    __in DWORD dwFlags)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(char);

    hr = StringExValidateDestA(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        if (dwFlags & (~STRSAFE_VALID_FLAGS))
        {
            hr = STRSAFE_E_INVALID_PARAMETER;

            if (cchDest != 0)
            {
                *pszDest = '\0';
            }
        }
        else if (cchDest == 0)
        {
            if (pszDest == NULL)
            {
                hr = STRSAFE_E_INVALID_PARAMETER;
            }
            else
            {
                hr = STRSAFE_E_INSUFFICIENT_BUFFER;
            }
        }
        else
        {
            size_t cchNewDestLength = 0;

            hr = StringGetsWorkerA(pszDest, cchDest, &cchNewDestLength);

            pszDestEnd = pszDest + cchNewDestLength;
            cchRemaining = cchDest - cchNewDestLength;

            if (SUCCEEDED(hr)                           &&
                (dwFlags & STRSAFE_FILL_BEHIND_NULL)    &&
                (cchRemaining > 1))
            {
                size_t cbRemaining;

                // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                cbRemaining = (cchRemaining * sizeof(char)) + (cbDest % sizeof(char));

                // handle the STRSAFE_FILL_BEHIND_NULL flag
                StringExHandleFillBehindNullA(pszDestEnd, cbRemaining, dwFlags);
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cbDest != 0))
        {
            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsA(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr)                           ||
            (hr == STRSAFE_E_INSUFFICIENT_BUFFER)   ||
            (hr == STRSAFE_E_END_OF_FILE))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcbRemaining)
            {
                // safe to multiply cchRemaining * sizeof(char) since cchRemaining < STRSAFE_MAX_CCH and sizeof(char) is 1
                *pcbRemaining = (cchRemaining * sizeof(char)) + (cbDest % sizeof(char));
            }
        }
    }

    return hr;
}

STRSAFEAPI
StringCbGetsExW(
    __out_bcount(cbDest) STRSAFE_LPWSTR pszDest,
    __in size_t cbDest,
    __deref_opt_out_bcount(*pcbRemaining) STRSAFE_LPWSTR* ppszDestEnd,
    __out_opt size_t* pcbRemaining,
    __in DWORD dwFlags)
{
    HRESULT hr;
    size_t cchDest = cbDest / sizeof(wchar_t);

    hr = StringExValidateDestW(&pszDest,
                               &cchDest,
                               NULL,
                               STRSAFE_MAX_CCH,
                               dwFlags);

    if (SUCCEEDED(hr))
    {
        STRSAFE_LPWSTR pszDestEnd = pszDest;
        size_t cchRemaining = cchDest;

        if (dwFlags & (~STRSAFE_VALID_FLAGS))
        {
            hr = STRSAFE_E_INVALID_PARAMETER;

            if (cchDest != 0)
            {
                *pszDest = L'\0';
            }
        }
        else if (cchDest == 0)
        {
            if (pszDest == NULL)
            {
                hr = STRSAFE_E_INVALID_PARAMETER;
            }
            else
            {
                hr = STRSAFE_E_INSUFFICIENT_BUFFER;
            }
        }
        else
        {
            size_t cchNewDestLength = 0;

            hr = StringGetsWorkerW(pszDest, cchDest, &cchNewDestLength);

            pszDestEnd = pszDest + cchNewDestLength;
            cchRemaining = cchDest - cchNewDestLength;

            if (SUCCEEDED(hr) && (dwFlags & STRSAFE_FILL_BEHIND_NULL))
            {
                size_t cbRemaining;

                // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                cbRemaining = (cchRemaining * sizeof(wchar_t)) + (cbDest % sizeof(wchar_t));

                // handle the STRSAFE_FILL_BEHIND_NULL flag
                StringExHandleFillBehindNullW(pszDestEnd, cbRemaining, dwFlags);
            }
        }

        if (FAILED(hr)                                                                              &&
            (dwFlags & (STRSAFE_NO_TRUNCATION | STRSAFE_FILL_ON_FAILURE | STRSAFE_NULL_ON_FAILURE)) &&
            (cbDest != 0))
        {
            // handle the STRSAFE_FILL_ON_FAILURE, STRSAFE_NULL_ON_FAILURE, and STRSAFE_NO_TRUNCATION flags
            StringExHandleOtherFlagsW(pszDest,
                                      cbDest,
                                      0,
                                      &pszDestEnd,
                                      &cchRemaining,
                                      dwFlags);
        }

        if (SUCCEEDED(hr)                           ||
            (hr == STRSAFE_E_INSUFFICIENT_BUFFER)   ||
            (hr == STRSAFE_E_END_OF_FILE))
        {
            if (ppszDestEnd)
            {
                *ppszDestEnd = pszDestEnd;
            }

            if (pcbRemaining)
            {
                // safe to multiply cchRemaining * sizeof(wchar_t) since cchRemaining < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
                *pcbRemaining = (cchRemaining * sizeof(wchar_t)) + (cbDest % sizeof(wchar_t));
            }
        }
    }

    return hr;
}
#endif  // !STRSAFE_NO_CB_FUNCTIONS

#endif  // !STRSAFE_LIB_IMPL

#ifndef STRSAFE_NO_CCH_FUNCTIONS
/*++

STDAPI
StringCchLength(
    __in_opt LPCTSTR psz         OPTIONAL,
    __in size_t  cchMax,
    __out_opt size_t* pcchLength  OPTIONAL
    );

Routine Description:

    This routine is a safer version of the C built-in function 'strlen'.
    It is used to make sure a string is not larger than a given length, and
    it optionally returns the current length in characters not including
    the null terminator.

    This function returns a hresult, and not a pointer.  It returns
    S_OK if the string is non-null and the length including the null
    terminator is less than or equal to cchMax characters.

Arguments:

    psz         -   string to check the length of

    cchMax      -   maximum number of characters including the null terminator
                    that psz is allowed to contain

    pcch        -   if the function succeeds and pcch is non-null, the current length
                    in characters of psz excluding the null terminator will be returned.
                    This out parameter is equivalent to the return value of strlen(psz)

Notes:
    psz can be null but the function will fail

    cchMax should be greater than zero or the function will fail

Return Value:

    S_OK           -   psz is non-null and the length including the null
                       terminator is less than or equal to cchMax characters

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function.

--*/
#ifdef UNICODE
#define StringCchLength  StringCchLengthW
#else
#define StringCchLength  StringCchLengthA
#endif // !UNICODE

STRSAFEAPI
StringCchLengthA(
    __in_opt STRSAFE_LPCSTR psz,
    __in size_t cchMax,
    __out_opt size_t* pcchLength)
{
    HRESULT hr;

    if ((psz == NULL) || (cchMax > STRSAFE_MAX_CCH))
    {
        hr = STRSAFE_E_INVALID_PARAMETER;
    }
    else
    {
        hr = StringLengthWorkerA(psz, cchMax, pcchLength);
    }

    if (FAILED(hr) && pcchLength)
    {
        *pcchLength = 0;
    }

    return hr;
}

STRSAFEAPI
StringCchLengthW(
    __in_opt STRSAFE_LPCWSTR psz,
    __in size_t cchMax,
    __out_opt size_t* pcchLength)
{
    HRESULT hr;

    if ((psz == NULL) || (cchMax > STRSAFE_MAX_CCH))
    {
        hr = STRSAFE_E_INVALID_PARAMETER;
    }
    else
    {
        hr = StringLengthWorkerW(psz, cchMax, pcchLength);
    }

    if (FAILED(hr) && pcchLength)
    {
        *pcchLength = 0;
    }

    return hr;
}
#endif  // !STRSAFE_NO_CCH_FUNCTIONS


#ifndef STRSAFE_NO_CB_FUNCTIONS
/*++

STDAPI
StringCbLength(
    __in_opt LPCTSTR psz,        OPTIONAL
    __in size_t  cbMax,
    __out_opt size_t* pcbLength   OPTIONAL
    );

Routine Description:

    This routine is a safer version of the C built-in function 'strlen'.
    It is used to make sure a string is not larger than a given length, and
    it optionally returns the current length in bytes not including
    the null terminator.

    This function returns a hresult, and not a pointer.  It returns
    S_OK if the string is non-null and the length including the null
    terminator is less than or equal to cbMax bytes.

Arguments:

    psz         -   string to check the length of

    cbMax       -   maximum number of bytes including the null terminator
                    that psz is allowed to contain

    pcb         -   if the function succeeds and pcb is non-null, the current length
                    in bytes of psz excluding the null terminator will be returned.
                    This out parameter is equivalent to the return value of strlen(psz) * sizeof(TCHAR)

Notes:
    psz can be null but the function will fail

    cbMax should be greater than or equal to sizeof(TCHAR) or the function will fail

Return Value:

    S_OK           -   psz is non-null and the length including the null
                       terminator is less than or equal to cbMax bytes

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function.

--*/
#ifdef UNICODE
#define StringCbLength  StringCbLengthW
#else
#define StringCbLength  StringCbLengthA
#endif // !UNICODE

STRSAFEAPI
StringCbLengthA(
    __in_opt STRSAFE_LPCSTR psz,
    __in size_t cbMax,
    __out_opt size_t* pcbLength)
{
    HRESULT hr;
    size_t cchMax = cbMax / sizeof(char);
    size_t cchLength = 0;

    if ((psz == NULL) || (cchMax > STRSAFE_MAX_CCH))
    {
        hr = STRSAFE_E_INVALID_PARAMETER;
    }
    else
    {
        hr = StringLengthWorkerA(psz, cchMax, &cchLength);
    }

    if (pcbLength)
    {
        if (SUCCEEDED(hr))
        {
             // safe to multiply cchLength * sizeof(char) since cchLength < STRSAFE_MAX_CCH and sizeof(char) is 1
            *pcbLength = cchLength * sizeof(char);
        }
        else
        {
            *pcbLength = 0;
        }
    }

    return hr;
}

STRSAFEAPI
StringCbLengthW(
    __in_opt STRSAFE_LPCWSTR psz,
    __in size_t cbMax,
    __out_opt size_t* pcbLength)
{
    HRESULT hr;
    size_t cchMax = cbMax / sizeof(wchar_t);
    size_t cchLength = 0;

    if ((psz == NULL) || (cchMax > STRSAFE_MAX_CCH))
    {
        hr = STRSAFE_E_INVALID_PARAMETER;
    }
    else
    {
        hr = StringLengthWorkerW(psz, cchMax, &cchLength);
    }

    if (pcbLength)
    {
        if (SUCCEEDED(hr))
        {
            // safe to multiply cchLength * sizeof(wchar_t) since cchLength < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
            *pcbLength = cchLength * sizeof(wchar_t);
        }
        else
        {
            *pcbLength = 0;
        }
    }

    return hr;
}
#endif  // !STRSAFE_NO_CB_FUNCTIONS

#ifndef STRSAFE_NO_CCH_FUNCTIONS
/*++

STDAPI
UnalignedStringCchLength(
    __in_opt LPCUTSTR    psz         OPTIONAL,
    __in size_t      cchMax,
    __out_opt size_t*     pcchLength  OPTIONAL
    );

Routine Description:

    This routine is a version of StringCchLength that accepts an unaligned string pointer.

    This function returns a hresult, and not a pointer.  It returns
    S_OK if the string is non-null and the length including the null
    terminator is less than or equal to cchMax characters.

Arguments:

    psz         -   string to check the length of

    cchMax      -   maximum number of characters including the null terminator
                    that psz is allowed to contain

    pcch        -   if the function succeeds and pcch is non-null, the current length
                    in characters of psz excluding the null terminator will be returned.
                    This out parameter is equivalent to the return value of strlen(psz)

Notes:
    psz can be null but the function will fail

    cchMax should be greater than zero or the function will fail

Return Value:

    S_OK           -   psz is non-null and the length including the null
                       terminator is less than or equal to cchMax characters

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function.

--*/
#ifdef UNICODE
#define UnalignedStringCchLength    UnalignedStringCchLengthW
#else
#define UnalignedStringCchLength    StringCchLengthA
#endif // !UNICODE

#ifdef ALIGNMENT_MACHINE
STRSAFEAPI
UnalignedStringCchLengthW(
    __in_opt STRSAFE_LPCUWSTR psz,
    __in size_t cchMax,
    __out_opt size_t* pcchLength)
{
    HRESULT hr;

    if ((psz == NULL) || (cchMax > STRSAFE_MAX_CCH))
    {
        hr = STRSAFE_E_INVALID_PARAMETER;
    }
    else
    {
        hr = UnalignedStringLengthWorkerW(psz, cchMax, pcchLength);
    }

    if (FAILED(hr) && pcchLength)
    {
        *pcchLength = 0;
    }

    return hr;
}
#else
#define UnalignedStringCchLengthW   StringCchLengthW
#endif  // !ALIGNMENT_MACHINE
#endif  // !STRSAFE_NO_CCH_FUNCTIONS

#ifndef STRSAFE_NO_CB_FUNCTIONS
/*++

STDAPI
UnalignedStringCbLength(
    __in_opt LPCUTSTR    psz,        OPTIONAL
    __in size_t      cbMax,
    __out_opt size_t*     pcbLength   OPTIONAL
    );

Routine Description:

    This routine is a version of StringCbLength that accepts an unaligned string pointer.

    This function returns a hresult, and not a pointer.  It returns
    S_OK if the string is non-null and the length including the null
    terminator is less than or equal to cbMax bytes.

Arguments:

    psz         -   string to check the length of

    cbMax       -   maximum number of bytes including the null terminator
                    that psz is allowed to contain

    pcb         -   if the function succeeds and pcb is non-null, the current length
                    in bytes of psz excluding the null terminator will be returned.
                    This out parameter is equivalent to the return value of strlen(psz) * sizeof(TCHAR)

Notes:
    psz can be null but the function will fail

    cbMax should be greater than or equal to sizeof(TCHAR) or the function will fail

Return Value:

    S_OK           -   psz is non-null and the length including the null
                       terminator is less than or equal to cbMax bytes

    failure        -   you can use the macro HRESULT_CODE() to get a win32
                       error code for all hresult failure cases

    It is strongly recommended to use the SUCCEEDED() / FAILED() macros to test the
    return value of this function.

--*/
#ifdef UNICODE
#define UnalignedStringCbLength UnalignedStringCbLengthW
#else
#define UnalignedStringCbLength StringCbLengthA
#endif // !UNICODE

#ifdef ALIGNMENT_MACHINE
STRSAFEAPI
UnalignedStringCbLengthW(
    __in_opt STRSAFE_LPCUWSTR psz,
    __in size_t cbMax,
    __out_opt size_t* pcbLength)
{
    HRESULT hr;
    size_t cchMax = cbMax / sizeof(wchar_t);
    size_t cchLength = 0;

    if ((psz == NULL) || (cchMax > STRSAFE_MAX_CCH))
    {
        hr = STRSAFE_E_INVALID_PARAMETER;
    }
    else
    {
        hr = UnalignedStringLengthWorkerW(psz, cchMax, &cchLength);
    }

    if (pcbLength)
    {
        if (SUCCEEDED(hr))
        {
            // safe to multiply cchLength * sizeof(wchar_t) since cchLength < STRSAFE_MAX_CCH and sizeof(wchar_t) is 2
            *pcbLength = cchLength * sizeof(wchar_t);
        }
        else
        {
            *pcbLength = 0;
        }
    }

    return hr;
}
#else
#define UnalignedStringCbLengthW    StringCbLengthW
#endif  // !ALIGNMENT_MACHINE
#endif  // !STRSAFE_NO_CB_FUNCTIONS


#endif  // !STRSAFE_LIB_IMPL


// Below here are the worker functions that actually do the work

#if defined(STRSAFE_LIB_IMPL) || !defined(STRSAFE_LIB)

STRSAFEWORKERAPI
StringLengthWorkerA(
    __in STRSAFE_LPCSTR psz,
    __in size_t cchMax,
    __out_opt size_t* pcchLength)
{
    HRESULT hr = S_OK;
    size_t cchOriginalMax = cchMax;

    while (cchMax && (*psz != '\0'))
    {
        psz++;
        cchMax--;
    }

    if (cchMax == 0)
    {
        // the string is longer than cchMax
        hr = STRSAFE_E_INVALID_PARAMETER;
    }

    if (pcchLength)
    {
        if (SUCCEEDED(hr))
        {
            *pcchLength = cchOriginalMax - cchMax;
        }
        else
        {
            *pcchLength = 0;
        }
    }

    return hr;
}

STRSAFEWORKERAPI
StringLengthWorkerW(
    __in STRSAFE_LPCWSTR psz,
    __in size_t cchMax,
    __out_opt size_t* pcchLength)
{
    HRESULT hr = S_OK;
    size_t cchOriginalMax = cchMax;

    while (cchMax && (*psz != L'\0'))
    {
        psz++;
        cchMax--;
    }

    if (cchMax == 0)
    {
        // the string is longer than cchMax
        hr = STRSAFE_E_INVALID_PARAMETER;
    }

    if (pcchLength)
    {
        if (SUCCEEDED(hr))
        {
            *pcchLength = cchOriginalMax - cchMax;
        }
        else
        {
            *pcchLength = 0;
        }
    }

    return hr;
}

#ifdef ALIGNMENT_MACHINE
STRSAFEWORKERAPI
UnalignedStringLengthWorkerW(
    __in STRSAFE_LPCUWSTR psz,
    __in size_t cchMax,
    __out_opt size_t* pcchLength)
{
    HRESULT hr = S_OK;
    size_t cchOriginalMax = cchMax;

    while (cchMax && (*psz != L'\0'))
    {
        psz++;
        cchMax--;
    }

    if (cchMax == 0)
    {
        // the string is longer than cchMax
        hr = STRSAFE_E_INVALID_PARAMETER;
    }

    if (pcchLength)
    {
        if (SUCCEEDED(hr))
        {
            *pcchLength = cchOriginalMax - cchMax;
        }
        else
        {
            *pcchLength = 0;
        }
    }

    return hr;
}
#endif  // ALIGNMENT_MACHINE

STRSAFEWORKERAPI
StringExValidateSrcA(
    __deref_inout_opt STRSAFE_LPCSTR* ppszSrc,
    __inout_opt size_t* pcchToRead,
    __in size_t cchMax,
    __in DWORD dwFlags)
{
    HRESULT hr = S_OK;

    if (pcchToRead && (*pcchToRead >= cchMax))
    {
        hr = STRSAFE_E_INVALID_PARAMETER;
    }
    else if ((dwFlags & STRSAFE_IGNORE_NULLS) && (*ppszSrc == NULL))
    {
        *ppszSrc = "";

        if (pcchToRead)
        {
            *pcchToRead = 0;
        }
    }

    return hr;
}

STRSAFEWORKERAPI
StringExValidateSrcW(
    __deref_inout_opt STRSAFE_LPCWSTR* ppszSrc,
    __inout_opt size_t* pcchToRead,
    __in size_t cchMax,
    __in DWORD dwFlags)
{
    HRESULT hr = S_OK;

    if (pcchToRead && (*pcchToRead >= cchMax))
    {
        hr = STRSAFE_E_INVALID_PARAMETER;
    }
    else if ((dwFlags & STRSAFE_IGNORE_NULLS) && (*ppszSrc == NULL))
    {
        *ppszSrc = L"";

        if (pcchToRead)
        {
            *pcchToRead = 0;
        }
    }

    return hr;
}

STRSAFEWORKERAPI
StringValidateDestA(
    __in_ecount_opt(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __out_opt size_t* pcchDestLength,
    __in size_t cchMax)
{
    HRESULT hr = S_OK;

    if ((cchDest == 0) || (cchDest > cchMax))
    {
        hr = STRSAFE_E_INVALID_PARAMETER;
    }

    if (pcchDestLength)
    {
        if (SUCCEEDED(hr))
        {
            hr = StringLengthWorkerA(pszDest, cchDest, pcchDestLength);
        }
        else
        {
            *pcchDestLength = 0;
        }
    }

    return hr;
}

STRSAFEWORKERAPI
StringValidateDestW(
    __in_ecount_opt(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __out_opt size_t* pcchDestLength,
    __in size_t cchMax)
{
    HRESULT hr = S_OK;

    if ((cchDest == 0) || (cchDest > cchMax))
    {
        hr = STRSAFE_E_INVALID_PARAMETER;
    }

    if (pcchDestLength)
    {
        if (SUCCEEDED(hr))
        {
            hr = StringLengthWorkerW(pszDest, cchDest, pcchDestLength);
        }
        else
        {
            *pcchDestLength = 0;
        }
    }

    return hr;
}

STRSAFEWORKERAPI
StringExValidateDestA(
    __deref_inout_opt STRSAFE_LPSTR* ppszDest,
    __inout size_t* pcchDest,
    __out_opt size_t* pcchDestLength,
    __in size_t cchMax,
    __in DWORD dwFlags)
{
    HRESULT hr = S_OK;

    if (dwFlags & STRSAFE_IGNORE_NULLS)
    {
        if (((*ppszDest == NULL) && (*pcchDest != 0))   ||
            (*pcchDest > cchMax))
        {
            hr = STRSAFE_E_INVALID_PARAMETER;
        }

        if (pcchDestLength)
        {
            if (FAILED(hr) || (*pcchDest == 0))
            {
                *pcchDestLength = 0;
            }
            else
            {
                hr = StringLengthWorkerA(*ppszDest, *pcchDest, pcchDestLength);
            }
        }
    }
    else
    {
        hr = StringValidateDestA(*ppszDest, *pcchDest, pcchDestLength, cchMax);
    }

    return hr;
}

STRSAFEWORKERAPI
StringExValidateDestW(
    __deref_inout_opt STRSAFE_LPWSTR* ppszDest,
    __inout size_t* pcchDest,
    __out_opt size_t* pcchDestLength,
    __in size_t cchMax,
    __in DWORD dwFlags)
{
    HRESULT hr = S_OK;

    if (dwFlags & STRSAFE_IGNORE_NULLS)
    {
        if (((*ppszDest == NULL) && (*pcchDest != 0))   ||
            (*pcchDest > cchMax))
        {
            hr = STRSAFE_E_INVALID_PARAMETER;
        }

        if (pcchDestLength)
        {
            if (FAILED(hr) || (*pcchDest == 0))
            {
                *pcchDestLength = 0;
            }
            else
            {
                hr = StringLengthWorkerW(*ppszDest, *pcchDest, pcchDestLength);
            }
        }
    }
    else
    {
        hr = StringValidateDestW(*ppszDest, *pcchDest, pcchDestLength, cchMax);
    }

    return hr;
}

STRSAFEWORKERAPI
StringCopyWorkerA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __out_opt size_t* pcchNewDestLength,
    __in STRSAFE_LPCSTR pszSrc,
    __in size_t cchToCopy)
{
    HRESULT hr = S_OK;
    size_t cchNewDestLength = 0;

    // ASSERT(cchDest != 0);

    while (cchDest && cchToCopy && (*pszSrc != '\0'))
    {
        *pszDest++ = *pszSrc++;
        cchDest--;
        cchToCopy--;

        cchNewDestLength++;
    }

    if (cchDest == 0)
    {
        // we are going to truncate pszDest
        pszDest--;
        cchNewDestLength--;

        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
    }

    *pszDest= '\0';

    if (pcchNewDestLength)
    {
        *pcchNewDestLength = cchNewDestLength;
    }

    return hr;
}

STRSAFEWORKERAPI
StringCopyWorkerW(
    __out_ecount(cchDest) wchar_t* pszDest,
    __in size_t cchDest,
    __out_opt size_t* pcchNewDestLength,
    __in STRSAFE_LPCWSTR pszSrc,
    __in size_t cchToCopy)
{
    HRESULT hr = S_OK;
    size_t cchNewDestLength = 0;

    // ASSERT(cchDest != 0);

    while (cchDest && cchToCopy && (*pszSrc != L'\0'))
    {
        *pszDest++ = *pszSrc++;
        cchDest--;
        cchToCopy--;

        cchNewDestLength++;
    }

    if (cchDest == 0)
    {
        // we are going to truncate pszDest
        pszDest--;
        cchNewDestLength--;

        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
    }

    *pszDest= L'\0';

    if (pcchNewDestLength)
    {
        *pcchNewDestLength = cchNewDestLength;
    }

    return hr;
}

STRSAFEWORKERAPI
StringVPrintfWorkerA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __out_opt size_t* pcchNewDestLength,
    __in __format_string STRSAFE_LPCSTR pszFormat,
    __in va_list argList)
{
    HRESULT hr = S_OK;
    int iRet;
    size_t cchMax;
    size_t cchNewDestLength = 0;

    // leave the last space for the null terminator
    cchMax = cchDest - 1;

#if (STRSAFE_USE_SECURE_CRT == 1) && !defined(STRSAFE_LIB_IMPL)
    iRet = _vsnprintf_s(pszDest, cchDest, cchMax, pszFormat, argList);
#else
    iRet = _vsnprintf(pszDest, cchMax, pszFormat, argList);
#endif
    // ASSERT((iRet < 0) || (((size_t)iRet) <= cchMax));

    if ((iRet < 0) || (((size_t)iRet) > cchMax))
    {
        // need to null terminate the string
        pszDest += cchMax;
        *pszDest = '\0';

        cchNewDestLength = cchMax;

        // we have truncated pszDest
        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
    }
    else if (((size_t)iRet) == cchMax)
    {
        // need to null terminate the string
        pszDest += cchMax;
        *pszDest = '\0';

        cchNewDestLength = cchMax;
    }
    else
    {
        cchNewDestLength = (size_t)iRet;
    }

    if (pcchNewDestLength)
    {
        *pcchNewDestLength = cchNewDestLength;
    }

    return hr;
}

#ifdef STRSAFE_LOCALE_FUNCTIONS
STRSAFELOCALEWORKERAPI
StringVPrintf_lWorkerA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __out_opt size_t* pcchNewDestLength,
    __in __format_string STRSAFE_LPCSTR pszFormat,
    __in _locale_t locale,
    __in va_list argList)
{
    HRESULT hr = S_OK;
    int iRet;
    size_t cchMax;
    size_t cchNewDestLength = 0;

    // leave the last space for the null terminator
    cchMax = cchDest - 1;

#if (STRSAFE_USE_SECURE_CRT == 1) && !defined(STRSAFE_LIB_IMPL)
    iRet = _vsnprintf_s_l(pszDest, cchDest, cchMax, pszFormat, locale, argList);
#else
    iRet = _vsnprintf_l(pszDest, cchMax, pszFormat, locale, argList);
#endif
    // ASSERT((iRet < 0) || (((size_t)iRet) <= cchMax));

    if ((iRet < 0) || (((size_t)iRet) > cchMax))
    {
        // need to null terminate the string
        pszDest += cchMax;
        *pszDest = '\0';

        cchNewDestLength = cchMax;

        // we have truncated pszDest
        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
    }
    else if (((size_t)iRet) == cchMax)
    {
        // need to null terminate the string
        pszDest += cchMax;
        *pszDest = '\0';

        cchNewDestLength = cchMax;
    }
    else
    {
        cchNewDestLength = (size_t)iRet;
    }

    if (pcchNewDestLength)
    {
        *pcchNewDestLength = cchNewDestLength;
    }

    return hr;
}
#endif  // STRSAFE_LOCALE_FUNCTIONS

STRSAFEWORKERAPI
StringVPrintfWorkerW(
    __out_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __out_opt size_t* pcchNewDestLength,
    __in __format_string STRSAFE_LPCWSTR pszFormat,
    __in va_list argList)
{
    HRESULT hr = S_OK;
    int iRet;
    size_t cchMax;
    size_t cchNewDestLength = 0;

    // leave the last space for the null terminator
    cchMax = cchDest - 1;

#if (STRSAFE_USE_SECURE_CRT == 1) && !defined(STRSAFE_LIB_IMPL)
    iRet = _vsnwprintf_s(pszDest, cchDest, cchMax, pszFormat, argList);
#else
    iRet = _vsnwprintf(pszDest, cchMax, pszFormat, argList);
#endif
    // ASSERT((iRet < 0) || (((size_t)iRet) <= cchMax));

    if ((iRet < 0) || (((size_t)iRet) > cchMax))
    {
        // need to null terminate the string
        pszDest += cchMax;
        *pszDest = L'\0';

        cchNewDestLength = cchMax;

        // we have truncated pszDest
        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
    }
    else if (((size_t)iRet) == cchMax)
    {
        // need to null terminate the string
        pszDest += cchMax;
        *pszDest = L'\0';

        cchNewDestLength = cchMax;
    }
    else
    {
        cchNewDestLength = (size_t)iRet;
    }

    if (pcchNewDestLength)
    {
        *pcchNewDestLength = cchNewDestLength;
    }

    return hr;
}

#ifdef STRSAFE_LOCALE_FUNCTIONS
STRSAFELOCALEWORKERAPI
StringVPrintf_lWorkerW(
    __out_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __out_opt size_t* pcchNewDestLength,
    __in __format_string STRSAFE_LPCWSTR pszFormat,
    __in _locale_t locale,
    __in va_list argList)
{
    HRESULT hr = S_OK;
    int iRet;
    size_t cchMax;
    size_t cchNewDestLength = 0;

    // leave the last space for the null terminator
    cchMax = cchDest - 1;

#if (STRSAFE_USE_SECURE_CRT == 1) && !defined(STRSAFE_LIB_IMPL)
    iRet = _vsnwprintf_s_l(pszDest, cchDest, cchMax, pszFormat, locale, argList);
#else
    iRet = _vsnwprintf_l(pszDest, cchMax, pszFormat, locale, argList);
#endif
    // ASSERT((iRet < 0) || (((size_t)iRet) <= cchMax));

    if ((iRet < 0) || (((size_t)iRet) > cchMax))
    {
        // need to null terminate the string
        pszDest += cchMax;
        *pszDest = L'\0';

        cchNewDestLength = cchMax;

        // we have truncated pszDest
        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
    }
    else if (((size_t)iRet) == cchMax)
    {
        // need to null terminate the string
        pszDest += cchMax;
        *pszDest = L'\0';

        cchNewDestLength = cchMax;
    }
    else
    {
        cchNewDestLength = (size_t)iRet;
    }

    if (pcchNewDestLength)
    {
        *pcchNewDestLength = cchNewDestLength;
    }

    return hr;
}
#endif  // STRSAFE_LOCALE_FUNCTIONS

#endif  // defined(STRSAFE_LIB_IMPL) || !defined(STRSAFE_LIB)

#ifndef STRSAFE_LIB_IMPL
// the StringGetsWorkerA/W functions always run inline since we do not want to
// have a different strsafe lib versions each type of c runtime (eg msvcrt,
// libcmt, etc..)

STRSAFEAPI
StringGetsWorkerA(
    __out_ecount(cchDest) STRSAFE_LPSTR pszDest,
    __in size_t cchDest,
    __out_opt size_t* pcchNewDestLength)
{
    HRESULT hr = S_OK;
    size_t cchNewDestLength = 0;

    if (cchDest == 1)
    {
        *pszDest = '\0';

        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
    }
    else
    {
        while (cchDest > 1)
        {
            char ch;
            int i = getc(stdin);

            if (i == EOF)
            {
                if (cchNewDestLength == 0)
                {
                    // we failed to read anything from stdin
                    hr = STRSAFE_E_END_OF_FILE;
                }

                break;
            }

            ch = (char)i;

            if (ch == '\n')
            {
                break;
            }

            *pszDest++ = ch;
            cchDest--;

            cchNewDestLength++;
        }

        *pszDest = '\0';
    }

    if (pcchNewDestLength)
    {
        *pcchNewDestLength = cchNewDestLength;
    }

    return hr;
}

STRSAFEAPI
StringGetsWorkerW(
    __out_ecount(cchDest) STRSAFE_LPWSTR pszDest,
    __in size_t cchDest,
    __out_opt size_t* pcchNewDestLength)
{
    HRESULT hr = S_OK;
    size_t cchNewDestLength = 0;

    if (cchDest == 1)
    {
        *pszDest = '\0';

        hr = STRSAFE_E_INSUFFICIENT_BUFFER;
    }
    else
    {
        while (cchDest > 1)
        {
            wchar_t ch = getwc(stdin);
            // ASSERT(sizeof(wchar_t) == sizeof(wint_t));

            if (ch == WEOF)
            {
                if (cchNewDestLength == 0)
                {
                    // we failed to read anything from stdin
                    hr = STRSAFE_E_END_OF_FILE;
                }

                break;
            }

            if (ch == L'\n')
            {
                break;
            }

            *pszDest++ = ch;
            cchDest--;

            cchNewDestLength++;
        }

        *pszDest = L'\0';
    }

    if (pcchNewDestLength)
    {
        *pcchNewDestLength = cchNewDestLength;
    }

    return hr;
}

#endif  // !STRSAFE_LIB_IMPL

#if defined(STRSAFE_LIB_IMPL) || !defined(STRSAFE_LIB)

STRSAFEWORKERAPI
StringExHandleFillBehindNullA(
    __out_bcount(cbRemaining) STRSAFE_LPSTR pszDestEnd,
    __in size_t cbRemaining,
    __in DWORD dwFlags)
{
    if (cbRemaining > sizeof(char))
    {
        memset(pszDestEnd + 1, STRSAFE_GET_FILL_PATTERN(dwFlags), cbRemaining - sizeof(char));
    }

    return S_OK;
}

STRSAFEWORKERAPI
StringExHandleFillBehindNullW(
    __out_bcount(cbRemaining) STRSAFE_LPWSTR pszDestEnd,
    __in size_t cbRemaining,
    __in DWORD dwFlags)
{
    if (cbRemaining > sizeof(wchar_t))
    {
        memset(pszDestEnd + 1, STRSAFE_GET_FILL_PATTERN(dwFlags), cbRemaining - sizeof(wchar_t));
    }

    return S_OK;
}

STRSAFEWORKERAPI
StringExHandleOtherFlagsA(
    __out_bcount(cbDest) STRSAFE_LPSTR pszDest,
    __in size_t cbDest,
    __in size_t cchOriginalDestLength,
    __deref_out_ecount(*pcchRemaining) STRSAFE_LPSTR* ppszDestEnd,
    __out size_t* pcchRemaining,
    __in DWORD dwFlags)
{
    size_t cchDest = cbDest / sizeof(char);

    if ((cchDest > 0) && (dwFlags & STRSAFE_NO_TRUNCATION))
    {
        char* pszOriginalDestEnd;

        pszOriginalDestEnd = pszDest + cchOriginalDestLength;

        *ppszDestEnd = pszOriginalDestEnd;
        *pcchRemaining = cchDest - cchOriginalDestLength;

        // null terminate the end of the original string
        *pszOriginalDestEnd = '\0';
    }

    if (dwFlags & STRSAFE_FILL_ON_FAILURE)
    {
        memset(pszDest, STRSAFE_GET_FILL_PATTERN(dwFlags), cbDest);

        if (STRSAFE_GET_FILL_PATTERN(dwFlags) == 0)
        {
            *ppszDestEnd = pszDest;
            *pcchRemaining = cchDest;
        }
        else if (cchDest > 0)
        {
            char* pszDestEnd;

            pszDestEnd = pszDest + cchDest - 1;

            *ppszDestEnd = pszDestEnd;
            *pcchRemaining = 1;

            // null terminate the end of the string
            *pszDestEnd = L'\0';
        }
    }

    if ((cchDest > 0) && (dwFlags & STRSAFE_NULL_ON_FAILURE))
    {
        *ppszDestEnd = pszDest;
        *pcchRemaining = cchDest;

        // null terminate the beginning of the string
        *pszDest = '\0';
    }

    return S_OK;
}

STRSAFEWORKERAPI
StringExHandleOtherFlagsW(
    __out_bcount(cbDest) STRSAFE_LPWSTR pszDest,
    __in size_t cbDest,
    __in size_t cchOriginalDestLength,
    __deref_out_ecount(*pcchRemaining)STRSAFE_LPWSTR* ppszDestEnd,
    __out size_t* pcchRemaining,
    __in DWORD dwFlags)
{
    size_t cchDest = cbDest / sizeof(wchar_t);

    if ((cchDest > 0) && (dwFlags & STRSAFE_NO_TRUNCATION))
    {
        wchar_t* pszOriginalDestEnd;

        pszOriginalDestEnd = pszDest + cchOriginalDestLength;

        *ppszDestEnd = pszOriginalDestEnd;
        *pcchRemaining = cchDest - cchOriginalDestLength;

        // null terminate the end of the original string
        *pszOriginalDestEnd = L'\0';
    }

    if (dwFlags & STRSAFE_FILL_ON_FAILURE)
    {
        memset(pszDest, STRSAFE_GET_FILL_PATTERN(dwFlags), cbDest);

        if (STRSAFE_GET_FILL_PATTERN(dwFlags) == 0)
        {
            *ppszDestEnd = pszDest;
            *pcchRemaining = cchDest;
        }
        else if (cchDest > 0)
        {
            wchar_t* pszDestEnd;

            pszDestEnd = pszDest + cchDest - 1;

            *ppszDestEnd = pszDestEnd;
            *pcchRemaining = 1;

            // null terminate the end of the string
            *pszDestEnd = L'\0';
        }
    }

    if ((cchDest > 0) && (dwFlags & STRSAFE_NULL_ON_FAILURE))
    {
        *ppszDestEnd = pszDest;
        *pcchRemaining = cchDest;

        // null terminate the beginning of the string
        *pszDest = L'\0';
    }

    return S_OK;
}


#endif  // defined(STRSAFE_LIB_IMPL) || !defined(STRSAFE_LIB)


// Do not call these functions, they are worker functions for internal use within this file
#ifdef DEPRECATE_SUPPORTED
#pragma deprecated(StringLengthWorkerA)
#pragma deprecated(StringLengthWorkerW)
#pragma deprecated(UnalignedStringLengthWorkerW)
#pragma deprecated(StringExValidateSrcA)
#pragma deprecated(StringExValidateSrcW)
#pragma deprecated(StringValidateDestA)
#pragma deprecated(StringValidateDestW)
#pragma deprecated(StringExValidateDestA)
#pragma deprecated(StringExValidateDestW)
#pragma deprecated(StringCopyWorkerA)
#pragma deprecated(StringCopyWorkerW)
#pragma deprecated(StringVPrintfWorkerA)
#pragma deprecated(StringVPrintfWorkerW)
#pragma deprecated(StringGetsWorkerA)
#pragma deprecated(StringGetsWorkerW)
#pragma deprecated(StringExHandleFillBehindNullA)
#pragma deprecated(StringExHandleFillBehindNullW)
#pragma deprecated(StringExHandleOtherFlagsA)
#pragma deprecated(StringExHandleOtherFlagsW)
#else
#define StringLengthWorkerA             StringLengthWorkerA_instead_use_StringCchLengthA_or_StringCbLengthA
#define StringLengthWorkerW             StringLengthWorkerW_instead_use_StringCchLengthW_or_StringCbLengthW
#define UnalignedStringLengthWorkerW    UnalignedStringLengthWorkerW_instead_use_UnalignedStringCchLengthW
#define StringExValidateSrcA            StringExValidateSrcA_do_not_call_this_function
#define StringExValidateSrcW            StringExValidateSrcW_do_not_call_this_function
#define StringValidateDestA             StringValidateDestA_do_not_call_this_function
#define StringValidateDestW             StringValidateDestW_do_not_call_this_function
#define StringExValidateDestA           StringExValidateDestA_do_not_call_this_function
#define StringExValidateDestW           StringExValidateDestW_do_not_call_this_function
#define StringCopyWorkerA               StringCopyWorkerA_instead_use_StringCchCopyA_or_StringCbCopyA
#define StringCopyWorkerW               StringCopyWorkerW_instead_use_StringCchCopyW_or_StringCbCopyW
#define StringVPrintfWorkerA            StringVPrintfWorkerA_instead_use_StringCchVPrintfA_or_StringCbVPrintfA
#define StringVPrintfWorkerW            StringVPrintfWorkerW_instead_use_StringCchVPrintfW_or_StringCbVPrintfW
#define StringGetsWorkerA               StringGetsWorkerA_instead_use_StringCchGetsA_or_StringCbGetsA
#define StringGetsWorkerW               StringGetsWorkerW_instead_use_StringCchGetsW_or_StringCbGetsW
#define StringExHandleFillBehindNullA   StringExHandleFillBehindNullA_do_not_call_this_function
#define StringExHandleFillBehindNullW   StringExHandleFillBehindNullW_do_not_call_this_function
#define StringExHandleOtherFlagsA       StringExHandleOtherFlagsA_do_not_call_this_function
#define StringExHandleOtherFlagsW       StringExHandleOtherFlagsW_do_not_call_this_function
#endif // !DEPRECATE_SUPPORTED

// OFFICEDEV: StrSafe is not the right place to deprecate functions 'belonging'
// to Windows, Shlwapi, or the CRT; see O12:1108
#define STRSAFE_NO_DEPRECATE


#ifndef STRSAFE_NO_DEPRECATE
// Deprecate all of the unsafe functions to generate compiletime errors. If you do not want
// this then you can #define STRSAFE_NO_DEPRECATE before including this file
#ifdef DEPRECATE_SUPPORTED

// First all the names that are a/w variants (or shouldn't be #defined by now anyway)
#pragma deprecated(lstrcpyA)
#pragma deprecated(lstrcpyW)
#pragma deprecated(lstrcatA)
#pragma deprecated(lstrcatW)
#pragma deprecated(wsprintfA)
#pragma deprecated(wsprintfW)

#pragma deprecated(StrCpyW)
#pragma deprecated(StrCatW)
#pragma deprecated(StrNCatA)
#pragma deprecated(StrNCatW)
#pragma deprecated(StrCatNA)
#pragma deprecated(StrCatNW)
#pragma deprecated(wvsprintfA)
#pragma deprecated(wvsprintfW)

#pragma deprecated(strcpy)
#pragma deprecated(wcscpy)
#pragma deprecated(strcat)
#pragma deprecated(wcscat)
#pragma deprecated(sprintf)
#pragma deprecated(swprintf)
#pragma deprecated(vsprintf)
#pragma deprecated(vswprintf)
#pragma deprecated(_snprintf)
#pragma deprecated(_snwprintf)
#pragma deprecated(_vsnprintf)
#pragma deprecated(_vsnwprintf)
#pragma deprecated(gets)
#pragma deprecated(_getws)

// Then all the windows.h names
#undef lstrcpy
#undef lstrcat
#undef wsprintf
#undef wvsprintf
#pragma deprecated(lstrcpy)
#pragma deprecated(lstrcat)
#pragma deprecated(wsprintf)
#pragma deprecated(wvsprintf)
#ifdef UNICODE
#define lstrcpy    lstrcpyW
#define lstrcat    lstrcatW
#define wsprintf   wsprintfW
#define wvsprintf  wvsprintfW
#else
#define lstrcpy    lstrcpyA
#define lstrcat    lstrcatA
#define wsprintf   wsprintfA
#define wvsprintf  wvsprintfA
#endif

// Then the shlwapi names
#undef StrCpyA
#undef StrCpy
#undef StrCatA
#undef StrCat
#undef StrNCat
#undef StrCatN
#pragma deprecated(StrCpyA)
#pragma deprecated(StrCatA)
#pragma deprecated(StrCatN)
#pragma deprecated(StrCpy)
#pragma deprecated(StrCat)
#pragma deprecated(StrNCat)
#define StrCpyA lstrcpyA
#define StrCatA lstrcatA
#define StrCatN StrNCat
#ifdef UNICODE
#define StrCpy  StrCpyW
#define StrCat  StrCatW
#define StrNCat StrNCatW
#else
#define StrCpy  lstrcpyA
#define StrCat  lstrcatA
#define StrNCat StrNCatA
#endif

#undef _tcscpy
#undef _ftcscpy
#undef _tcscat
#undef _ftcscat
#undef _stprintf
#undef _sntprintf
#undef _vstprintf
#undef _vsntprintf
#undef _getts
#pragma deprecated(_tcscpy)
#pragma deprecated(_ftcscpy)
#pragma deprecated(_tcscat)
#pragma deprecated(_ftcscat)
#pragma deprecated(_stprintf)
#pragma deprecated(_sntprintf)
#pragma deprecated(_vstprintf)
#pragma deprecated(_vsntprintf)
#pragma deprecated(_getts)
#ifdef _UNICODE
#define _tcscpy     wcscpy
#define _ftcscpy    wcscpy
#define _tcscat     wcscat
#define _ftcscat    wcscat
#define _stprintf   swprintf
#define _sntprintf  _snwprintf
#define _vstprintf  vswprintf
#define _vsntprintf _vsnwprintf
#define _getts      _getws
#else
#define _tcscpy     strcpy
#define _ftcscpy    strcpy
#define _tcscat     strcat
#define _ftcscat    strcat
#define _stprintf   sprintf
#define _sntprintf  _snprintf
#define _vstprintf  vsprintf
#define _vsntprintf _vsnprintf
#define _getts      gets
#endif

#else // DEPRECATE_SUPPORTED

#undef strcpy
#define strcpy      strcpy_instead_use_StringCchCopyA_or_StringCbCopyA;

#undef wcscpy
#define wcscpy      wcscpy_instead_use_StringCchCopyW_or_StringCbCopyW;

#undef strcat
#define strcat      strcat_instead_use_StringCchCatA_or_StringCbCatA;

#undef wcscat
#define wcscat      wcscat_instead_use_StringCchCatW_or_StringCbCatW;

#undef sprintf
#define sprintf     sprintf_instead_use_StringCchPrintfA_or_StringCbPrintfA;

#undef swprintf
#define swprintf    swprintf_instead_use_StringCchPrintfW_or_StringCbPrintfW;

#undef vsprintf
#define vsprintf    vsprintf_instead_use_StringCchVPrintfA_or_StringCbVPrintfA;

#undef vswprintf
#define vswprintf   vswprintf_instead_use_StringCchVPrintfW_or_StringCbVPrintfW;

#undef _snprintf
#define _snprintf   _snprintf_instead_use_StringCchPrintfA_or_StringCbPrintfA;

#undef _snwprintf
#define _snwprintf  _snwprintf_instead_use_StringCchPrintfW_or_StringCbPrintfW;

#undef _vsnprintf
#define _vsnprintf  _vsnprintf_instead_use_StringCchVPrintfA_or_StringCbVPrintfA;

#undef _vsnwprintf
#define _vsnwprintf _vsnwprintf_instead_use_StringCchVPrintfW_or_StringCbVPrintfW;

#undef strcpyA
#define strcpyA     strcpyA_instead_use_StringCchCopyA_or_StringCbCopyA;

#undef strcpyW
#define strcpyW     strcpyW_instead_use_StringCchCopyW_or_StringCbCopyW;

#undef lstrcpy
#define lstrcpy     lstrcpy_instead_use_StringCchCopy_or_StringCbCopy;

#undef lstrcpyA
#define lstrcpyA    lstrcpyA_instead_use_StringCchCopyA_or_StringCbCopyA;

#undef lstrcpyW
#define lstrcpyW    lstrcpyW_instead_use_StringCchCopyW_or_StringCbCopyW;

#undef StrCpy
#define StrCpy      StrCpy_instead_use_StringCchCopy_or_StringCbCopy;

#undef StrCpyA
#define StrCpyA     StrCpyA_instead_use_StringCchCopyA_or_StringCbCopyA;

#undef StrCpyW
#define StrCpyW     StrCpyW_instead_use_StringCchCopyW_or_StringCbCopyW;

#undef _tcscpy
#define _tcscpy     _tcscpy_instead_use_StringCchCopy_or_StringCbCopy;

#undef _ftcscpy
#define _ftcscpy    _ftcscpy_instead_use_StringCchCopy_or_StringCbCopy;

#undef lstrcat
#define lstrcat     lstrcat_instead_use_StringCchCat_or_StringCbCat;

#undef lstrcatA
#define lstrcatA    lstrcatA_instead_use_StringCchCatA_or_StringCbCatA;

#undef lstrcatW
#define lstrcatW    lstrcatW_instead_use_StringCchCatW_or_StringCbCatW;

#undef StrCat
#define StrCat      StrCat_instead_use_StringCchCat_or_StringCbCat;

#undef StrCatA
#define StrCatA     StrCatA_instead_use_StringCchCatA_or_StringCbCatA;

#undef StrCatW
#define StrCatW     StrCatW_instead_use_StringCchCatW_or_StringCbCatW;

#undef StrNCat
#define StrNCat     StrNCat_instead_use_StringCchCatN_or_StringCbCatN;

#undef StrNCatA
#define StrNCatA    StrNCatA_instead_use_StringCchCatNA_or_StringCbCatNA;

#undef StrNCatW
#define StrNCatW    StrNCatW_instead_use_StringCchCatNW_or_StringCbCatNW;

#undef StrCatN
#define StrCatN     StrCatN_instead_use_StringCchCatN_or_StringCbCatN;

#undef StrCatNA
#define StrCatNA    StrCatNA_instead_use_StringCchCatNA_or_StringCbCatNA;

#undef StrCatNW
#define StrCatNW    StrCatNW_instead_use_StringCchCatNW_or_StringCbCatNW;

#undef _tcscat
#define _tcscat     _tcscat_instead_use_StringCchCat_or_StringCbCat;

#undef _ftcscat
#define _ftcscat    _ftcscat_instead_use_StringCchCat_or_StringCbCat;

#undef wsprintf
#define wsprintf    wsprintf_instead_use_StringCchPrintf_or_StringCbPrintf;

#undef wsprintfA
#define wsprintfA   wsprintfA_instead_use_StringCchPrintfA_or_StringCbPrintfA;

#undef wsprintfW
#define wsprintfW   wsprintfW_instead_use_StringCchPrintfW_or_StringCbPrintfW;

#undef wvsprintf
#define wvsprintf   wvsprintf_instead_use_StringCchVPrintf_or_StringCbVPrintf;

#undef wvsprintfA
#define wvsprintfA  wvsprintfA_instead_use_StringCchVPrintfA_or_StringCbVPrintfA;

#undef wvsprintfW
#define wvsprintfW  wvsprintfW_instead_use_StringCchVPrintfW_or_StringCbVPrintfW;

#undef _vstprintf
#define _vstprintf  _vstprintf_instead_use_StringCchVPrintf_or_StringCbVPrintf;

#undef _vsntprintf
#define _vsntprintf _vsntprintf_instead_use_StringCchVPrintf_or_StringCbVPrintf;

#undef _stprintf
#define _stprintf   _stprintf_instead_use_StringCchPrintf_or_StringCbPrintf;

#undef _sntprintf
#define _sntprintf  _sntprintf_instead_use_StringCchPrintf_or_StringCbPrintf;

#undef _getts
#define _getts      _getts_instead_use_StringCchGets_or_StringCbGets;

#undef gets
#define gets        _gets_instead_use_StringCchGetsA_or_StringCbGetsA;

#undef _getws
#define _getws      _getws_instead_use_StringCchGetsW_or_StringCbGetsW;

#endif  // DEPRECATE_SUPPORTED
#endif  // !STRSAFE_NO_DEPRECATE

#pragma warning(pop)

#endif  // _STRSAFE_H_INCLUDED_