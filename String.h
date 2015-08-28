#pragma once
#include <string>
using namespace std;

wstring loadstring(DWORD dwID);
wstring formatV(wstring szMsg, va_list argList);
wstring format(const LPWSTR szMsg, ...);
wstring formatmessage(DWORD dwID, ...);
LPTSTR wstringToLPTSTR(wstring src);
CString wstringToCString(wstring src);
wstring LPCTSTRToWstring(LPCTSTR src);
wstring LPCSTRToWstring(LPCSTR src);
wstring stringToWstring(string src);
void wstringToLower(wstring src);
ULONG wstringToUlong(wstring src, int radix);

// Unicode support
_Check_return_ HRESULT AnsiToUnicode(_In_opt_z_ LPCSTR pszA, _Out_z_cap_(cchszA) LPWSTR* ppszW, size_t cchszA = -1);
_Check_return_ HRESULT UnicodeToAnsi(_In_z_ LPCWSTR pszW, _Out_z_cap_(cchszW) LPSTR* ppszA, size_t cchszW = -1);

bool IsNullOrEmptyW(LPWSTR szStr);
bool IsNullOrEmptyA(LPCSTR szStr);

#ifdef UNICODE
#define IsNullOrEmpty IsNullOrEmptyW
#else
#define IsNullOrEmpty IsNullOrEmptyA
#endif