#pragma once
#include <string>
using namespace std;

wstring loadstring(DWORD dwID);
wstring format(const LPWSTR fmt, ...);
wstring formatmessage(DWORD dwID, ...);
LPTSTR wstringToLPTSTR(wstring src);
LPWSTR wstringToLPWSTR(wstring src);
CString wstringToCString(::wstring src);
wstring LPTSTRToWstring(LPTSTR src);
wstring LPSTRToWstring(LPSTR src);
_Check_return_ LPWSTR CStringToLPWSTR(CString szCString);

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