#pragma once
#include <string>
using namespace std;

extern wstring emptystring; 
wstring loadstring(DWORD dwID);
wstring formatV(wstring const& szMsg, va_list argList);
wstring format(LPCWSTR szMsg, ...);
wstring formatmessagesys(DWORD dwID);
wstring formatmessage(DWORD dwID, ...);
wstring formatmessage(wstring const szMsg, ...);
LPTSTR wstringToLPTSTR(wstring const& src);
CString wstringToCString(wstring const& src);
wstring LPCTSTRToWstring(LPCTSTR src);
wstring LPCSTRToWstring(LPCSTR src);
void wstringToLower(wstring src);
ULONG wstringToUlong(wstring const& src, int radix);
ULONG CStringToUlong(CString const& src, int radix);
wstring StripCarriage(wstring szString);

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