#pragma once
#include <string>

std::wstring loadstring(DWORD dwID);
std::wstring format(const LPWSTR fmt, ...);
std::wstring formatmessage(DWORD dwID, ...);
LPTSTR wstringToLPTSTR(std::wstring src);

// Unicode support
_Check_return_ HRESULT AnsiToUnicode(_In_opt_z_ LPCSTR pszA, _Out_z_cap_(cchszA) LPWSTR* ppszW, size_t cchszA = -1);
_Check_return_ HRESULT UnicodeToAnsi(_In_z_ LPCWSTR pszW, _Out_z_cap_(cchszW) LPSTR* ppszA, size_t cchszW = -1);