#pragma once
#include <string>

std::wstring loadstring(DWORD dwID);
std::wstring format(const LPWSTR fmt, ...);
std::wstring formatmessage(DWORD dwID, ...);
LPTSTR wstringToLPTSTR(std::wstring src);