#include "stdafx.h"
#include "string.h"

std::wstring format(const LPWSTR fmt, ...)
{
	LPWSTR buffer = NULL;
	va_list vl;
	va_start(vl, fmt);
	int len = _vscwprintf(fmt, vl);
	if (0 != len)
	{
		len++;
		buffer = new wchar_t[len];
		memset(buffer, 0, sizeof(wchar_t)* len);
		(void)_vsnwprintf_s(buffer, len, len, fmt, vl);
	}

	std::wstring ret(buffer);
	va_end(vl);
	delete[] buffer;
	return ret;
}

std::wstring loadstring(DWORD dwID)
{
	std::wstring fmtString;
	LPWSTR buffer = 0;
	size_t len = ::LoadStringW(NULL, dwID, (PWCHAR)&buffer, 0);

	if (len)
	{
		fmtString.assign(buffer, len);
	}

	return fmtString;
}

std::wstring formatmessage(DWORD dwID, ...)
{
	std::wstring format = loadstring(dwID);

	LPWSTR buffer = NULL;
	std::wstring ret;
	va_list vl;
	va_start(vl, dwID);
	DWORD dw = FormatMessageW(FORMAT_MESSAGE_FROM_STRING | FORMAT_MESSAGE_ALLOCATE_BUFFER, format.c_str(), 0, 0, (LPWSTR)&buffer, 0, &vl);
	if (dw)
	{
		ret = std::wstring(buffer);
		(void)LocalFree(buffer);
	}

	va_end(vl);
	return ret;
}

LPTSTR wstringToLPTSTR(std::wstring src)
{
	size_t cch = src.length();
	if (!cch) return NULL;

	cch++; // Null terminator
	LPTSTR dst = new TCHAR[cch];
	if (dst)
	{
		HRESULT hRes = S_OK;
		EC_H(StringCchPrintf(dst, cch, _T("%ws"), src.c_str()))
	}

	return dst;
}