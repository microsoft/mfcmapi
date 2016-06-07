#include "stdafx.h"
#include "string.h"
#include <algorithm>
#include <locale>
#include <codecvt>

wstring emptystring = L"";

wstring formatV(wstring const& szMsg, va_list argList)
{
	int len = _vscwprintf(szMsg.c_str(), argList);
	if (0 != len)
	{
		len++;
		LPWSTR buffer = new wchar_t[len];
		memset(buffer, 0, sizeof(wchar_t)* len);
		if (_vsnwprintf_s(buffer, len, _TRUNCATE, szMsg.c_str(), argList) > 0)
		{
			wstring szOut(buffer);
			delete[] buffer;
			return szOut;
		}

		delete[] buffer;
	}

	return L"";
}

// Takes format strings with %x %d...
wstring format(LPCWSTR szMsg, ...)
{
	va_list argList;
	va_start(argList, szMsg);
	wstring ret = formatV(szMsg, argList);
	va_end(argList);
	return ret;
}

wstring loadstring(DWORD dwID)
{
	wstring fmtString;
	LPWSTR buffer = 0;
	size_t len = ::LoadStringW(NULL, dwID, (PWCHAR)&buffer, 0);

	if (len)
	{
		fmtString.assign(buffer, len);
	}

	return fmtString;
}

wstring formatmessageV(wstring const& szMsg, va_list argList)
{
	LPWSTR buffer = NULL;
	DWORD dw = FormatMessageW(FORMAT_MESSAGE_FROM_STRING | FORMAT_MESSAGE_ALLOCATE_BUFFER, szMsg.c_str(), 0, 0, (LPWSTR)&buffer, 0, &argList);
	if (dw)
	{
		wstring ret = wstring(buffer);
		(void)LocalFree(buffer);
		return ret;
	}

	return L"";
}

wstring formatmessagesys(DWORD dwID)
{
	LPWSTR buffer = NULL;
	DWORD dw = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM, 0, dwID, 0, (LPWSTR)&buffer, 0, 0);
	if (dw)
	{
		wstring ret = wstring(buffer);
		(void)LocalFree(buffer);
		return ret;
	}

	return L"";
}

// Takes format strings with %1 %2 %3...
wstring formatmessage(DWORD dwID, ...)
{
	va_list argList;
	va_start(argList, dwID);
	wstring ret = formatmessageV(loadstring(dwID), argList);
	va_end(argList);
	return ret;
}

// Takes format strings with %1 %2 %3...
wstring formatmessage(wstring const szMsg, ...)
{
	va_list argList;
	va_start(argList, szMsg);
	wstring ret = formatmessageV(szMsg, argList);
	va_end(argList);
	return ret;
}

// Allocates with new. Free with delete[]
LPTSTR wstringToLPTSTR(wstring const& src)
{
	LPTSTR dst = NULL;
#ifdef UNICODE
	size_t cch = src.length();
	if (!cch) return NULL;

	cch++; // Null terminator
	dst = new WCHAR[cch];
	if (dst)
	{
		memcpy(dst, src.c_str(), cch * sizeof(WCHAR));
	}
#else
	HRESULT hRes = S_OK;
	EC_H(UnicodeToAnsi(src.c_str(), &dst));
#endif

	return dst;
}

CString wstringToCString(wstring const& src)
{
	return src.c_str();
}

CStringA wstringToCStringA(wstring const& src)
{
	return src.c_str();
}

wstring LPCTSTRToWstring(LPCTSTR src)
{
#ifdef UNICODE
	return src ? src : L"";
#else
	return LPCSTRToWstring(src);
#endif
}

wstring LPCSTRToWstring(LPCSTR src)
{
	if (!src) return L"";
	string ansi = src;
	return wstring(ansi.begin(), ansi.end());
}

CString LPCSTRToCString(LPCSTR src)
{
	return src;
}

CStringA LPCTSTRToCStringA(LPCTSTR src)
{
	return src;
}

void wstringToLower(wstring src)
{
	transform(src.begin(), src.end(), src.begin(), ::tolower);
}

// Converts a wstring to a number. Will return 0 if string is empty or contains non-numeric data.
ULONG wstringToUlong(wstring const& src, int radix)
{
	if (src.empty()) return 0;

	LPWSTR szEndPtr = NULL;
	ULONG ulArg = wcstoul(src.c_str(), &szEndPtr, radix);

	// if szEndPtr is pointing to something other than NULL, this must be a string
	if (!szEndPtr || *szEndPtr)
	{
		ulArg = NULL;
	}

	return ulArg;
}

// Converts a CString to a number. Will return 0 if string is empty or contains non-numeric data.
ULONG CStringToUlong(CString const& src, int radix)
{
	if (IsNullOrEmpty((LPCTSTR)src)) return 0;

	LPTSTR szEndPtr = NULL;
	ULONG ulArg = _tcstoul((LPCTSTR)src, &szEndPtr, radix);

	// if szEndPtr is pointing to something other than NULL, this must be a string
	if (!szEndPtr || *szEndPtr)
	{
		ulArg = NULL;
	}

	return ulArg;
}

wstring StripCarriage(wstring szString)
{
	szString.erase(remove(szString.begin(), szString.end(), L'\r'), szString.end());
	return szString;
}

// if cchszA == -1, MultiByteToWideChar will compute the length
// Delete with delete[]
_Check_return_ HRESULT AnsiToUnicode(_In_opt_z_ LPCSTR pszA, _Out_z_cap_(cchszA) LPWSTR* ppszW, size_t cchszA)
{
	HRESULT hRes = S_OK;
	if (!ppszW) return MAPI_E_INVALID_PARAMETER;
	*ppszW = NULL;
	if (NULL == pszA) return S_OK;
	if (!cchszA) return S_OK;

	// Get our buffer size
	int iRet = 0;
	EC_D(iRet, MultiByteToWideChar(
		CP_UTF8,
		0,
		pszA,
		(int)cchszA,
		NULL,
		NULL));
	if (SUCCEEDED(hRes) && 0 != iRet)
	{
		// MultiByteToWideChar returns num of chars
		LPWSTR pszW = new WCHAR[iRet];

		EC_D(iRet, MultiByteToWideChar(
			CP_UTF8,
			0,
			pszA,
			(int)cchszA,
			pszW,
			iRet));
		if (SUCCEEDED(hRes))
		{
			*ppszW = pszW;
		}
		else
		{
			delete[] pszW;
		}
	}

	return hRes;
}

// if cchszW == -1, WideCharToMultiByte will compute the length
// Delete with delete[]
_Check_return_ HRESULT UnicodeToAnsi(_In_z_ LPCWSTR pszW, _Out_z_cap_(cchszW) LPSTR* ppszA, size_t cchszW)
{
	HRESULT hRes = S_OK;
	if (!ppszA) return MAPI_E_INVALID_PARAMETER;
	*ppszA = NULL;
	if (NULL == pszW) return S_OK;

	// Get our buffer size
	int iRet = 0;
	EC_D(iRet, WideCharToMultiByte(
		CP_UTF8,
		0,
		pszW,
		(int)cchszW,
		NULL,
		NULL,
		NULL,
		NULL));
	if (SUCCEEDED(hRes) && 0 != iRet)
	{
		// WideCharToMultiByte returns num of bytes
		LPSTR pszA = (LPSTR) new BYTE[iRet];

		EC_D(iRet, WideCharToMultiByte(
			CP_UTF8,
			0,
			pszW,
			(int)cchszW,
			pszA,
			iRet,
			NULL,
			NULL));
		if (SUCCEEDED(hRes))
		{
			*ppszA = pszA;
		}
		else
		{
			delete[] pszA;
		}
	}

	return hRes;
}

bool IsNullOrEmptyW(LPCWSTR szStr)
{
	return !szStr || !szStr[0];
}

bool IsNullOrEmptyA(LPCSTR szStr)
{
	return !szStr || !szStr[0];
}