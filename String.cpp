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
_Check_return_ HRESULT AnsiToUnicode(_In_opt_z_ LPCSTR pszA, _Out_z_cap_(cchszA) LPWSTR* ppszW, _Out_ size_t* cchszW, size_t cchszA)
{
	HRESULT hRes = S_OK;
	if (!ppszW || *cchszW) return MAPI_E_INVALID_PARAMETER;
	*ppszW = NULL;
	*cchszW = 0;
	if (NULL == pszA) return S_OK;
	if (!cchszA) return S_OK;

	// Get our buffer size
	int iRet = 0;
	EC_D(iRet, MultiByteToWideChar(
		CP_ACP,
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
			CP_ACP,
			0,
			pszA,
			(int)cchszA,
			pszW,
			iRet));
		if (SUCCEEDED(hRes))
		{
			*ppszW = pszW;
			*cchszW = iRet;
		}
		else
		{
			delete[] pszW;
		}
	}

	return hRes;
}

// if cchszA == -1, MultiByteToWideChar will compute the length
// Delete with delete[]
_Check_return_ HRESULT AnsiToUnicode(_In_opt_z_ LPCSTR pszA, _Out_z_cap_(cchszA) LPWSTR* ppszW, size_t cchszA)
{
	size_t cchsW = 0;
	return AnsiToUnicode(pszA, ppszW, &cchsW, cchszA);
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
		CP_ACP,
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
			CP_ACP,
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

wstring BinToTextString(_In_ LPSBinary lpBin, bool bMultiLine)
{
	if (!lpBin || !lpBin->cb || !lpBin->lpb) return L"";

	wstring szBin;

	ULONG i;
	for (i = 0; i < lpBin->cb; i++)
	{
		// Any printable extended ASCII character gets mapped directly
		if (lpBin->lpb[i] >= 0x20 &&
			lpBin->lpb[i] <= 0xFE)
		{
			szBin += lpBin->lpb[i];
		}
		// If we allow multiple lines, we accept tab, LF and CR
		else if (bMultiLine &&
			(lpBin->lpb[i] == 9 || // Tab
				lpBin->lpb[i] == 10 || // Line Feed
				lpBin->lpb[i] == 13))  // Carriage Return
		{
			szBin += lpBin->lpb[i];
		}
		// Everything else is a dot
		else
		{
			szBin += L'.';
		}
	}

	return szBin;
}

wstring BinToHexString(_In_opt_count_(cb) LPBYTE lpb, size_t cb, bool bPrependCB)
{
	wstring lpsz;

	if (bPrependCB)
	{
		lpsz = format(L"cb: %u lpb: ", (UINT)cb); // STRING_OK
	}

	if (!cb || !lpb)
	{
		lpsz += L"NULL";
	}
	else
	{
		ULONG i = 0;
		for (i = 0; i < cb; i++)
		{
			BYTE bLow = (BYTE)((lpb[i]) & 0xf);
			BYTE bHigh = (BYTE)((lpb[i] >> 4) & 0xf);
			wchar_t szLow = (wchar_t)((bLow <= 0x9) ? L'0' + bLow : L'A' + bLow - 0xa);
			wchar_t szHigh = (wchar_t)((bHigh <= 0x9) ? L'0' + bHigh : L'A' + bHigh - 0xa);

			lpsz += szHigh;
			lpsz += szLow;
		}
	}

	return lpsz;
}

wstring BinToHexString(_In_opt_ LPSBinary lpBin, bool bPrependCB)
{
	if (!lpBin) return L"";

	return BinToHexString(
		lpBin->lpb,
		lpBin->cb,
		bPrependCB);
}

bool stripPrefix(wstring& str, wstring prefix)
{
	size_t length = prefix.length();
	if (str.compare(0, length, prefix) == 0)
	{
		str.erase(0, length);
		return true;
	}

	return false;
}

vector<BYTE> HexStringToBin(_In_ wstring lpsz)
{
	// remove junk
	WCHAR szJunk[] = L"\r\n\t -.,\\/'{}`\""; // STRING_OK
	for (unsigned int i = 0; i < _countof(szJunk); ++i)
	{
		lpsz.erase(std::remove(lpsz.begin(), lpsz.end(), szJunk[i]), lpsz.end());
	}

	// strip one (and only one) prefix
	stripPrefix(lpsz, L"0x") ||
		stripPrefix(lpsz, L"0X") ||
		stripPrefix(lpsz, L"x") ||
		stripPrefix(lpsz, L"X");

	size_t iCur = 0;
	WCHAR szTmp[3] = { 0 };
	size_t cchStrLen = lpsz.length();
	vector<BYTE> lpb;

	// We have a clean string now. If it's of odd length, we're done.
	if (cchStrLen % 2 != 0) return vector<BYTE>();

	// convert two characters at a time
	while (iCur < cchStrLen)
	{
		// Check for valid hex characters
		if (!isxdigit(lpsz[iCur]) || !isxdigit(lpsz[iCur + 1]))
		{
			return vector<BYTE>();
		}

		szTmp[0] = lpsz[iCur];
		szTmp[1] = lpsz[iCur + 1];
		lpb.push_back((BYTE)wcstol(szTmp, NULL, 16));
		iCur += 2;
	}

	return lpb;
}

// Converts byte vector to LPBYTE allocated with new
LPBYTE ByteVectorToLPBYTE(vector<BYTE> bin)
{
	if (bin.empty()) return NULL;

	LPBYTE lpBin = new BYTE[bin.size()];
	if (lpBin != NULL)
	{
		memset(lpBin, 0, bin.size());
		memcpy(lpBin, &bin[0], bin.size());
		return lpBin;
	}

	return NULL;
}