#include "stdafx.h"
#include <Interpret/String.h>
#include <locale>
#include <sstream>
#include <iterator>
#include <functional>

wstring emptystring = L"";

wstring formatV(LPCWSTR szMsg, va_list argList)
{
	auto len = _vscwprintf(szMsg, argList);
	if (0 != len)
	{
		len++;
		auto buffer = new wchar_t[len];
		memset(buffer, 0, sizeof(wchar_t)* len);
		if (_vsnwprintf_s(buffer, len, _TRUNCATE, szMsg, argList) > 0)
		{
			wstring szOut(buffer);
			delete[] buffer;
			return szOut;
		}

		delete[] buffer;
	}

	return L"";
}

#ifdef CHECKFORMATPARAMS
#undef format
#endif

// Takes format strings with %x %d...
wstring format(LPCWSTR szMsg, ...)
{
	va_list argList;
	va_start(argList, szMsg);
	auto ret = formatV(szMsg, argList);
	va_end(argList);
	return ret;
}

wstring loadstring(HINSTANCE hInstance, DWORD dwID)
{
	wstring fmtString;
	LPWSTR buffer = nullptr;
	size_t len = LoadStringW(hInstance, dwID, reinterpret_cast<PWCHAR>(&buffer), 0);

	if (len)
	{
		fmtString.assign(buffer, len);
	}

	return fmtString;
}

wstring loadstring(DWORD dwID)
{
	return loadstring(nullptr, dwID);
}

wstring formatmessageV(const wstring& szMsg, va_list argList)
{
	LPWSTR buffer = nullptr;
	auto dw = FormatMessageW(FORMAT_MESSAGE_FROM_STRING | FORMAT_MESSAGE_ALLOCATE_BUFFER, szMsg.c_str(), 0, 0, reinterpret_cast<LPWSTR>(&buffer), 0, &argList);
	if (dw)
	{
		auto ret = wstring(buffer);
		(void)LocalFree(buffer);
		return ret;
	}

	return L"";
}

wstring formatmessagesys(DWORD dwID)
{
	LPWSTR buffer = nullptr;
	auto dw = FormatMessageW(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM, nullptr, dwID, 0, reinterpret_cast<LPWSTR>(&buffer), 0, nullptr);
	if (dw)
	{
		auto ret = wstring(buffer);
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
	auto ret = formatmessageV(loadstring(dwID), argList);
	va_end(argList);
	return ret;
}

// Takes format strings with %1 %2 %3...
wstring formatmessage(const wstring szMsg, ...)
{
	va_list argList;
	va_start(argList, szMsg);
	auto ret = formatmessageV(szMsg, argList);
	va_end(argList);
	return ret;
}

tstring wstringTotstring(const wstring& src)
{
#ifdef _UNICODE
	return src;
#else
	return wstringTostring(src);
#endif
}

string wstringTostring(const wstring& src)
{
	return string(src.begin(), src.end());
}

wstring stringTowstring(const string& src)
{
	return wstring(src.begin(), src.end());
}

wstring LPCTSTRToWstring(LPCTSTR src)
{
#ifdef UNICODE
	return src ? src : L"";
#else
	return LPCSTRToWstring(src);
#endif
}

// Careful calling this with string as it *will* lose embedded null
// use stringTowstring instead
wstring LPCSTRToWstring(LPCSTR src)
{
	if (!src) return L"";
	string ansi = src;
	return wstring(ansi.begin(), ansi.end());
}

wstring wstringToLower(const wstring& src)
{
	auto dst = src;
	transform(dst.begin(), dst.end(), dst.begin(), tolower);
	return dst;
}

// Converts a wstring to a ulong. Will return 0 if string is empty or contains non-numeric data.
ULONG wstringToUlong(const wstring& src, int radix, bool rejectInvalidCharacters)
{
	if (src.empty()) return 0;

	LPWSTR szEndPtr = nullptr;
	auto ulArg = wcstoul(src.c_str(), &szEndPtr, radix);

	if (rejectInvalidCharacters)
	{
		// if szEndPtr is pointing to something other than NULL, this must be a string
		if (!szEndPtr || *szEndPtr)
		{
			ulArg = NULL;
		}
	}

	return ulArg;
}

// Converts a wstring to a long. Will return 0 if string is empty or contains non-numeric data.
long wstringToLong(const wstring& src, int radix)
{
	if (src.empty()) return 0;

	LPWSTR szEndPtr = nullptr;
	auto lArg = wcstol(src.c_str(), &szEndPtr, radix);

	// if szEndPtr is pointing to something other than NULL, this must be a string
	if (!szEndPtr || *szEndPtr)
	{
		lArg = NULL;
	}

	return lArg;
}

// Converts a wstring to a double. Will return 0 if string is empty or contains non-numeric data.
double wstringToDouble(const wstring& src)
{
	if (src.empty()) return 0;

	LPWSTR szEndPtr = nullptr;
	auto dArg = wcstod(src.c_str(), &szEndPtr);

	// if szEndPtr is pointing to something other than NULL, this must be a string
	if (!szEndPtr || *szEndPtr)
	{
		dArg = NULL;
	}

	return dArg;
}

__int64 wstringToInt64(const wstring& src)
{
	if (src.empty()) return 0;

	return _wtoi64(src.c_str());
}

wstring strip(const wstring& str, std::function<bool(const WCHAR&)> func)
{
	wstring result;
	result.reserve(str.length());
	remove_copy_if(str.begin(), str.end(), back_inserter(result), func);
	return result;
}

wstring StripCharacter(const wstring& szString, const WCHAR& character)
{
	return strip(szString, [character](const WCHAR& chr)
	{
		return chr == character;
	});
}

wstring StripCarriage(const wstring& szString)
{
	return StripCharacter(szString, L'\r');
}

wstring CleanString(const wstring& szString)
{
	return strip(szString, [](const WCHAR & chr)
	{
		return wstring(L", \r\n").find(chr) != wstring::npos;
	});
}

wstring TrimString(const wstring& szString)
{
	size_t first = szString.find_first_not_of(' ');
	if (first == string::npos)
		return emptystring;
	size_t last = szString.find_last_not_of(' ');
	return szString.substr(first, (last - first + 1));
}

wstring replace(const wstring& str, std::function<bool(const WCHAR&)> func, const WCHAR& chr)
{
	wstring result;
	result.reserve(str.length());
	replace_copy_if(str.begin(), str.end(), back_inserter(result), func, chr);
	return result;
}

wstring ScrubStringForXML(const wstring& szString)
{
	return replace(szString, [](const WCHAR& chr)
	{
		// Replace anything less than 0x20 except tab, carriage return and linefeed
		return chr < 0x20 && wstring(L"\t\r\n").find(chr) == wstring::npos;
	}, L'.');
}

// Processes szFileIn, replacing non file system characters with underscores
// Do NOT call with full path - just file names
wstring SanitizeFileName(const wstring& szFileIn)
{
	return replace(szFileIn, [](const WCHAR& chr)
	{
		return wstring(L"^&*-+=[]\\|;:\",<>/?\r\n").find(chr) != wstring::npos;
	}, L'_');
}

wstring indent(int iIndent)
{
	return wstring(iIndent, L'\t');
}

// Converts binary data to a string, assuming source string was unicode
wstring BinToTextStringW(const vector<BYTE>& lpByte, bool bMultiLine)
{
	SBinary bin = { 0 };
	bin.cb = static_cast<ULONG>(lpByte.size());
	bin.lpb = const_cast<LPBYTE>(lpByte.data());
	return BinToTextStringW(&bin, bMultiLine);
}

// Converts binary data to a string, assuming source string was unicode
wstring BinToTextStringW(_In_ const LPSBinary lpBin, bool bMultiLine)
{
	if (!lpBin || !lpBin->cb || lpBin->cb % sizeof WCHAR || !lpBin->lpb) return L"";

	wstring szBin(reinterpret_cast<LPWSTR>(lpBin->lpb), lpBin->cb / sizeof WCHAR);
	std::replace_if(szBin.begin(), szBin.end(), [bMultiLine](const WCHAR & chr)
	{
		// Any printable extended ASCII character gets mapped directly
		if (chr >= 0x20 &&
			chr <= 0xFE)
		{
			return false;
		}
		// If we allow multiple lines, we accept tab, LF and CR
		else if (bMultiLine &&
			(chr == 9 || // Tab
				chr == 10 || // Line Feed
				chr == 13))  // Carriage Return
		{
			return false;
		}

		return true;
	}, L'.');

	return szBin;
}

// Converts binary data to a string, assuming source string was single byte
wstring BinToTextString(_In_ const LPSBinary lpBin, bool bMultiLine)
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

wstring BinToHexString(_In_opt_count_(cb) const BYTE* lpb, size_t cb, bool bPrependCB)
{
	wstring lpsz;

	if (bPrependCB)
	{
		lpsz = format(L"cb: %u lpb: ", static_cast<UINT>(cb)); // STRING_OK
	}

	if (!cb || !lpb)
	{
		lpsz += L"NULL";
	}
	else
	{
		for (ULONG i = 0; i < cb; i++)
		{
			auto bLow = static_cast<BYTE>(lpb[i] & 0xf);
			auto bHigh = static_cast<BYTE>(lpb[i] >> 4 & 0xf);
			auto szLow = static_cast<wchar_t>(bLow <= 0x9 ? L'0' + bLow : L'A' + bLow - 0xa);
			auto szHigh = static_cast<wchar_t>(bHigh <= 0x9 ? L'0' + bHigh : L'A' + bHigh - 0xa);

			lpsz += szHigh;
			lpsz += szLow;
		}
	}

	return lpsz;
}

wstring BinToHexString(const vector<BYTE>& lpByte, bool bPrependCB)
{
	SBinary sBin = { 0 };
	sBin.cb = static_cast<ULONG>(lpByte.size());
	sBin.lpb = const_cast<LPBYTE>(lpByte.data());
	return BinToHexString(&sBin, bPrependCB);
}

wstring BinToHexString(_In_opt_ const SBinary* lpBin, bool bPrependCB)
{
	if (!lpBin) return L"";

	return BinToHexString(
		lpBin->lpb,
		lpBin->cb,
		bPrependCB);
}

bool stripPrefix(wstring& str, const wstring& prefix)
{
	auto length = prefix.length();
	if (str.compare(0, length, prefix) == 0)
	{
		str.erase(0, length);
		return true;
	}

	return false;
}

// Converts hex string in lpsz to a binary buffer.
// If cbTarget != 0, caps the number of bytes converted at cbTarget
vector<BYTE> HexStringToBin(_In_ const wstring& input, size_t cbTarget)
{
	// If our target is odd, we can't convert
	if (cbTarget % 2 != 0) return vector<BYTE>();

	// remove junk
	wstring szJunk = L"\r\n\t -.,\\/'{}`\""; // STRING_OK
	auto lpsz = strip(input, [szJunk](const WCHAR & chr)
	{
		return szJunk.find(chr) != wstring::npos;
	});

	// strip one (and only one) prefix
	stripPrefix(lpsz, L"0x") ||
		stripPrefix(lpsz, L"0X") ||
		stripPrefix(lpsz, L"x") ||
		stripPrefix(lpsz, L"X");

	auto cchStrLen = lpsz.length();

	vector<BYTE> lpb;
	WCHAR szTmp[3] = { 0 };
	size_t iCur = 0;
	size_t cbConverted = 0;

	// convert two characters at a time
	while (iCur < cchStrLen && (cbTarget == 0 || cbConverted < cbTarget))
	{
		// Check for valid hex characters
		if (lpsz[iCur] > 255 || lpsz[iCur + 1] > 255 || !isxdigit(lpsz[iCur]) || !isxdigit(lpsz[iCur + 1]))
		{
			return vector<BYTE>();
		}

		szTmp[0] = lpsz[iCur];
		szTmp[1] = lpsz[iCur + 1];
		lpb.push_back(static_cast<BYTE>(wcstol(szTmp, nullptr, 16)));
		iCur += 2;
		cbConverted++;
	}

	return lpb;
}

// Converts byte vector to LPBYTE allocated with new
LPBYTE ByteVectorToLPBYTE(const vector<BYTE>& bin)
{
	if (bin.empty()) return nullptr;

	auto lpBin = new BYTE[bin.size()];
	if (lpBin != nullptr)
	{
		memset(lpBin, 0, bin.size());
		memcpy(lpBin, &bin[0], bin.size());
		return lpBin;
	}

	return nullptr;
}

vector<wstring> split(const wstring& str, const wchar_t delim)
{
	auto ss = std::wstringstream(str);
	wstring item;
	vector<wstring> elems;
	while (getline(ss, item, delim))
	{
		elems.push_back(item);
	}

	return elems;
}

wstring join(const vector<wstring>& elems, const wstring& delim)
{
	std::wstringstream ss;
	for (size_t i = 0; i < elems.size(); ++i)
	{
		if (i != 0)
			ss << delim;
		ss << elems[i];
	}

	return ss.str();
}

wstring join(const vector<wstring>& elems, const wchar_t delim)
{
	return join(elems, wstring(1, delim));
}