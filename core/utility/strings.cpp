#include <core/stdafx.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>

namespace strings
{
	std::wstring emptystring;

	std::wstring formatV(LPCWSTR szMsg, va_list argList)
	{
		auto len = _vscwprintf(szMsg, argList);
		if (0 != len)
		{
			auto buffer = std::wstring(len + 1, '\0'); // Include extra since _vsnwprintf_s writes a null terminator
			if (_vsnwprintf_s(const_cast<wchar_t*>(buffer.c_str()), len + 1, _TRUNCATE, szMsg, argList) > 0)
			{
				buffer.resize(len); // Resize to exact length before returning
				return buffer;
			}
		}

		return emptystring;
	}

#ifdef CHECKFORMATPARAMS
#undef format
#endif

	// Takes format strings with %x %d...
	std::wstring format(LPCWSTR szMsg, ...)
	{
		va_list argList;
		va_start(argList, szMsg);
		auto ret = formatV(szMsg, argList);
		va_end(argList);
		return ret;
	}

	HINSTANCE g_testInstance = nullptr;
	// By default, we call LoadStringW with a null hInstance.
	// This will try to load the string from the executable, which is fine for MFCMAPI and MrMAPI
	// In our unit tests, we must load strings from UnitTest.dll, so we use setTestInstance
	// to populate an appropriate HINSTANCE
	void setTestInstance(HINSTANCE hInstance) { g_testInstance = hInstance; }

	std::wstring loadstring(DWORD dwID)
	{
		std::wstring fmtString;
		LPWSTR buffer = nullptr;
		const size_t len = LoadStringW(g_testInstance, dwID, reinterpret_cast<PWCHAR>(&buffer), 0);

		if (len)
		{
			fmtString.assign(buffer, len);
		}

		return fmtString;
	}

	std::wstring formatmessageV(LPCWSTR szMsg, va_list argList)
	{
		LPWSTR buffer = nullptr;
		const auto dw = FormatMessageW(
			FORMAT_MESSAGE_FROM_STRING | FORMAT_MESSAGE_ALLOCATE_BUFFER,
			szMsg,
			0,
			0,
			reinterpret_cast<LPWSTR>(&buffer),
			0,
			&argList);
		if (dw)
		{
			auto ret = std::wstring(buffer);
			(void) LocalFree(buffer);
			return ret;
		}

		return L"";
	}

	std::wstring formatmessagesys(DWORD dwID)
	{
		LPWSTR buffer = nullptr;
		const auto dw = FormatMessageW(
			FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM,
			nullptr,
			dwID,
			0,
			reinterpret_cast<LPWSTR>(&buffer),
			0,
			nullptr);
		if (dw)
		{
			auto ret = std::wstring(buffer);
			(void) LocalFree(buffer);
			return ret;
		}

		return L"";
	}

	// Takes format strings with %1 %2 %3...
	std::wstring formatmessage(DWORD dwID, ...)
	{
		va_list argList;
		va_start(argList, dwID);
		auto ret = formatmessageV(loadstring(dwID).c_str(), argList);
		va_end(argList);
		return ret;
	}

	// Takes format strings with %1 %2 %3...
	std::wstring formatmessage(LPCWSTR szMsg, ...)
	{
		va_list argList;
		va_start(argList, szMsg);
		auto ret = formatmessageV(szMsg, argList);
		va_end(argList);
		return ret;
	}

	// This function should be used as little as possible as for many strings this conversion does not make sense
	std::basic_string<TCHAR> wstringTotstring(const std::wstring& src)
	{
#ifdef _UNICODE
		return src;
#else
		return wstringTostring(src);
#endif
	}

	// This function should be used as little as possible as for many strings this conversion does not make sense
	std::string wstringTostring(const std::wstring& src)
	{
		std::string dst;
		dst.reserve(src.length());
		std::transform(src.begin(), src.end(), std::back_inserter(dst), [](auto c) { return static_cast<char>(c); });
		return dst;
	}

	std::wstring stringTowstring(const std::string& src)
	{
		std::wstring dst;
		dst.reserve(src.length());
		std::transform(
			src.begin(), src.end(), std::back_inserter(dst), [](auto c) { return static_cast<wchar_t>(c & 255); });
		return dst;
	}

	std::wstring LPCTSTRToWstring(LPCTSTR src)
	{
#ifdef UNICODE
		return src ? src : L"";
#else
		return LPCSTRToWstring(src);
#endif
	}

	// Careful calling this with string as it *will* lose embedded null
	// use stringTowstring instead
	std::wstring LPCSTRToWstring(LPCSTR src)
	{
		if (!src) return L"";
		std::string ansi = src;
		return {ansi.begin(), ansi.end()};
	}

	// Converts wstring to LPCWSTR allocated with new
	LPCWSTR wstringToLPCWSTR(const std::wstring& src)
	{
		const auto cch = src.length() + 1;
		const auto cb = cch * sizeof WCHAR;

		const auto lpBin = new (std::nothrow) WCHAR[cch];
		if (lpBin != nullptr)
		{
			memset(lpBin, 0, cb);
			memcpy(lpBin, &src[0], cb);
		}

		return lpBin;
	}

	std::wstring wstringToLower(const std::wstring& src)
	{
		auto dst = src;
		std::transform(src.begin(), src.end(), dst.begin(), towlower);
		return dst;
	}

	bool tryWstringToUlong(ULONG& out, const std::wstring& src, int radix, bool rejectInvalidCharacters)
	{
		// Default our out to 0 for failures
		out = 0;
		if (src.empty()) return false;

		LPWSTR szEndPtr = nullptr;
		out = wcstoul(src.c_str(), &szEndPtr, radix);

		if (rejectInvalidCharacters)
		{
			// if szEndPtr is pointing to something other than NULL, this must be a string, so our conversion failed
			if (!szEndPtr || *szEndPtr)
			{
				out = 0;
				return false;
			}
		}

		return true;
	}

	// Converts a std::wstring to a ulong. Will return 0 if string is empty or contains non-numeric data.
	ULONG wstringToUlong(const std::wstring& src, int radix, bool rejectInvalidCharacters)
	{
		ULONG ulArg{};
		if (tryWstringToUlong(ulArg, src, radix, rejectInvalidCharacters))
		{
			return ulArg;
		}

		return 0;
	}

	// Converts a std::wstring to a long. Will return 0 if string is empty or contains non-numeric data.
	long wstringToLong(const std::wstring& src, int radix)
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

	// Converts a std::wstring to a double. Will return 0 if string is empty or contains non-numeric data.
	double wstringToDouble(const std::wstring& src)
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

	__int64 wstringToInt64(const std::wstring& src)
	{
		if (src.empty()) return 0;

		return _wtoi64(src.c_str());
	}

	std::wstring strip(const std::wstring& str, const std::function<bool(const WCHAR&)>& func)
	{
		std::wstring result;
		result.reserve(str.length());
		remove_copy_if(str.begin(), str.end(), back_inserter(result), func);
		return result;
	}

	std::wstring StripCharacter(const std::wstring& szString, const WCHAR& character)
	{
		return strip(szString, [character](const WCHAR& chr) { return chr == character; });
	}

	std::wstring StripCarriage(const std::wstring& szString) { return StripCharacter(szString, L'\r'); }

	std::wstring StripCRLF(const std::wstring& szString)
	{
		return strip(szString, [](const WCHAR& chr) {
			// Remove carriage returns
			return std::wstring(L", \r\n").find(chr) != std::wstring::npos;
		});
	}

	std::wstring trimWhitespace(const std::wstring& szString)
	{
		static const auto whitespace = {L'\0', L' ', L'\r', L'\n', L'\t'};
		const auto first = szString.find_first_not_of(whitespace);
		if (first == std::string::npos) return emptystring;
		const auto last = szString.find_last_not_of(whitespace);
		return szString.substr(first, last - first + 1);
	}

	std::wstring trim(const std::wstring& szString)
	{
		const auto first = szString.find_first_not_of(' ');
		if (first == std::string::npos) return emptystring;
		const auto last = szString.find_last_not_of(' ');
		return szString.substr(first, last - first + 1);
	}

	std::wstring replace(const std::wstring& str, const std::function<bool(const WCHAR&)>& func, const WCHAR& chr)
	{
		std::wstring result;
		result.reserve(str.length());
		replace_copy_if(str.begin(), str.end(), back_inserter(result), func, chr);
		return result;
	}

	std::wstring ScrubStringForXML(const std::wstring& szString)
	{
		return replace(
			szString,
			[](const WCHAR& chr) {
				// Replace anything less than 0x20 except tab, carriage return and linefeed
				return chr < 0x20 && std::wstring(L"\t\r\n").find(chr) == std::wstring::npos;
			},
			L'.');
	}

	// Processes szFileIn, replacing non file system characters with underscores
	// Do NOT call with full path - just file names
	std::wstring SanitizeFileName(const std::wstring& szFileIn)
	{
		return replace(
			szFileIn,
			[](const WCHAR& chr) {
				// Remove non file system characters
				return std::wstring(L"^&*-+=[]\\|;:\",<>/?\r\n").find(chr) != std::wstring::npos;
			},
			L'_');
	}

	std::wstring indent(int iIndent) { return std::wstring(iIndent, L'\t'); }

	// Find valid printable Unicode characters
	bool InvalidCharacter(ULONG chr, bool bMultiLine)
	{
		// Remove range of control characters
		if (chr >= 0x80 && chr <= 0x9F) return true;
		// Any printable Unicode character gets mapped directly
		if (chr >= 0x20)
		{
			return false;
		}
		// If we allow multiple lines, we accept tab, LF and CR
		else if (
			bMultiLine && (chr == 9 || // Tab
						   chr == 10 || // Line Feed
						   chr == 13)) // Carriage Return
		{
			return false;
		}

		return true;
	}

	std::string RemoveInvalidCharactersA(const std::string& szString, bool bMultiLine)
	{
		if (szString.empty()) return szString;
		auto szBin(szString);
		const auto nullTerminated = szBin.back() == '\0';
		std::replace_if(
			szBin.begin(),
			szBin.end(),
			[bMultiLine](const char& chr) { return InvalidCharacter(0xFF & chr, bMultiLine); },
			'.');

		if (nullTerminated) szBin.back() = '\0';
		return szBin;
	}

	std::wstring RemoveInvalidCharactersW(const std::wstring& szString, bool bMultiLine)
	{
		if (szString.empty()) return szString;
		auto szBin(szString);
		const auto nullTerminated = szBin.back() == L'\0';
		std::replace_if(
			szBin.begin(),
			szBin.end(),
			[bMultiLine](const WCHAR& chr) { return InvalidCharacter(chr, bMultiLine); },
			L'.');

		if (nullTerminated) szBin.back() = L'\0';
		return szBin;
	}

	// Converts binary data to a string, assuming source string was unicode
	std::wstring BinToTextStringW(const std::vector<BYTE>& lpByte, bool bMultiLine)
	{
		SBinary bin = {static_cast<ULONG>(lpByte.size()), const_cast<LPBYTE>(lpByte.data())};
		return BinToTextStringW(&bin, bMultiLine);
	}

	// Converts binary data to a string, assuming source string was unicode
	std::wstring BinToTextStringW(_In_opt_ const SBinary* lpBin, bool bMultiLine)
	{
		if (!lpBin || !lpBin->cb || lpBin->cb % sizeof WCHAR || !lpBin->lpb) return L"";

		const std::wstring szBin(reinterpret_cast<LPWSTR>(lpBin->lpb), lpBin->cb / sizeof WCHAR);
		return RemoveInvalidCharactersW(szBin, bMultiLine);
	}

	std::wstring BinToTextString(const std::vector<BYTE>& lpByte, bool bMultiLine)
	{
		SBinary bin = {static_cast<ULONG>(lpByte.size()), const_cast<LPBYTE>(lpByte.data())};
		return BinToTextString(&bin, bMultiLine);
	}

	// Converts binary data to a string, assuming source string was single byte
	std::wstring BinToTextString(_In_opt_ const SBinary* lpBin, bool bMultiLine)
	{
		if (!lpBin || !lpBin->cb || !lpBin->lpb) return L"";

		std::wstring szBin;
		szBin.reserve(lpBin->cb);

		for (ULONG i = 0; i < lpBin->cb; i++)
		{
			szBin += InvalidCharacter(lpBin->lpb[i], bMultiLine) ? L'.' : lpBin->lpb[i];
		}

		return szBin;
	}

	std::wstring BinToHexString(_In_opt_count_(cb) const BYTE* lpb, size_t cb, bool bPrependCB)
	{
		std::wstring lpsz;

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
				const auto bLow = static_cast<BYTE>(lpb[i] & 0xf);
				const auto bHigh = static_cast<BYTE>(lpb[i] >> 4 & 0xf);
				const auto szLow = static_cast<wchar_t>(bLow <= 0x9 ? L'0' + bLow : L'A' + bLow - 0xa);
				const auto szHigh = static_cast<wchar_t>(bHigh <= 0x9 ? L'0' + bHigh : L'A' + bHigh - 0xa);

				lpsz += szHigh;
				lpsz += szLow;
			}
		}

		return lpsz;
	}

	std::wstring BinToHexString(const std::vector<BYTE>& lpByte, bool bPrependCB)
	{
		SBinary sBin = {static_cast<ULONG>(lpByte.size()), const_cast<LPBYTE>(lpByte.data())};
		return BinToHexString(&sBin, bPrependCB);
	}

	std::wstring BinToHexString(_In_opt_ const SBinary* lpBin, bool bPrependCB)
	{
		if (!lpBin) return L"";

		return BinToHexString(lpBin->lpb, lpBin->cb, bPrependCB);
	}

	bool stripPrefix(std::wstring& str, const std::wstring& prefix)
	{
		const auto length = prefix.length();
		if (str.compare(0, length, prefix) == 0)
		{
			str.erase(0, length);
			return true;
		}

		return false;
	}

	// Converts hex string in lpsz to a binary buffer.
	// If cbTarget != 0, caps the number of bytes converted at cbTarget
	std::vector<BYTE> HexStringToBin(_In_ const std::wstring& input, size_t cbTarget)
	{
		// remove junk
		auto lpsz = strip(input, [](const WCHAR& chr) { return IsFilteredHex(chr); });

		// strip one (and only one) prefix
		stripPrefix(lpsz, L"0x") || stripPrefix(lpsz, L"0X") || stripPrefix(lpsz, L"x") || stripPrefix(lpsz, L"X");

		const auto cchStrLen = lpsz.length();
		// If our input is odd, we can't convert
		// Unless we've capped our output at less than the input
		if (cchStrLen % 2 != 0 && (cbTarget == 0 || cbTarget * 2 > cchStrLen)) return std::vector<BYTE>();

		std::vector<BYTE> lpb;
		WCHAR szTmp[3] = {0};
		size_t iCur = 0;
		size_t cbConverted = 0;

		// convert two characters at a time
		while (iCur + 1 < cchStrLen && (cbTarget == 0 || cbConverted < cbTarget))
		{
			// Check for valid hex characters
			if (lpsz[iCur] > 255 || lpsz[iCur + 1] > 255 || !isxdigit(lpsz[iCur]) || !isxdigit(lpsz[iCur + 1]))
			{
				return std::vector<BYTE>();
			}

			szTmp[0] = lpsz[iCur];
			szTmp[1] = lpsz[iCur + 1];
			lpb.push_back(static_cast<BYTE>(wcstol(szTmp, nullptr, 16)));
			iCur += 2;
			cbConverted++;
		}

		return lpb;
	}

	// Converts vector<BYTE> to LPBYTE allocated with new
	LPBYTE ByteVectorToLPBYTE(const std::vector<BYTE>& bin)
	{
		if (bin.empty()) return nullptr;

		const auto lpBin = new (std::nothrow) BYTE[bin.size()];
		if (lpBin != nullptr)
		{
			memset(lpBin, 0, bin.size());
			memcpy(lpBin, &bin[0], bin.size());
		}

		return lpBin;
	}

	std::vector<std::wstring> split(const std::wstring& str, const wchar_t delim)
	{
		auto ss = std::wstringstream(str);
		std::wstring item;
		std::vector<std::wstring> elems;
		while (getline(ss, item, delim))
		{
			elems.push_back(item);
		}

		return elems;
	}

	std::wstring join(const std::vector<std::wstring>& elems, const std::wstring& delim)
	{
		if (elems.empty()) return emptystring;

		std::wstringstream ss;
		auto iter = elems.begin();
		while (true)
		{
			ss << *iter;
			if (++iter == elems.end()) break;
			ss << delim;
		}

		return ss.str();
	}

	std::wstring join(const std::vector<std::wstring>& elems, const wchar_t delim)
	{
		return join(elems, std::wstring(1, delim));
	}

	// clang-format off
	static const char pBase64[] = {
		0x3e, 0x7f, 0x7f, 0x7f, 0x3f, 0x34, 0x35, 0x36,
		0x37, 0x38, 0x39, 0x3a, 0x3b, 0x3c, 0x3d, 0x7f,
		0x7f, 0x7f, 0x7f, 0x7f, 0x7f, 0x7f, 0x00, 0x01,
		0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09,
		0x0a, 0x0b, 0x0c, 0x0d, 0x0e, 0x0f, 0x10, 0x11,
		0x12, 0x13, 0x14, 0x15, 0x16, 0x17, 0x18, 0x19,
		0x7f, 0x7f, 0x7f, 0x7f, 0x7f, 0x7f, 0x1a, 0x1b,
		0x1c, 0x1d, 0x1e, 0x1f, 0x20, 0x21, 0x22, 0x23,
		0x24, 0x25, 0x26, 0x27, 0x28, 0x29, 0x2a, 0x2b,
		0x2c, 0x2d, 0x2e, 0x2f, 0x30, 0x31, 0x32, 0x33
	};
	// clang-format on

	std::vector<BYTE> Base64Decode(const std::wstring& szEncodedStr)
	{
		const auto cchLen = szEncodedStr.length();
		std::vector<BYTE> lpb;
		if (cchLen % 4) return lpb;

		// look for padding at the end
		const auto posEqual = szEncodedStr.find(L'=');
		if (posEqual != std::wstring::npos)
		{
			auto suffix = szEncodedStr.substr(posEqual);
			if (suffix.length() >= 3 || suffix.find_first_not_of(L'=') != std::wstring::npos) return lpb;
		}

		auto szEncodedStrPtr = szEncodedStr.c_str();
		WCHAR c[4] = {0};
		BYTE bTmp[3] = {0}; // output

		while (*szEncodedStrPtr)
		{
			auto iOutlen = 3;
			for (auto i = 0; i < 4; i++)
			{
				c[i] = *(szEncodedStrPtr + i);
				if (c[i] == L'=')
				{
					iOutlen = i - 1;
					break;
				}

				if (c[i] < 0x2b || c[i] > 0x7a) return std::vector<BYTE>();

				c[i] = pBase64[c[i] - 0x2b];
			}

			bTmp[0] = static_cast<BYTE>(c[0] << 2 | c[1] >> 4);
			bTmp[1] = static_cast<BYTE>((c[1] & 0x0f) << 4 | c[2] >> 2);
			bTmp[2] = static_cast<BYTE>((c[2] & 0x03) << 6 | c[3]);

			for (auto i = 0; i < iOutlen; i++)
			{
				lpb.push_back(bTmp[i]);
			}

			szEncodedStrPtr += 4;
		}

		return lpb;
	}

	// clang-format off
	static const // Base64 Index into encoding
		char pIndex[] = { // and decoding table.
		0x41, 0x42, 0x43, 0x44, 0x45, 0x46, 0x47, 0x48,
		0x49, 0x4a, 0x4b, 0x4c, 0x4d, 0x4e, 0x4f, 0x50,
		0x51, 0x52, 0x53, 0x54, 0x55, 0x56, 0x57, 0x58,
		0x59, 0x5a, 0x61, 0x62, 0x63, 0x64, 0x65, 0x66,
		0x67, 0x68, 0x69, 0x6a, 0x6b, 0x6c, 0x6d, 0x6e,
		0x6f, 0x70, 0x71, 0x72, 0x73, 0x74, 0x75, 0x76,
		0x77, 0x78, 0x79, 0x7a, 0x30, 0x31, 0x32, 0x33,
		0x34, 0x35, 0x36, 0x37, 0x38, 0x39, 0x2b, 0x2f
	};
	// clang-format on

	std::wstring Base64Encode(size_t cbSourceBuf, _In_count_(cbSourceBuf) const BYTE* lpSourceBuffer)
	{
		std::wstring szEncodedStr;
		size_t cbBuf = 0;

		// Using integer division to round down here
		while (cbBuf < cbSourceBuf / 3 * 3) // encode each 3 byte octet.
		{
			szEncodedStr += pIndex[lpSourceBuffer[cbBuf] >> 2];
			szEncodedStr += pIndex[((lpSourceBuffer[cbBuf] & 0x03) << 4) + (lpSourceBuffer[cbBuf + 1] >> 4)];
			szEncodedStr += pIndex[((lpSourceBuffer[cbBuf + 1] & 0x0f) << 2) + (lpSourceBuffer[cbBuf + 2] >> 6)];
			szEncodedStr += pIndex[lpSourceBuffer[cbBuf + 2] & 0x3f];
			cbBuf += 3; // Next octet.
		}

		if (cbSourceBuf - cbBuf != 0) // Partial octet remaining?
		{
			szEncodedStr += pIndex[lpSourceBuffer[cbBuf] >> 2]; // Yes, encode it.

			if (cbSourceBuf - cbBuf == 1) // End of octet?
			{
				szEncodedStr += pIndex[(lpSourceBuffer[cbBuf] & 0x03) << 4];
				szEncodedStr += L'=';
				szEncodedStr += L'=';
			}
			else
			{ // No, one more part.
				szEncodedStr += pIndex[((lpSourceBuffer[cbBuf] & 0x03) << 4) + (lpSourceBuffer[cbBuf + 1] >> 4)];
				szEncodedStr += pIndex[(lpSourceBuffer[cbBuf + 1] & 0x0f) << 2];
				szEncodedStr += L'=';
			}
		}

		return szEncodedStr;
	}

	std::wstring CurrencyToString(const CURRENCY& curVal)
	{
		auto szCur = format(L"%05I64d", curVal.int64); // STRING_OK
		if (szCur.length() > 4)
		{
			szCur.insert(szCur.length() - 4, L"."); // STRING_OK
		}

		return szCur;
	}

	void
	FileTimeToString(_In_ const FILETIME& fileTime, _In_ std::wstring& PropString, _In_opt_ std::wstring& AltPropString)
	{
		PropString = std::wstring{};
		AltPropString = std::wstring{};
		auto SysTime = SYSTEMTIME{};

		if (FileTimeToSystemTime(&fileTime, &SysTime) != 0)
		{
			wchar_t szTimeStr[MAX_PATH] = {};
			wchar_t szDateStr[MAX_PATH] = {};

			// shove millisecond info into our format string since GetTimeFormat doesn't use it
			auto szFormatStr = formatmessage(IDS_FILETIMEFORMAT, SysTime.wMilliseconds);

			GetTimeFormatW(LOCALE_USER_DEFAULT, NULL, &SysTime, szFormatStr.c_str(), szTimeStr, MAX_PATH);
			GetDateFormatW(LOCALE_USER_DEFAULT, NULL, &SysTime, nullptr, szDateStr, MAX_PATH);

			PropString = format(L"%ws %ws", szTimeStr, szDateStr); // STRING_OK
		}

		if (PropString.empty())
		{
			PropString = loadstring(IDS_INVALIDSYSTIME);
		}

		AltPropString = formatmessage(IDS_FILETIMEALTFORMAT, fileTime.dwLowDateTime, fileTime.dwHighDateTime);
	}

	bool IsFilteredHex(const WCHAR& chr)
	{
		const std::wstring szJunk = L"\r\n\t -.,\\/'{}`\""; // STRING_OK
		return szJunk.find(chr) != std::wstring::npos;
	}

	size_t OffsetToFilteredOffset(const std::wstring& szString, size_t offset)
	{
		// An offset in a string can range from 0 (before the first character) to length
		// So for a string such as "AB", valid offsets are 0, 1, 2
		// Example maps for " A B ->"AB":
		// 0->1
		// 1->3
		// 2->4
		if (szString.empty() || offset > szString.length()) return static_cast<size_t>(-1);
		auto found = size_t{};
		auto lastFoundLocation = size_t{};
		for (auto i = size_t{}; i < szString.length(); i++)
		{
			if (!IsFilteredHex(szString[i]))
			{
				found++;
				lastFoundLocation = i;
			}

			if (found == offset + 1) return lastFoundLocation;
		}

		if (found == offset) return lastFoundLocation + 1;
		return static_cast<size_t>(-1);
	}

	// Returns true if str begins with prefix
	bool beginsWith(const std::wstring& str, const std::wstring& prefix)
	{
		if (str.empty()) return false;
		if (prefix.empty()) return false;
		if (prefix.length() > str.length()) return false;

		const auto strLower = wstringToLower(str);
		const auto prefixLower = wstringToLower(prefix);
		const auto res = std::mismatch(prefixLower.begin(), prefixLower.end(), strLower.begin());

		return res.first == prefixLower.end();
	}

	bool endsWith(const std::wstring& str, const std::wstring& ending)
	{
		if (str.length() >= ending.length())
		{
			return 0 == str.compare(str.length() - ending.length(), ending.length(), ending);
		}

		return false;
	}

	std::wstring ensureCRLF(const std::wstring& str)
	{
		if (endsWith(str, L"\r\n"))
		{
			return str;
		}

		return str + L"\r\n";
	}

	// Check that an _SPropValue really contains a string which matches the given ulPropType
	_Check_return_ bool CheckStringProp(_In_opt_ const _SPropValue* lpProp, ULONG ulPropType)
	{
		if (ulPropType != PT_STRING8 && ulPropType != PT_UNICODE)
		{
			output::DebugPrint(
				output::DBGGeneric, L"CheckStringProp: Called with invalid ulPropType of 0x%X\n", ulPropType);
			return false;
		}

		if (!lpProp)
		{
			output::DebugPrint(output::DBGGeneric, L"CheckStringProp: lpProp is NULL\n");
			return false;
		}

		if (PROP_TYPE(lpProp->ulPropTag) == PT_ERROR)
		{
			output::DebugPrint(output::DBGGeneric, L"CheckStringProp: lpProp->ulPropTag is of type PT_ERROR\n");
			return false;
		}

		if (ulPropType != PROP_TYPE(lpProp->ulPropTag))
		{
			output::DebugPrint(
				output::DBGGeneric, L"CheckStringProp: lpProp->ulPropTag is not of type 0x%X\n", ulPropType);
			return false;
		}

		if (lpProp->Value.LPSZ == nullptr)
		{
			output::DebugPrint(output::DBGGeneric, L"CheckStringProp: lpProp->Value.LPSZ is NULL\n");
			return false;
		}

		if (ulPropType == PT_STRING8 && lpProp->Value.lpszA[0] == NULL)
		{
			output::DebugPrint(output::DBGGeneric, L"CheckStringProp: lpProp->Value.lpszA[0] is NULL\n");
			return false;
		}

		if (ulPropType == PT_UNICODE && lpProp->Value.lpszW[0] == NULL)
		{
			output::DebugPrint(output::DBGGeneric, L"CheckStringProp: lpProp->Value.lpszW[0] is NULL\n");
			return false;
		}

		return true;
	}

	std::map<std::wstring, std::wstring> tokenize(const std::wstring str)
	{
		auto ret = std::map<std::wstring, std::wstring>();
		// Start with: a: b c: d
		const auto tokens = strings::split(str, L' ');
		// Now we have a:, b, c:, d
		// Look for a token that ends in :
		// Then pair that with the next token and add to the return
		std::wstring label;
		for (const auto token : tokens)
		{
			if (label.empty())
			{
				if (token.back() == L':')
				{
					label = token.substr(0, token.size() - 1);
				}
			}
			else
			{
				if (token.back() != L':')
				{
					ret[label] = token;
					label.clear();
				}
				else
				{
					label.clear();
				}
			}
		}

		return ret;
	}
} // namespace strings