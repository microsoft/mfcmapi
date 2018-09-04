#pragma once
#include <string>
#include <vector>
#include <functional>

namespace strings
{
// Enable this macro to build with parameter checking for format()
// Do NOT check in with this macro enabled!
//#define CHECKFORMATPARAMS

#ifdef _UNICODE
	typedef std::wstring tstring;
#else
	typedef std::string tstring;
#endif

	extern std::wstring emptystring;
	void setTestInstance(HINSTANCE hInstance);
	const std::wstring loadstring(DWORD dwID);
	const std::wstring formatV(LPCWSTR szMsg, va_list argList);
	const std::wstring format(LPCWSTR szMsg, ...);
#ifdef CHECKFORMATPARAMS
#undef format
#define format(fmt, ...) (wprintf(fmt, __VA_ARGS__), format(fmt, __VA_ARGS__))
#endif

	const std::wstring formatmessagesys(DWORD dwID);
	const std::wstring formatmessage(DWORD dwID, ...);
	const std::wstring formatmessage(LPCWSTR szMsg, ...);
	const tstring wstringTotstring(const std::wstring& src);
	const std::string wstringTostring(const std::wstring& src);
	const std::wstring stringTowstring(const std::string& src);
	const std::wstring LPCTSTRToWstring(LPCTSTR src);
	const std::wstring LPCSTRToWstring(LPCSTR src);
	LPCWSTR wstringToLPCWSTR(const std::wstring& src);
	const std::wstring wstringToLower(const std::wstring& src);
	ULONG wstringToUlong(const std::wstring& src, int radix, bool rejectInvalidCharacters = true);
	long wstringToLong(const std::wstring& src, int radix);
	double wstringToDouble(const std::wstring& src);
	__int64 wstringToInt64(const std::wstring& src);

	const std::wstring StripCharacter(const std::wstring& szString, const WCHAR& character);
	const std::wstring StripCarriage(const std::wstring& szString);
	const std::wstring StripCRLF(const std::wstring& szString);
	const std::wstring trimWhitespace(const std::wstring& szString);
	const std::wstring trim(const std::wstring& szString);
	const std::wstring replace(const std::wstring& str, const std::function<bool(const WCHAR&)>& func, const WCHAR& chr);
	const std::wstring ScrubStringForXML(const std::wstring& szString);
	const std::wstring SanitizeFileName(const std::wstring& szFileIn);
	const std::wstring indent(int iIndent);

	const std::string RemoveInvalidCharactersA(const std::string& szString, bool bMultiLine = true);
	const std::wstring RemoveInvalidCharactersW(const std::wstring& szString, bool bMultiLine = true);
	const std::wstring BinToTextStringW(const std::vector<BYTE>& lpByte, bool bMultiLine);
	const std::wstring BinToTextStringW(_In_ const SBinary* lpBin, bool bMultiLine);
	const std::wstring BinToTextString(const std::vector<BYTE>& lpByte, bool bMultiLine);
	const std::wstring BinToTextString(_In_ const SBinary* lpBin, bool bMultiLine);
	const std::wstring BinToHexString(const std::vector<BYTE>& lpByte, bool bPrependCB);
	const std::wstring BinToHexString(_In_opt_count_(cb) const BYTE* lpb, size_t cb, bool bPrependCB);
	const std::wstring BinToHexString(_In_opt_ const SBinary* lpBin, bool bPrependCB);
	const std::vector<BYTE> HexStringToBin(_In_ const std::wstring& input, size_t cbTarget = 0);
	LPBYTE ByteVectorToLPBYTE(const std::vector<BYTE>& bin);

	const std::vector<std::wstring> split(const std::wstring& str, wchar_t delim);
	const std::wstring join(const std::vector<std::wstring>& elems, const std::wstring& delim);
	const std::wstring join(const std::vector<std::wstring>& elems, wchar_t delim);

	// Base64 functions
	const std::vector<BYTE> Base64Decode(const std::wstring& szEncodedStr);
	const std::wstring Base64Encode(size_t cbSourceBuf, _In_count_(cbSourceBuf) const BYTE* lpSourceBuffer);

	const std::wstring CurrencyToString(const CURRENCY& curVal);

	void FileTimeToString(
		_In_ const FILETIME& fileTime,
		_In_ std::wstring& PropString,
		_In_opt_ std::wstring& AltPropString);
}