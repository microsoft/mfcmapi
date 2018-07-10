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
	std::wstring loadstring(DWORD dwID);
	std::wstring formatV(LPCWSTR szMsg, va_list argList);
	std::wstring format(LPCWSTR szMsg, ...);
#ifdef CHECKFORMATPARAMS
#undef format
#define format(fmt, ...) (wprintf(fmt, __VA_ARGS__), format(fmt, __VA_ARGS__))
#endif

	std::wstring formatmessagesys(DWORD dwID);
	std::wstring formatmessage(DWORD dwID, ...);
	std::wstring formatmessage(LPCWSTR szMsg, ...);
	tstring wstringTotstring(const std::wstring& src);
	std::string wstringTostring(const std::wstring& src);
	std::wstring stringTowstring(const std::string& src);
	std::wstring LPCTSTRToWstring(LPCTSTR src);
	std::wstring LPCSTRToWstring(LPCSTR src);
	LPCWSTR wstringToLPCWSTR(const std::wstring& src);
	std::wstring wstringToLower(const std::wstring& src);
	ULONG wstringToUlong(const std::wstring& src, int radix, bool rejectInvalidCharacters = true);
	long wstringToLong(const std::wstring& src, int radix);
	double wstringToDouble(const std::wstring& src);
	__int64 wstringToInt64(const std::wstring& src);

	std::wstring StripCharacter(const std::wstring& szString, const WCHAR& character);
	std::wstring StripCarriage(const std::wstring& szString);
	std::wstring StripCRLF(const std::wstring& szString);
	std::wstring trim(const std::wstring& szString);
	std::wstring replace(const std::wstring& str, const std::function<bool(const WCHAR&)>& func, const WCHAR& chr);
	std::wstring ScrubStringForXML(const std::wstring& szString);
	std::wstring SanitizeFileName(const std::wstring& szFileIn);
	std::wstring indent(int iIndent);

	std::string RemoveInvalidCharactersA(const std::string& szString, bool bMultiLine = true);
	std::wstring RemoveInvalidCharactersW(const std::wstring& szString, bool bMultiLine = true);
	std::wstring BinToTextStringW(const std::vector<BYTE>& lpByte, bool bMultiLine);
	std::wstring BinToTextStringW(_In_ const SBinary* lpBin, bool bMultiLine);
	std::wstring BinToTextString(const std::vector<BYTE>& lpByte, bool bMultiLine);
	std::wstring BinToTextString(_In_ const SBinary* lpBin, bool bMultiLine);
	std::wstring BinToHexString(const std::vector<BYTE>& lpByte, bool bPrependCB);
	std::wstring BinToHexString(_In_opt_count_(cb) const BYTE* lpb, size_t cb, bool bPrependCB);
	std::wstring BinToHexString(_In_opt_ const SBinary* lpBin, bool bPrependCB);
	std::vector<BYTE> HexStringToBin(_In_ const std::wstring& input, size_t cbTarget = 0);
	LPBYTE ByteVectorToLPBYTE(const std::vector<BYTE>& bin);

	std::vector<std::wstring> split(const std::wstring& str, wchar_t delim);
	std::wstring join(const std::vector<std::wstring>& elems, const std::wstring& delim);
	std::wstring join(const std::vector<std::wstring>& elems, wchar_t delim);

	// Base64 functions
	std::vector<BYTE> Base64Decode(const std::wstring& szEncodedStr);
	std::wstring Base64Encode(size_t cbSourceBuf, _In_count_(cbSourceBuf) const BYTE* lpSourceBuffer);

	std::wstring CurrencyToString(const CURRENCY& curVal);

	void FileTimeToString(
		_In_ const FILETIME& fileTime,
		_In_ std::wstring& PropString,
		_In_opt_ std::wstring& AltPropString);
}