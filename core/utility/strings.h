#pragma once

namespace strings
{
	// Enable this macro to build with parameter checking for format()
	// Do NOT check in with this macro enabled!
	//#define CHECKFORMATPARAMS

	extern std::wstring emptystring;
	void setTestInstance(HINSTANCE hInstance) noexcept;
	std::wstring loadstring(DWORD dwID);
	std::wstring formatV(LPCWSTR szMsg, va_list argList);
	std::wstring format(LPCWSTR szMsg, ...);
#ifdef CHECKFORMATPARAMS
#undef format
// Yes, this is awkward, but it allows wprintf to be evaluated for format paramater checking without breaking namespaces.
#define format(fmt, ...) format((wprintf(fmt, __VA_ARGS__), fmt), __VA_ARGS__)
#endif

	std::wstring formatmessagesys(DWORD dwID);
	std::wstring formatmessage(DWORD dwID, ...);
	std::wstring formatmessage(LPCWSTR szMsg, ...);
	std::basic_string<TCHAR> wstringTotstring(const std::wstring& src);
	std::string wstringTostring(const std::wstring& src);
	std::wstring stringTowstring(const std::string& src);
	std::wstring LPCTSTRToWstring(LPCTSTR src);
	std::wstring LPCSTRToWstring(LPCSTR src);
	LPCWSTR wstringToLPCWSTR(const std::wstring& src) noexcept;
	std::wstring wstringToLower(const std::wstring& src);
	inline bool compareInsensitive(const std::wstring& lhs, const std::wstring& rhs) noexcept
	{
		return wstringToLower(lhs) == wstringToLower(rhs);
	}
	bool
	tryWstringToUlong(ULONG& out, const std::wstring& src, int radix, bool rejectInvalidCharacters = true) noexcept;
	ULONG wstringToUlong(const std::wstring& src, int radix, bool rejectInvalidCharacters = true) noexcept;
	long wstringToLong(const std::wstring& src, int radix) noexcept;
	double wstringToDouble(const std::wstring& src) noexcept;
	__int64 wstringToInt64(const std::wstring& src) noexcept;
	__int64 wstringToCurrency(const std::wstring& src);

	std::wstring StripCharacter(const std::wstring& szString, const WCHAR& character);
	std::wstring StripCarriage(const std::wstring& szString);
	std::wstring StripCRLF(const std::wstring& szString);
	std::wstring trimWhitespace(const std::wstring& szString);
	std::wstring trim(const std::wstring& szString);
	std::wstring replace(const std::wstring& str, const std::function<bool(const WCHAR&)>& func, const WCHAR& chr);
	std::wstring ScrubStringForXML(const std::wstring& szString);
	std::wstring SanitizeFileName(const std::wstring& szFileIn);
	std::wstring indent(int iIndent);

	std::string RemoveInvalidCharactersA(const std::string& szString, bool bMultiLine = true);
	std::wstring RemoveInvalidCharactersW(const std::wstring& szString, bool bMultiLine = true);
	std::wstring BinToTextStringW(const std::vector<BYTE>& lpByte, bool bMultiLine);
	std::wstring BinToTextStringW(_In_opt_ const SBinary* lpBin, bool bMultiLine);
	std::wstring BinToTextString(const std::vector<BYTE>& lpByte, bool bMultiLine);
	std::wstring BinToTextString(_In_opt_ const SBinary* lpBin, bool bMultiLine);
	std::wstring BinToHexString(const std::vector<BYTE>& lpByte, bool bPrependCB);
	std::wstring BinToHexString(_In_opt_count_(cb) const BYTE* lpb, size_t cb, bool bPrependCB);
	std::wstring BinToHexString(_In_opt_ const SBinary* lpBin, bool bPrependCB);
	bool stripPrefix(std::wstring& str, const std::wstring& prefix);
	std::vector<BYTE> HexStringToBin(_In_ const std::wstring& input, size_t cbTarget = 0);
	LPBYTE ByteVectorToLPBYTE(const std::vector<BYTE>& bin) noexcept;

	std::vector<std::wstring> split(const std::wstring& str, wchar_t delim);
	std::wstring join(const std::vector<std::wstring>& elems, const std::wstring& delim, bool bSkipEmpty = false);
	std::wstring join(const std::vector<std::wstring>& elems, wchar_t delim, bool bSkipEmpty = false);

	// Base64 functions
	std::vector<BYTE> Base64Decode(const std::wstring& szEncodedStr);
	std::wstring Base64Encode(size_t cbSourceBuf, _In_count_(cbSourceBuf) const BYTE* lpSourceBuffer);

	std::wstring CurrencyToString(const CURRENCY& curVal);

	void FileTimeToString(
		_In_ const FILETIME& fileTime,
		_In_ std::wstring& PropString,
		_In_opt_ std::wstring& AltPropString);

	bool IsFilteredHex(const WCHAR& chr);
	size_t OffsetToFilteredOffset(const std::wstring& szString, size_t offset);

	bool beginsWith(const std::wstring& str, const std::wstring& prefix);
	bool endsWith(const std::wstring& str, const std::wstring& ending);
	std::wstring ensureCRLF(const std::wstring& str);

	_Check_return_ bool CheckStringProp(_In_opt_ const _SPropValue* lpProp, ULONG ulPropType);

	// Tokenize strings of the form "a: b c: d" into a map where a->b, c->d, etc
	std::map<std::wstring, std::wstring> tokenize(const std::wstring str);

	std::wstring MAPINAMEIDToString(_In_ const MAPINAMEID& mapiNameId);
	std::wstring collapseTree(const std::wstring& src);
} // namespace strings