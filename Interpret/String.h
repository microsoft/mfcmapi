#pragma once
#include <string>
#include <vector>

// Enable this macro to build with parameter checking for format()
// Do NOT check in with this macro enabled!
//#define CHECKFORMATPARAMS

namespace std
{
#ifdef _UNICODE 
	typedef wstring tstring;
#else
	typedef string tstring;
#endif
}

using std::wstring;
using std::string;
using std::tstring;
using std::vector;

extern wstring emptystring;
wstring loadstring(DWORD dwID);
wstring formatV(LPCWSTR szMsg, va_list argList);
wstring format(LPCWSTR szMsg, ...);
#ifdef CHECKFORMATPARAMS
#undef format
#define format(fmt,...) (wprintf(fmt,__VA_ARGS__), format(fmt,__VA_ARGS__))
#endif

wstring formatmessagesys(DWORD dwID);
wstring formatmessage(DWORD dwID, ...);
wstring formatmessage(const wstring szMsg, ...);
tstring wstringTotstring(const wstring& src);
string wstringTostring(const wstring& src);
wstring stringTowstring(const string& src);
wstring LPCTSTRToWstring(LPCTSTR src);
wstring LPCSTRToWstring(LPCSTR src);
wstring wstringToLower(const wstring& src);
ULONG wstringToUlong(const wstring& src, int radix, bool rejectInvalidCharacters = true);
long wstringToLong(const wstring& src, int radix);
double wstringToDouble(const wstring& src);
__int64 wstringToInt64(const wstring& src);

wstring StripCharacter(const wstring& szString, const WCHAR& character);
wstring StripCarriage(const wstring& szString);
wstring CleanString(const wstring& szString);
wstring TrimString(const wstring& szString);
wstring ScrubStringForXML(const wstring& szString);
wstring SanitizeFileNameW(const wstring& szFileIn);
wstring indent(int iIndent);

wstring BinToTextStringW(const vector<BYTE>& lpByte, bool bMultiLine);
wstring BinToTextStringW(_In_ const LPSBinary lpBin, bool bMultiLine);
wstring BinToTextString(_In_ const LPSBinary lpBin, bool bMultiLine);
wstring BinToHexString(const vector<BYTE>& lpByte, bool bPrependCB);
wstring BinToHexString(_In_opt_count_(cb) const BYTE* lpb, size_t cb, bool bPrependCB);
wstring BinToHexString(_In_opt_ const SBinary* lpBin, bool bPrependCB);
vector<BYTE> HexStringToBin(_In_ const wstring& input, size_t cbTarget = 0);
LPBYTE ByteVectorToLPBYTE(const vector<BYTE>& bin);

vector<wstring> split(const wstring& str, const wchar_t delim);
wstring join(const vector<wstring>& elems, const wstring& delim);
wstring join(const vector<wstring>& elems, const wchar_t delim);