#pragma once
#include <string>
#include <vector>
using namespace std;

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
};

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
wstring formatmessage(wstring const szMsg, ...);
tstring wstringTotstring(wstring const& src);
string wstringTostring(wstring const& src);
wstring stringTowstring(string const& src);
wstring LPCTSTRToWstring(LPCTSTR src);
wstring LPCSTRToWstring(LPCSTR src);
wstring wstringToLower(wstring const& src);
ULONG wstringToUlong(wstring const& src, int radix, bool rejectInvalidCharacters = true);
long wstringToLong(wstring const& src, int radix);
double wstringToDouble(wstring const& src);
__int64 wstringToInt64(wstring const& src);

wstring StripCharacter(wstring szString, WCHAR character);
wstring StripCarriage(wstring const& szString);
wstring CleanString(wstring szString);
wstring ScrubStringForXML(wstring szString);
wstring SanitizeFileNameW(wstring szFileIn);
wstring indent(int iIndent);

wstring BinToTextString(_In_ LPSBinary lpBin, bool bMultiLine);
wstring BinToHexString(_In_opt_count_(cb) LPBYTE lpb, size_t cb, bool bPrependCB);
wstring BinToHexString(_In_opt_ LPSBinary lpBin, bool bPrependCB);
vector<BYTE> HexStringToBin(_In_ wstring lpsz, size_t cbTarget = 0);
LPBYTE ByteVectorToLPBYTE(vector<BYTE> const& bin);

vector<wstring> split(const wstring &str, const wchar_t delim);
wstring join(const vector<wstring> elems, const wchar_t delim);