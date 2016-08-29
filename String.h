#pragma once
#include <string>
#include <vector>
using namespace std;

extern wstring emptystring; 
wstring loadstring(DWORD dwID);
wstring formatV(wstring const& szMsg, va_list argList);
wstring format(LPCWSTR szMsg, ...);
wstring formatmessagesys(DWORD dwID);
wstring formatmessage(DWORD dwID, ...);
wstring formatmessage(wstring const szMsg, ...);
LPTSTR wstringToLPTSTR(wstring const& src);
CString wstringToCString(wstring const& src);
CStringA wstringToCStringA(wstring const& src);
string wstringTostring(wstring const& src);
wstring LPCTSTRToWstring(LPCTSTR src);
wstring LPCSTRToWstring(LPCSTR src);
CString LPCSTRToCString(LPCSTR src);
CStringA LPCTSTRToCStringA(LPCTSTR src);
void wstringToLower(wstring src);
ULONG wstringToUlong(wstring const& src, int radix, bool rejectInvalidCharacters = true);
long wstringToLong(wstring const& src, int radix);
double wstringToDouble(wstring const& src);
__int64 wstringToInt64(wstring const& src);

wstring StripCharacter(wstring szString, WCHAR character);
wstring StripCarriage(wstring szString);
wstring CleanString(wstring szString);
void CleanPropString(_In_ CString* lpString);

// Unicode support
_Check_return_ HRESULT AnsiToUnicode(_In_opt_z_ LPCSTR pszA, _Out_z_cap_(cchszA) LPWSTR* ppszW, _Out_ size_t* cchszW, size_t cchszA);
_Check_return_ HRESULT AnsiToUnicode(_In_opt_z_ LPCSTR pszA, _Out_z_cap_(cchszA) LPWSTR* ppszW, size_t cchszA = -1);
_Check_return_ HRESULT UnicodeToAnsi(_In_z_ LPCWSTR pszW, _Out_z_cap_(cchszW) LPSTR* ppszA, size_t cchszW = -1);

bool IsNullOrEmptyW(LPCWSTR szStr);
bool IsNullOrEmptyA(LPCSTR szStr);

#ifdef UNICODE
#define IsNullOrEmpty IsNullOrEmptyW
#else
#define IsNullOrEmpty IsNullOrEmptyA
#endif

wstring BinToTextString(_In_ LPSBinary lpBin, bool bMultiLine);
wstring BinToHexString(_In_opt_count_(cb) LPBYTE lpb, size_t cb, bool bPrependCB);
wstring BinToHexString(_In_opt_ LPSBinary lpBin, bool bPrependCB);
vector<BYTE> HexStringToBin(_In_ wstring lpsz, size_t cbTarget = 0);
LPBYTE ByteVectorToLPBYTE(vector<BYTE>& bin);