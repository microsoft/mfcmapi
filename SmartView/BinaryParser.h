#pragma once

// CBinaryParser - helper class for parsing binary data without
// worrying about whether you've run off the end of your buffer.
class CBinaryParser
{
public:
	CBinaryParser();
	CBinaryParser(size_t cbBin, _In_count_(cbBin) LPBYTE lpBin);

	bool Empty();
	void Init(size_t cbBin, _In_count_(cbBin) LPBYTE lpBin);
	void Advance(size_t cbAdvance);
	void Rewind();
	size_t GetCurrentOffset();
	LPBYTE GetCurrentAddress();
	// Moves the parser to an offset obtained from GetCurrentOffset
	void SetCurrentOffset(size_t stOffset);
	size_t RemainingBytes();
	void GetBYTE(_Out_ BYTE* pBYTE);
	void GetWORD(_Out_ WORD* pWORD);
	void GetDWORD(_Out_ DWORD* pDWORD);
	void GetLARGE_INTEGER(_Out_ LARGE_INTEGER* pLARGE_INTEGER);
	void GetBYTES(size_t cbBytes, size_t cbMaxBytes, _Out_ LPBYTE* ppBYTES);
	void GetBYTESNoAlloc(size_t cbBytes, size_t cbMaxBytes, _In_count_(cbBytes) LPBYTE pBYTES);
	void GetStringA(size_t cchChar, _Deref_out_z_ LPSTR* ppStr);
	void GetStringW(size_t cchChar, _Deref_out_z_ LPWSTR* ppStr);
	void GetStringA(_Deref_out_opt_z_ LPSTR* ppStr);
	void GetStringW(_Deref_out_opt_z_ LPWSTR* ppStr);
	size_t GetRemainingData(_Out_ LPBYTE* ppRemainingBYTES);

private:
	bool CheckRemainingBytes(size_t cbBytes);
	size_t m_cbBin;
	LPBYTE m_lpBin;
	LPBYTE m_lpEnd;
	LPBYTE m_lpCur;
};