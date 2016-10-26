#pragma once

// CBinaryParser - helper class for parsing binary data without
// worrying about whether you've run off the end of your buffer.
class CBinaryParser
{
public:
	CBinaryParser();
	CBinaryParser(size_t cbBin, _In_count_(cbBin) LPBYTE lpBin);

	bool Empty() const;
	void Init(size_t cbBin, _In_count_(cbBin) LPBYTE lpBin);
	void Advance(size_t cbAdvance);
	void Rewind();
	size_t GetCurrentOffset() const;
	LPBYTE GetCurrentAddress() const;
	// Moves the parser to an offset obtained from GetCurrentOffset
	void SetCurrentOffset(size_t stOffset);
	size_t RemainingBytes() const;
	void GetBYTE(_Out_ BYTE* pBYTE);
	void GetWORD(_Out_ WORD* pWORD);
	void GetDWORD(_Out_ DWORD* pDWORD);
	template <typename T> T Get()
	{
		if (!CheckRemainingBytes(sizeof T)) return T();
		auto ret = *reinterpret_cast<T *>(m_lpCur);
		m_lpCur += sizeof T;
		return ret;
	}

	string GetStringA(size_t cchChar = -1);
	wstring GetStringW(size_t cchChar = -1);
	vector<BYTE> GetBYTES(size_t cbBytes, size_t cbMaxBytes = -1);
	vector<BYTE> GetRemainingData();

private:
	bool CheckRemainingBytes(size_t cbBytes) const;
	size_t m_cbBin;
	LPBYTE m_lpBin;
	LPBYTE m_lpEnd;
	LPBYTE m_lpCur;
};