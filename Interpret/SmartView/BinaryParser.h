#pragma once

// CBinaryParser - helper class for parsing binary data without
// worrying about whether you've run off the end of your buffer.
class CBinaryParser
{
public:
	CBinaryParser();
	CBinaryParser(size_t cbBin, _In_count_(cbBin) const BYTE* lpBin);

	bool Empty() const;
	void Init(size_t cbBin, _In_count_(cbBin) const BYTE* lpBin);
	void Advance(size_t cbAdvance);
	void Rewind();
	size_t GetCurrentOffset() const;
	const BYTE* GetCurrentAddress() const;
	// Moves the parser to an offset obtained from GetCurrentOffset
	void SetCurrentOffset(size_t stOffset);
	size_t RemainingBytes() const;
	template <typename T> T Get()
	{
		if (!CheckRemainingBytes(sizeof T)) return T();
		auto ret = *reinterpret_cast<const T *>(GetCurrentAddress());
		m_Offset += sizeof T;
		return ret;
	}

	string GetStringA(size_t cchChar = -1);
	wstring GetStringW(size_t cchChar = -1);
	vector<BYTE> GetBYTES(size_t cbBytes, size_t cbMaxBytes = -1);
	vector<BYTE> GetRemainingData();

private:
	bool CheckRemainingBytes(size_t cbBytes) const;
	vector<BYTE> m_Bin;
	size_t m_Offset;
};