#include <StdAfx.h>
#include <Interpret/SmartView/BinaryParser.h>

namespace smartview
{
	static std::wstring CLASS = L"CBinaryParser";

	void CBinaryParser::Init(size_t cbBin, _In_count_(cbBin) const BYTE* lpBin)
	{
		m_Bin = lpBin && cbBin ? std::vector<BYTE>(lpBin, lpBin + cbBin) : std::vector<BYTE>();
		m_Offset = 0;
	}

	void CBinaryParser::Rewind() { m_Offset = 0; }

	void CBinaryParser::SetCurrentOffset(size_t stOffset) { m_Offset = stOffset; }

	// If we're before the end of the buffer, return the count of remaining bytes
	// If we're at or past the end of the buffer, return 0
	// If we're before the beginning of the buffer, return 0
	size_t CBinaryParser::RemainingBytes() const
	{
		if (m_Offset > m_Bin.size()) return 0;
		return m_Bin.size() - m_Offset;
	}

	bool CBinaryParser::CheckRemainingBytes(size_t cbBytes) const { return cbBytes <= RemainingBytes(); }
} // namespace smartview