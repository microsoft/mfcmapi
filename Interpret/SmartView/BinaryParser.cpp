#include <StdAfx.h>
#include <Interpret/SmartView/BinaryParser.h>

namespace smartview
{
	void block::addBlockBytes(blockBytes child)
	{
		auto _block = child;
		_block.text = strings::BinToHexString(child, true);
		children.push_back(_block);
	}

	static std::wstring CLASS = L"CBinaryParser";

	CBinaryParser::CBinaryParser() { m_Offset = 0; }

	CBinaryParser::CBinaryParser(size_t cbBin, _In_count_(cbBin) const BYTE* lpBin) { Init(cbBin, lpBin); }

	void CBinaryParser::Init(size_t cbBin, _In_count_(cbBin) const BYTE* lpBin)
	{
		output::DebugPrintEx(
			DBGSmartView, CLASS, L"Init", L"cbBin = 0x%08X = %u\n", static_cast<int>(cbBin), static_cast<UINT>(cbBin));
		m_Bin = lpBin && cbBin ? std::vector<BYTE>(lpBin, lpBin + cbBin) : std::vector<BYTE>();
		m_Offset = 0;
	}

	bool CBinaryParser::Empty() const { return m_Bin.empty(); }

	void CBinaryParser::Advance(size_t cbAdvance) { m_Offset += cbAdvance; }

	void CBinaryParser::Rewind()
	{
		output::DebugPrintEx(DBGSmartView, CLASS, L"Rewind", L"Rewinding to the beginning of the stream\n");
		m_Offset = 0;
	}

	size_t CBinaryParser::GetCurrentOffset() const { return m_Offset; }

	const BYTE* CBinaryParser::GetCurrentAddress() const { return m_Bin.data() + m_Offset; }

	void CBinaryParser::SetCurrentOffset(size_t stOffset)
	{
		output::DebugPrintEx(
			DBGSmartView,
			CLASS,
			L"SetCurrentOffset",
			L"Setting offset 0x%08X = %u bytes.\n",
			static_cast<int>(stOffset),
			static_cast<UINT>(stOffset));
		m_Offset = stOffset;
	}

	// If we're before the end of the buffer, return the count of remaining bytes
	// If we're at or past the end of the buffer, return 0
	// If we're before the beginning of the buffer, return 0
	size_t CBinaryParser::RemainingBytes() const
	{
		if (m_Offset > m_Bin.size()) return 0;
		return m_Bin.size() - m_Offset;
	}

	bool CBinaryParser::CheckRemainingBytes(size_t cbBytes) const
	{
		const auto cbRemaining = RemainingBytes();
		if (cbBytes > cbRemaining)
		{
			output::DebugPrintEx(
				DBGSmartView,
				CLASS,
				L"CheckRemainingBytes",
				L"Bytes requested (0x%08X = %u) > remaining bytes (0x%08X = %u)\n",
				static_cast<int>(cbBytes),
				static_cast<UINT>(cbBytes),
				static_cast<int>(cbRemaining),
				static_cast<UINT>(cbRemaining));
			output::DebugPrintEx(
				DBGSmartView, CLASS, L"CheckRemainingBytes", L"Total Bytes: 0x%08X = %u\n", m_Bin.size(), m_Bin.size());
			output::DebugPrintEx(
				DBGSmartView, CLASS, L"CheckRemainingBytes", L"Current offset: 0x%08X = %d\n", m_Offset, m_Offset);
			return false;
		}

		return true;
	}
}