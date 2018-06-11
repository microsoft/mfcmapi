#include <StdAfx.h>
#include <Interpret/SmartView/BinaryParser.h>

namespace smartview
{
	static std::wstring CLASS = L"CBinaryParser";

	CBinaryParser::CBinaryParser()
	{
		m_Offset = 0;
	}

	CBinaryParser::CBinaryParser(size_t cbBin, _In_count_(cbBin) const BYTE* lpBin)
	{
		Init(cbBin, lpBin);
	}

	void CBinaryParser::Init(size_t cbBin, _In_count_(cbBin) const BYTE* lpBin)
	{
		output::DebugPrintEx(DBGSmartView, CLASS, L"Init", L"cbBin = 0x%08X = %u\n", static_cast<int>(cbBin), static_cast<UINT>(cbBin));
		m_Bin = lpBin && cbBin ? std::vector<BYTE>(lpBin, lpBin + cbBin) : std::vector<BYTE>();
		m_Offset = 0;
	}

	bool CBinaryParser::Empty() const
	{
		return m_Bin.empty();
	}

	void CBinaryParser::Advance(size_t cbAdvance)
	{
		m_Offset += cbAdvance;
	}

	void CBinaryParser::Rewind()
	{
		output::DebugPrintEx(DBGSmartView, CLASS, L"Rewind", L"Rewinding to the beginning of the stream\n");
		m_Offset = 0;
	}

	size_t CBinaryParser::GetCurrentOffset() const
	{
		return m_Offset;
	}

	const BYTE* CBinaryParser::GetCurrentAddress() const
	{
		return m_Bin.data() + m_Offset;
	}

	void CBinaryParser::SetCurrentOffset(size_t stOffset)
	{
		output::DebugPrintEx(DBGSmartView, CLASS, L"SetCurrentOffset", L"Setting offset 0x%08X = %u bytes.\n", static_cast<int>(stOffset), static_cast<UINT>(stOffset));
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
			output::DebugPrintEx(DBGSmartView, CLASS, L"CheckRemainingBytes", L"Bytes requested (0x%08X = %u) > remaining bytes (0x%08X = %u)\n", static_cast<int>(cbBytes), static_cast<UINT>(cbBytes), static_cast<int>(cbRemaining), static_cast<UINT>(cbRemaining));
			output::DebugPrintEx(DBGSmartView, CLASS, L"CheckRemainingBytes", L"Total Bytes: 0x%08X = %u\n", m_Bin.size(), m_Bin.size());
			output::DebugPrintEx(DBGSmartView, CLASS, L"CheckRemainingBytes", L"Current offset: 0x%08X = %d\n", m_Offset, m_Offset);
			return false;
		}

		return true;
	}

	std::string CBinaryParser::GetStringA(size_t cchChar)
	{
		if (cchChar == -1)
		{
			const auto hRes = StringCchLengthA(reinterpret_cast<LPCSTR>(GetCurrentAddress()), (m_Bin.size() - m_Offset) / sizeof CHAR, &cchChar);
			if (FAILED(hRes)) return "";
			cchChar += 1;
		}

		if (!cchChar || !CheckRemainingBytes(sizeof CHAR * cchChar)) return "";
		const std::string ret(reinterpret_cast<LPCSTR>(GetCurrentAddress()), cchChar);
		m_Offset += sizeof CHAR * cchChar;
		return strings::RemoveInvalidCharactersA(ret);
	}

	std::wstring CBinaryParser::GetStringW(size_t cchChar)
	{
		if (cchChar == -1)
		{
			const auto hRes = StringCchLengthW(reinterpret_cast<LPCWSTR>(GetCurrentAddress()), (m_Bin.size() - m_Offset) / sizeof WCHAR, &cchChar);
			if (FAILED(hRes)) return strings::emptystring;
			cchChar += 1;
		}

		if (!cchChar || !CheckRemainingBytes(sizeof WCHAR * cchChar)) return strings::emptystring;
		const std::wstring ret(reinterpret_cast<LPCWSTR>(GetCurrentAddress()), cchChar);
		m_Offset += sizeof WCHAR * cchChar;
		return strings::RemoveInvalidCharactersW(ret);
	}

	std::vector<BYTE> CBinaryParser::GetBYTES(size_t cbBytes, size_t cbMaxBytes)
	{
		if (!cbBytes || !CheckRemainingBytes(cbBytes)) return std::vector<BYTE>();
		if (cbMaxBytes != -1 && cbBytes > cbMaxBytes) return std::vector<BYTE>();
		std::vector<BYTE> ret(const_cast<LPBYTE>(GetCurrentAddress()), const_cast<LPBYTE>(GetCurrentAddress() + cbBytes));
		m_Offset += cbBytes;
		return ret;
	}

	std::vector<BYTE> CBinaryParser::GetRemainingData()
	{
		return GetBYTES(RemainingBytes());
	}
}