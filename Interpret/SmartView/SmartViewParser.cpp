#include <StdAfx.h>
#include <Interpret/SmartView/SmartViewParser.h>

namespace smartview
{
	SmartViewParser::SmartViewParser()
	{
		m_bParsed = false;
		m_bEnableJunk = true;
	}

	void SmartViewParser::Init(size_t cbBin, _In_count_(cbBin) const BYTE* lpBin) { m_Parser.Init(cbBin, lpBin); }

	void SmartViewParser::DisableJunkParsing() { m_bEnableJunk = false; }

	size_t SmartViewParser::GetCurrentOffset() const { return m_Parser.GetCurrentOffset(); }

	void SmartViewParser::EnsureParsed()
	{
		if (m_bParsed || m_Parser.Empty()) return;
		Parse();
		ParseBlocks();

		if (this->hasData() && m_bEnableJunk && m_Parser.RemainingBytes())
		{
			auto junkData = getJunkData();
			addBlock(junkData, JunkDataToString(junkData));
		}

		m_bParsed = true;
	}

	_Check_return_ std::wstring SmartViewParser::ToString()
	{
		if (m_Parser.Empty()) return L"";
		EnsureParsed();

		auto szParsedString = data.ToString();

		if (m_bEnableJunk)
		{
			szParsedString += JunkDataToString(m_Parser.RemainingBytes(), m_Parser.GetCurrentAddress());
		}

		// If we built a string with embedded nulls in it, replace them with dots.
		std::replace_if(
			szParsedString.begin(), szParsedString.end(), [](const WCHAR& chr) { return chr == L'\0'; }, L'.');

		return szParsedString;
	}

	_Check_return_ std::wstring SmartViewParser::JunkDataToString(const std::vector<BYTE>& lpJunkData) const
	{
		if (lpJunkData.empty()) return strings::emptystring;
		output::DebugPrint(
			DBGSmartView,
			L"Had 0x%08X = %u bytes left over.\n",
			static_cast<int>(lpJunkData.size()),
			static_cast<UINT>(lpJunkData.size()));
		auto szJunk = strings::formatmessage(IDS_JUNKDATASIZE, lpJunkData.size());
		szJunk += strings::BinToHexString(lpJunkData, true);
		return szJunk;
	}

	_Check_return_ std::wstring
	SmartViewParser::JunkDataToString(size_t cbJunkData, _In_count_(cbJunkData) const BYTE* lpJunkData) const
	{
		if (!cbJunkData || !lpJunkData) return L"";
		output::DebugPrint(
			DBGSmartView,
			L"Had 0x%08X = %u bytes left over.\n",
			static_cast<int>(cbJunkData),
			static_cast<UINT>(cbJunkData));
		auto szJunk = strings::formatmessage(IDS_JUNKDATASIZE, cbJunkData);
		szJunk += strings::BinToHexString(lpJunkData, static_cast<ULONG>(cbJunkData), true);
		return szJunk;
	}
}