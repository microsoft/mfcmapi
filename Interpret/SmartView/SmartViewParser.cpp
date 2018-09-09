#include <StdAfx.h>
#include <Interpret/SmartView/SmartViewParser.h>

namespace smartview
{
	SmartViewParser::SmartViewParser()
	{
		m_bParsed = false;
		m_bEnableJunk = true;
	}

	void SmartViewParser::init(size_t cbBin, _In_count_(cbBin) const BYTE* lpBin) { m_Parser.init(cbBin, lpBin); }

	void SmartViewParser::DisableJunkParsing() { m_bEnableJunk = false; }

	size_t SmartViewParser::GetCurrentOffset() const { return m_Parser.GetCurrentOffset(); }

	void SmartViewParser::EnsureParsed()
	{
		if (m_bParsed || m_Parser.empty()) return;
		Parse();
		ParseBlocks();

		if (this->hasData() && m_bEnableJunk && m_Parser.RemainingBytes())
		{
			const auto junkData = m_Parser.GetRemainingData();

			addLine();
			addHeader(L"Unparsed data size = 0x%1!08X!\r\n", junkData.size());
			addBlock(junkData);
		}

		m_bParsed = true;
	}

	_Check_return_ std::wstring SmartViewParser::ToString()
	{
		if (m_Parser.empty()) return L"";
		EnsureParsed();

		auto szParsedString = strings::trimWhitespace(data.ToString());

		// If we built a string with embedded nulls in it, replace them with dots.
		std::replace_if(
			szParsedString.begin(), szParsedString.end(), [](const WCHAR& chr) { return chr == L'\0'; }, L'.');

		return szParsedString;
	}
} // namespace smartview