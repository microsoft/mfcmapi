#include <core/stdafx.h>
#include <core/smartview/SmartViewParser.h>
#include <core/utility/strings.h>
#include <core/smartview/block/blockBytes.h>

namespace smartview
{
	void SmartViewParser::EnsureParsed()
	{
		if (m_bParsed || m_Parser->empty()) return;
		Parse();
		ParseBlocks();

		if (this->hasData() && m_bEnableJunk && m_Parser->RemainingBytes())
		{
			auto junkData = std::make_shared<blockBytes>(m_Parser, m_Parser->RemainingBytes());
			terminateBlock();
			addHeader(L"Unparsed data size = 0x%1!08X!\r\n", junkData->size());
			addChild(junkData);
		}

		m_bParsed = true;
	}

	std::wstring SmartViewParser::toString()
	{
		if (m_Parser->empty()) return L"";
		EnsureParsed();

		auto szParsedString = strings::trimWhitespace(data->toString());

		// If we built a string with embedded nulls in it, replace them with dots.
		std::replace_if(
			szParsedString.begin(), szParsedString.end(), [](const WCHAR& chr) { return chr == L'\0'; }, L'.');

		return szParsedString;
	}
} // namespace smartview