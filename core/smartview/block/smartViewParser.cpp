#include <core/stdafx.h>
#include <core/smartview/block/smartViewParser.h>
#include <core/utility/strings.h>
#include <core/smartview/block/blockBytes.h>

namespace smartview
{
	void smartViewParser::ensureParsed()
	{
		if (parsed || m_Parser->empty()) return;
		const auto startOffset = m_Parser->getOffset();

		parse();
		parseBlocks();

		if (this->hasData() && enableJunk && m_Parser->getSize())
		{
			auto junkData = blockBytes::parse(m_Parser, m_Parser->getSize());
			terminateBlock();
			addHeader(L"Unparsed data size = 0x%1!08X!\r\n", junkData->size());
			addChild(junkData);
		}

		const auto endOffset = m_Parser->getOffset();

		// Ensure we are tagged with offset and size so all top level blocks get proper highlighting
		setOffset(startOffset);
		setSize(endOffset - startOffset);

		parsed = true;
	}

	std::wstring smartViewParser::toString()
	{
		if (m_Parser->empty()) return L"";
		ensureParsed();

		auto parsedString = strings::trimWhitespace(block::toString());

		// If we built a string with embedded nulls in it, replace them with dots.
		std::replace_if(
			parsedString.begin(), parsedString.end(), [](const WCHAR& chr) { return chr == L'\0'; }, L'.');

		return parsedString;
	}
} // namespace smartview