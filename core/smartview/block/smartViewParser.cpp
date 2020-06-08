#include <core/stdafx.h>
#include <core/smartview/block/block.h>
#include <core/utility/strings.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/scratchBlock.h>

namespace smartview
{
	std::wstring smartViewParser::toString()
	{
		if (!m_Parser || m_Parser->empty()) return toBlockString();
		ensureParsed();

		auto parsedString = strings::trimWhitespace(toBlockString());

		// If we built a string with embedded nulls in it, replace them with dots.
		std::replace_if(
			parsedString.begin(), parsedString.end(), [](const WCHAR& chr) { return chr == L'\0'; }, L'.');

		return parsedString;
	}
} // namespace smartview