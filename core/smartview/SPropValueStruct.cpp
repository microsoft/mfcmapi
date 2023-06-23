#include <core/stdafx.h>
#include <core/smartview/SPropValueStruct.h>
#include <core/interpret/proptags.h>
#include <core/interpret/proptype.h>

namespace smartview
{
	void SPropValueStruct::parse()
	{
		const auto ulCurrOffset = parser->getOffset(); 

		PropType = blockT<WORD>::parse(parser);
		PropID = blockT<WORD>::parse(parser);

		ulPropTag = blockT<ULONG>::create(
			PROP_TAG(*PropType, *PropID), PropType->getSize() + PropID->getSize(), PropType->getOffset());

		if (m_doNickname) static_cast<void>(parser->advance(sizeof DWORD)); // reserved

		value = getPVParser(*ulPropTag, m_doNickname, m_doRuleProcessing);
		if (value)
		{
			value->block::parse(parser, false);
		}
	}

	void SPropValueStruct::parseBlocks()
	{
		setText(L"Property[%1!d!]", m_index);
		addChild(
			ulPropTag,
			L"PropTag = 0x%1!08X! (%2!ws!)",
			ulPropTag->getData(),
			proptype::TypeToString(*ulPropTag).c_str());

		const auto propTagNames = proptags::PropTagToPropName(*ulPropTag, false);
		if (!propTagNames.bestGuess.empty())
		{
			ulPropTag->addSubHeader(L"Name: %1!ws!", propTagNames.bestGuess.c_str());
		}

		if (!propTagNames.otherMatches.empty())
		{
			ulPropTag->addSubHeader(L"Other Matches: %1!ws!", propTagNames.otherMatches.c_str());
		}

		addChild(value);
	}
} // namespace smartview