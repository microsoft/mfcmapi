#include <core/stdafx.h>
#include <core/smartview/SPropValueStruct.h>
#include <core/interpret/proptags.h>

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

		value = getPVParser(*PropType);
		if (value)
		{
			value->parse(parser, *ulPropTag, m_doNickname, m_doRuleProcessing);
		}
	}

	void SPropValueStruct::parseBlocks()
	{
		setText(L"Property[%1!d!]\r\n", m_index);
		addChild(ulPropTag, L"Property = 0x%1!08X!", ulPropTag->getData());

		const auto propTagNames = proptags::PropTagToPropName(*ulPropTag, false);
		if (!propTagNames.bestGuess.empty())
		{
			// TODO: Add this as a child of ulPropTag
			terminateBlock();
			addHeader(L"Name: %1!ws!", propTagNames.bestGuess.c_str());
		}

		if (!propTagNames.otherMatches.empty())
		{
			// TODO: Add this as a child of ulPropTag
			terminateBlock();
			addHeader(L"Other Matches: %1!ws!", propTagNames.otherMatches.c_str());
		}

		terminateBlock();
		if (value)
		{
			const auto propString = value->PropBlock();
			if (!propString->empty())
			{
				addChild(propString, L"PropString = %1!ws!", propString->c_str());
			}

			const auto alt = value->AltPropBlock();
			if (!alt->empty())
			{
				addChild(alt, L" AltPropString = %1!ws!", alt->c_str());
			}

			const auto szSmartView = value->SmartViewBlock();
			if (!szSmartView->empty())
			{
				terminateBlock();
				addChild(szSmartView, L"Smart View: %1!ws!", szSmartView->c_str());
			}
		}
	}
} // namespace smartview