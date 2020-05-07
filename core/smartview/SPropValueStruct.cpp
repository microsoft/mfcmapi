#include <core/stdafx.h>
#include <core/smartview/SPropValueStruct.h>
#include <core/interpret/proptags.h>

namespace smartview
{
	void SPropValueStruct::parse()
	{
		const auto ulCurrOffset = m_Parser->getOffset();

		PropType = blockT<WORD>::parse(m_Parser);
		PropID = blockT<WORD>::parse(m_Parser);

		ulPropTag = blockT<ULONG>::create(
			PROP_TAG(*PropType, *PropID), PropType->getSize() + PropID->getSize(), PropType->getOffset());

		if (m_doNickname) static_cast<void>(m_Parser->advance(sizeof DWORD)); // reserved

		value = getPVParser(*PropType);
		if (value)
		{
			value->parse(m_Parser, *ulPropTag, m_doNickname, m_doRuleProcessing);
		}
	}

	void SPropValueStruct::parseBlocks()
	{
		auto propRoot = std::make_shared<block>();
		addChild(propRoot);
		propRoot->setText(L"Property[%1!d!]\r\n", m_index);
		propRoot->addChild(ulPropTag, L"Property = 0x%1!08X!", ulPropTag->getData());

		const auto propTagNames = proptags::PropTagToPropName(*ulPropTag, false);
		if (!propTagNames.bestGuess.empty())
		{
			// TODO: Add this as a child of ulPropTag
			propRoot->terminateBlock();
			propRoot->addHeader(L"Name: %1!ws!", propTagNames.bestGuess.c_str());
		}

		if (!propTagNames.otherMatches.empty())
		{
			// TODO: Add this as a child of ulPropTag
			propRoot->terminateBlock();
			propRoot->addHeader(L"Other Matches: %1!ws!", propTagNames.otherMatches.c_str());
		}

		propRoot->terminateBlock();
		if (value)
		{
			const auto propString = value->PropBlock();
			if (!propString->empty())
			{
				propRoot->addChild(propString, L"PropString = %1!ws!", propString->c_str());
			}

			const auto alt = value->AltPropBlock();
			if (!alt->empty())
			{
				propRoot->addChild(alt, L" AltPropString = %1!ws!", alt->c_str());
			}

			const auto szSmartView = value->SmartViewBlock();
			if (!szSmartView->empty())
			{
				propRoot->terminateBlock();
				propRoot->addChild(szSmartView, L"Smart View: %1!ws!", szSmartView->c_str());
			}
		}
	}
} // namespace smartview