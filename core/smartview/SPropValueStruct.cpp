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

		switch (*PropType)
		{
		case PT_I2:
			value = std::make_shared<I2BLock>();
			break;
		case PT_LONG:
			value = std::make_shared<LongBLock>();
			break;
		case PT_ERROR:
			value = std::make_shared<ErrorBlock>();
			break;
		case PT_R4:
			value = std::make_shared<R4BLock>();
			break;
		case PT_DOUBLE:
			value = std::make_shared<DoubleBlock>();
			break;
		case PT_BOOLEAN:
			value = std::make_shared<BooleanBlock>();
			break;
		case PT_I8:
			value = std::make_shared<I8Block>();
			break;
		case PT_SYSTIME:
			value = std::make_shared<FILETIMEBLock>();
			break;
		case PT_STRING8:
			value = std::make_shared<CountedStringA>();
			break;
		case PT_BINARY:
			value = std::make_shared<SBinaryBlock>();
			break;
		case PT_UNICODE:
			value = std::make_shared<CountedStringW>();
			break;
		case PT_CLSID:
			value = std::make_shared<CLSIDBlock>();
			break;
		case PT_MV_STRING8:
			value = std::make_shared<StringArrayA>();
			break;
		case PT_MV_UNICODE:
			value = std::make_shared<StringArrayW>();
			break;
		case PT_MV_BINARY:
			value = std::make_shared<SBinaryArrayBlock>();
			break;
		default:
			break;
		}

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