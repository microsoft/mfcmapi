#include <core/stdafx.h>
#include <core/smartview/PropertiesStruct.h>
#include <core/smartview/SmartView.h>
#include <core/interpret/proptags.h>

namespace smartview
{
	SBinaryBlock::SBinaryBlock(const std::shared_ptr<binaryParser>& parser)
	{
		cb = blockT<DWORD>::parse(parser);
		// Note that we're not placing a restriction on how large a multivalued binary property we can parse. May need to revisit this.
		lpb = blockBytes::parse(parser, *cb);
	}

	void PropertiesStruct::parse(const std::shared_ptr<binaryParser>& parser, DWORD cValues, bool bRuleCondition)
	{
		SetMaxEntries(cValues);
		if (bRuleCondition) EnableRuleConditionParsing();
		smartViewParser::parse(parser, false);
	}

	void PropertiesStruct::parse()
	{
		// For consistancy with previous parsings, we'll refuse to parse if asked to parse more than _MaxEntriesSmall
		// However, we may want to reconsider this choice.
		if (m_MaxEntries > _MaxEntriesSmall) return;

		DWORD dwPropCount = 0;

		// If we have a non-default max, it was computed elsewhere and we do expect to have that many entries. So we can reserve.
		if (m_MaxEntries != _MaxEntriesSmall)
		{
			m_Props.reserve(m_MaxEntries);
		}

		for (;;)
		{
			if (dwPropCount >= m_MaxEntries) break;
			auto oldSize = m_Parser->getSize();
			m_Props.emplace_back(std::make_shared<SPropValueStruct>(m_Parser, m_NickName, m_RuleCondition));
			auto newSize = m_Parser->getSize();
			if (newSize == 0 || newSize == oldSize) break;
			dwPropCount++;
		}
	}

	void PropertiesStruct::parseBlocks()
	{
		auto i = 0;
		for (const auto& prop : m_Props)
		{
			if (i != 0)
			{
				terminateBlock();
			}

			auto propBlock = std::make_shared<block>();
			addChild(propBlock);
			propBlock->setText(L"Property[%1!d!]\r\n", i++);
			propBlock->addChild(prop->ulPropTag, L"Property = 0x%1!08X!", prop->ulPropTag->getData());

			auto propTagNames = proptags::PropTagToPropName(*prop->ulPropTag, false);
			if (!propTagNames.bestGuess.empty())
			{
				// TODO: Add this as a child of ulPropTag
				propBlock->terminateBlock();
				propBlock->addHeader(L"Name: %1!ws!", propTagNames.bestGuess.c_str());
			}

			if (!propTagNames.otherMatches.empty())
			{
				// TODO: Add this as a child of ulPropTag
				propBlock->terminateBlock();
				propBlock->addHeader(L"Other Matches: %1!ws!", propTagNames.otherMatches.c_str());
			}

			propBlock->terminateBlock();
			propBlock->addChild(prop->PropBlock(), L"PropString = %1!ws! ", prop->PropBlock()->c_str());
			propBlock->addChild(prop->AltPropBlock(), L"AltPropString = %1!ws!", prop->AltPropBlock()->c_str());

			auto szSmartView = prop->SmartViewBlock();
			if (!szSmartView->empty())
			{
				propBlock->terminateBlock();
				propBlock->addChild(szSmartView, L"Smart View: %1!ws!", szSmartView->c_str());
			}
		}
	}

	SPropValueStruct::SPropValueStruct(
		const std::shared_ptr<binaryParser>& parser,
		bool doNickname,
		bool doRuleProcessing)
	{
		const auto ulCurrOffset = parser->getOffset();

		PropType = blockT<WORD>::parse(parser);
		PropID = blockT<WORD>::parse(parser);

		ulPropTag->setData(PROP_TAG(*PropType, *PropID));
		ulPropTag->setSize(PropType->getSize() + PropID->getSize());
		ulPropTag->setOffset(PropType->getOffset());
		dwAlignPad = 0;

		if (doNickname) (void) parser->advance(sizeof DWORD); // reserved

		switch (*PropType)
		{
		case PT_I2:
			// TODO: Insert proper property struct parsing here
			if (doNickname) Value.i = blockT<WORD>::parse(parser);
			if (doNickname) parser->advance(sizeof WORD);
			if (doNickname) parser->advance(sizeof DWORD);
			break;
		case PT_LONG:
			Value.l = blockT<LONG, DWORD>::parse(parser);
			if (doNickname) parser->advance(sizeof DWORD);
			break;
		case PT_ERROR:
			Value.err = blockT<SCODE, DWORD>::parse(parser);
			if (doNickname) parser->advance(sizeof DWORD);
			break;
		case PT_R4:
			Value.flt = blockT<float>::parse(parser);
			if (doNickname) parser->advance(sizeof DWORD);
			break;
		case PT_DOUBLE:
			Value.dbl = blockT<double>::parse(parser);
			break;
		case PT_BOOLEAN:
			if (doRuleProcessing)
			{
				Value.b = blockT<WORD, BYTE>::parse(parser);
			}
			else
			{
				Value.b = blockT<WORD>::parse(parser);
			}

			if (doNickname) parser->advance(sizeof WORD);
			if (doNickname) parser->advance(sizeof DWORD);
			break;
		case PT_I8:
			Value.li = blockT<LARGE_INTEGER>::parse(parser);
			break;
		case PT_SYSTIME:
			Value.ft.dwHighDateTime = blockT<DWORD>::parse(parser);
			Value.ft.dwLowDateTime = blockT<DWORD>::parse(parser);
			break;
		case PT_STRING8:
			if (doRuleProcessing)
			{
				Value.lpszA.str = blockStringA::parse(parser);
				Value.lpszA.cb->setData(static_cast<DWORD>(Value.lpszA.str->length()));
			}
			else
			{
				if (doNickname)
				{
					(void) parser->advance(sizeof LARGE_INTEGER); // union
					Value.lpszA.cb = blockT<DWORD>::parse(parser);
				}
				else
				{
					Value.lpszA.cb = blockT<DWORD, WORD>::parse(parser);
				}

				Value.lpszA.str = blockStringA::parse(parser, *Value.lpszA.cb);
			}

			break;
		case PT_BINARY:
			if (doNickname)
			{
				(void) parser->advance(sizeof LARGE_INTEGER); // union
			}

			if (doRuleProcessing || doNickname)
			{
				Value.bin.cb = blockT<DWORD>::parse(parser);
			}
			else
			{
				Value.bin.cb = blockT<DWORD, WORD>::parse(parser);
			}

			// Note that we're not placing a restriction on how large a binary property we can parse. May need to revisit this.
			Value.bin.lpb = blockBytes::parse(parser, *Value.bin.cb);
			break;
		case PT_UNICODE:
			if (doRuleProcessing)
			{
				Value.lpszW.str = blockStringW::parse(parser);
				Value.lpszW.cb->setData(static_cast<DWORD>(Value.lpszW.str->length()));
			}
			else
			{
				if (doNickname)
				{
					(void) parser->advance(sizeof LARGE_INTEGER); // union
					Value.lpszW.cb = blockT<DWORD>::parse(parser);
				}
				else
				{
					Value.lpszW.cb = blockT<DWORD, WORD>::parse(parser);
				}

				Value.lpszW.str = blockStringW::parse(parser, *Value.lpszW.cb / sizeof(WCHAR));
			}
			break;
		case PT_CLSID:
			if (doNickname) (void) parser->advance(sizeof LARGE_INTEGER); // union
			Value.lpguid = blockT<GUID>::parse(parser);
			break;
		case PT_MV_STRING8:
			if (doNickname)
			{
				(void) parser->advance(sizeof LARGE_INTEGER); // union
				Value.MVszA.cValues = blockT<DWORD>::parse(parser);
			}
			else
			{
				Value.MVszA.cValues = blockT<DWORD, WORD>::parse(parser);
			}

			if (Value.MVszA.cValues)
			//if (Value.MVszA.cValues && Value.MVszA.cValues < _MaxEntriesLarge)
			{
				Value.MVszA.lppszA.reserve(*Value.MVszA.cValues);
				for (ULONG j = 0; j < *Value.MVszA.cValues; j++)
				{
					Value.MVszA.lppszA.emplace_back(std::make_shared<blockStringA>(parser));
				}
			}
			break;
		case PT_MV_UNICODE:
			if (doNickname)
			{
				(void) parser->advance(sizeof LARGE_INTEGER); // union
				Value.MVszW.cValues = blockT<DWORD>::parse(parser);
			}
			else
			{
				Value.MVszW.cValues = blockT<DWORD, WORD>::parse(parser);
			}

			if (Value.MVszW.cValues && *Value.MVszW.cValues < _MaxEntriesLarge)
			{
				Value.MVszW.lppszW.reserve(*Value.MVszW.cValues);
				for (ULONG j = 0; j < *Value.MVszW.cValues; j++)
				{
					Value.MVszW.lppszW.emplace_back(std::make_shared<blockStringW>(parser));
				}
			}
			break;
		case PT_MV_BINARY:
			if (doNickname)
			{
				(void) parser->advance(sizeof LARGE_INTEGER); // union
				Value.MVbin.cValues = blockT<DWORD>::parse(parser);
			}
			else
			{
				Value.MVbin.cValues = blockT<DWORD, WORD>::parse(parser);
			}

			if (Value.MVbin.cValues && *Value.MVbin.cValues < _MaxEntriesLarge)
			{
				for (ULONG j = 0; j < *Value.MVbin.cValues; j++)
				{
					Value.MVbin.lpbin.emplace_back(std::make_shared<SBinaryBlock>(parser));
				}
			}
			break;
		default:
			break;
		}
	}
} // namespace smartview