#include <core/stdafx.h>
#include <core/smartview/PropertiesStruct.h>
#include <core/smartview/SmartView.h>
#include <core/interpret/proptags.h>

namespace smartview
{
	SBinaryBlock::SBinaryBlock(std::shared_ptr<binaryParser> parser)
	{
		cb.parse<DWORD>(parser);
		// Note that we're not placing a restriction on how large a multivalued binary property we can parse. May need to revisit this.
		lpb.parse(parser, cb);
	}

	void PropertiesStruct::init(std::shared_ptr<binaryParser> parser, DWORD cValues, bool bRuleCondition)
	{
		SetMaxEntries(cValues);
		if (bRuleCondition) EnableRuleConditionParsing();
		parse(parser, false);
	}

	void PropertiesStruct::Parse()
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
			m_Props.emplace_back(std::make_shared<SPropValueStruct>(m_Parser, m_NickName, m_RuleCondition));
			if (!m_Parser->RemainingBytes()) break;
			dwPropCount++;
		}
	}

	void PropertiesStruct::ParseBlocks()
	{
		auto i = 0;
		for (const auto& prop : m_Props)
		{
			if (i != 0)
			{
				terminateBlock();
			}

			auto propBlock = std::make_shared<block>();
			propBlock->setText(L"Property[%1!d!]\r\n", i++);
			propBlock->addChild(prop->ulPropTag, L"Property = 0x%1!08X!", prop->ulPropTag.getData());

			auto propTagNames = proptags::PropTagToPropName(prop->ulPropTag, false);
			if (!propTagNames.bestGuess.empty())
			{
				propBlock->terminateBlock();
				propBlock->addChild(prop->ulPropTag, L"Name: %1!ws!", propTagNames.bestGuess.c_str());
			}

			if (!propTagNames.otherMatches.empty())
			{
				propBlock->terminateBlock();
				propBlock->addChild(prop->ulPropTag, L"Other Matches: %1!ws!", propTagNames.otherMatches.c_str());
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

			addChild(propBlock);
		}
	}

	SPropValueStruct::SPropValueStruct(std::shared_ptr<binaryParser>& parser, bool doNickname, bool doRuleProcessing)
	{
		const auto ulCurrOffset = parser->GetCurrentOffset();

		PropType.parse<WORD>(parser);
		PropID.parse<WORD>(parser);

		ulPropTag.setData(PROP_TAG(PropType, PropID));
		ulPropTag.setSize(PropType.getSize() + PropID.getSize());
		ulPropTag.setOffset(PropType.getOffset());
		dwAlignPad = 0;

		if (doNickname) (void) parser->advance(sizeof DWORD); // reserved

		switch (PropType)
		{
		case PT_I2:
			// TODO: Insert proper property struct parsing here
			if (doNickname) Value.i.parse<WORD>(parser);
			if (doNickname) parser->advance(sizeof WORD);
			if (doNickname) parser->advance(sizeof DWORD);
			break;
		case PT_LONG:
			Value.l.parse<DWORD>(parser);
			if (doNickname) parser->advance(sizeof DWORD);
			break;
		case PT_ERROR:
			Value.err.parse<DWORD>(parser);
			if (doNickname) parser->advance(sizeof DWORD);
			break;
		case PT_R4:
			Value.flt.parse<float>(parser);
			if (doNickname) parser->advance(sizeof DWORD);
			break;
		case PT_DOUBLE:
			Value.dbl.parse<double>(parser);
			break;
		case PT_BOOLEAN:
			if (doRuleProcessing)
			{
				Value.b.parse<BYTE>(parser);
			}
			else
			{
				Value.b.parse<WORD>(parser);
			}

			if (doNickname) parser->advance(sizeof WORD);
			if (doNickname) parser->advance(sizeof DWORD);
			break;
		case PT_I8:
			Value.li.parse<LARGE_INTEGER>(parser);
			break;
		case PT_SYSTIME:
			Value.ft.dwHighDateTime.parse<DWORD>(parser);
			Value.ft.dwLowDateTime.parse<DWORD>(parser);
			break;
		case PT_STRING8:
			if (doRuleProcessing)
			{
				Value.lpszA.str = blockStringA::parse(parser);
				Value.lpszA.cb.setData(static_cast<DWORD>(Value.lpszA.str->length()));
			}
			else
			{
				if (doNickname)
				{
					(void) parser->advance(sizeof LARGE_INTEGER); // union
					Value.lpszA.cb.parse<DWORD>(parser);
				}
				else
				{
					Value.lpszA.cb.parse<WORD>(parser);
				}

				Value.lpszA.str = blockStringA::parse(parser, Value.lpszA.cb);
			}

			break;
		case PT_BINARY:
			if (doNickname)
			{
				(void) parser->advance(sizeof LARGE_INTEGER); // union
			}

			if (doRuleProcessing || doNickname)
			{
				Value.bin.cb.parse<DWORD>(parser);
			}
			else
			{
				Value.bin.cb.parse<WORD>(parser);
			}

			// Note that we're not placing a restriction on how large a binary property we can parse. May need to revisit this.
			Value.bin.lpb.parse(parser, Value.bin.cb);
			break;
		case PT_UNICODE:
			if (doRuleProcessing)
			{
				Value.lpszW.str = blockStringW::parse(parser);
				Value.lpszW.cb.setData(static_cast<DWORD>(Value.lpszW.str->length()));
			}
			else
			{
				if (doNickname)
				{
					(void) parser->advance(sizeof LARGE_INTEGER); // union
					Value.lpszW.cb.parse<DWORD>(parser);
				}
				else
				{
					Value.lpszW.cb.parse<WORD>(parser);
				}

				Value.lpszW.str = blockStringW::parse(parser, Value.lpszW.cb / sizeof(WCHAR));
			}
			break;
		case PT_CLSID:
			if (doNickname) (void) parser->advance(sizeof LARGE_INTEGER); // union
			Value.lpguid.parse<GUID>(parser);
			break;
		case PT_MV_STRING8:
			if (doNickname)
			{
				(void) parser->advance(sizeof LARGE_INTEGER); // union
				Value.MVszA.cValues.parse<DWORD>(parser);
			}
			else
			{
				Value.MVszA.cValues.parse<WORD>(parser);
			}

			if (Value.MVszA.cValues)
			//if (Value.MVszA.cValues && Value.MVszA.cValues < _MaxEntriesLarge)
			{
				Value.MVszA.lppszA.reserve(Value.MVszA.cValues);
				for (ULONG j = 0; j < Value.MVszA.cValues; j++)
				{
					Value.MVszA.lppszA.emplace_back(std::make_shared<blockStringA>(parser));
				}
			}
			break;
		case PT_MV_UNICODE:
			if (doNickname)
			{
				(void) parser->advance(sizeof LARGE_INTEGER); // union
				Value.MVszW.cValues.parse<DWORD>(parser);
			}
			else
			{
				Value.MVszW.cValues.parse<WORD>(parser);
			}

			if (Value.MVszW.cValues && Value.MVszW.cValues < _MaxEntriesLarge)
			{
				Value.MVszW.lppszW.reserve(Value.MVszW.cValues);
				for (ULONG j = 0; j < Value.MVszW.cValues; j++)
				{
					Value.MVszW.lppszW.emplace_back(std::make_shared<blockStringW>(parser));
				}
			}
			break;
		case PT_MV_BINARY:
			if (doNickname)
			{
				(void) parser->advance(sizeof LARGE_INTEGER); // union
				Value.MVbin.cValues.parse<DWORD>(parser);
			}
			else
			{
				Value.MVbin.cValues.parse<WORD>(parser);
			}

			if (Value.MVbin.cValues && Value.MVbin.cValues < _MaxEntriesLarge)
			{
				for (ULONG j = 0; j < Value.MVbin.cValues; j++)
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