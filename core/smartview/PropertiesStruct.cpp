#include <core/stdafx.h>
#include <core/smartview/PropertiesStruct.h>
#include <core/smartview/SmartView.h>
#include <core/interpret/proptags.h>

namespace smartview
{
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
		for (auto& prop : m_Props)
		{
			if (i != 0)
			{
				terminateBlock();
			}

			auto propBlock = block();
			propBlock.setText(L"Property[%1!d!]\r\n", i++);
			propBlock.addBlock(prop->ulPropTag, L"Property = 0x%1!08X!", prop->ulPropTag.getData());

			auto propTagNames = proptags::PropTagToPropName(prop->ulPropTag, false);
			if (!propTagNames.bestGuess.empty())
			{
				propBlock.terminateBlock();
				propBlock.addBlock(prop->ulPropTag, L"Name: %1!ws!", propTagNames.bestGuess.c_str());
			}

			if (!propTagNames.otherMatches.empty())
			{
				propBlock.terminateBlock();
				propBlock.addBlock(prop->ulPropTag, L"Other Matches: %1!ws!", propTagNames.otherMatches.c_str());
			}

			propBlock.terminateBlock();
			propBlock.addBlock(prop->PropBlock(), L"PropString = %1!ws! ", prop->PropBlock().c_str());
			propBlock.addBlock(prop->AltPropBlock(), L"AltPropString = %1!ws!", prop->AltPropBlock().c_str());

			auto& szSmartView = prop->SmartViewBlock();
			if (!szSmartView.empty())
			{
				propBlock.terminateBlock();
				propBlock.addBlock(szSmartView, L"Smart View: %1!ws!", szSmartView.c_str());
			}

			addBlock(propBlock);
		}
	}

	SPropValueStruct::SPropValueStruct(std::shared_ptr<binaryParser>& parser, bool doNickname, bool doRuleProcessing)
	{
		const auto ulCurrOffset = parser->GetCurrentOffset();

		PropType = parser->Get<WORD>();
		PropID = parser->Get<WORD>();

		ulPropTag.setData(PROP_TAG(PropType, PropID));
		ulPropTag.setSize(PropType.getSize() + PropID.getSize());
		ulPropTag.setOffset(PropType.getOffset());
		dwAlignPad = 0;

		if (doNickname) (void) parser->Get<DWORD>(); // reserved

		switch (PropType)
		{
		case PT_I2:
			// TODO: Insert proper property struct parsing here
			if (doNickname) Value.i = parser->Get<WORD>();
			if (doNickname) parser->Get<WORD>();
			if (doNickname) parser->Get<DWORD>();
			break;
		case PT_LONG:
			Value.l = parser->Get<DWORD>();
			if (doNickname) parser->Get<DWORD>();
			break;
		case PT_ERROR:
			Value.err = parser->Get<DWORD>();
			if (doNickname) parser->Get<DWORD>();
			break;
		case PT_R4:
			Value.flt = parser->Get<float>();
			if (doNickname) parser->Get<DWORD>();
			break;
		case PT_DOUBLE:
			Value.dbl = parser->Get<double>();
			break;
		case PT_BOOLEAN:
			if (doRuleProcessing)
			{
				Value.b = parser->Get<BYTE>();
			}
			else
			{
				Value.b = parser->Get<WORD>();
			}

			if (doNickname) parser->Get<WORD>();
			if (doNickname) parser->Get<DWORD>();
			break;
		case PT_I8:
			Value.li = parser->Get<LARGE_INTEGER>();
			break;
		case PT_SYSTIME:
			Value.ft.dwHighDateTime = parser->Get<DWORD>();
			Value.ft.dwLowDateTime = parser->Get<DWORD>();
			break;
		case PT_STRING8:
			if (doRuleProcessing)
			{
				Value.lpszA.str.parse(parser);
				Value.lpszA.cb.setData(static_cast<DWORD>(Value.lpszA.str.length()));
			}
			else
			{
				if (doNickname)
				{
					(void) parser->Get<LARGE_INTEGER>(); // union
					Value.lpszA.cb = parser->Get<DWORD>();
				}
				else
				{
					Value.lpszA.cb = parser->Get<WORD>();
				}

				Value.lpszA.str.parse(parser, Value.lpszA.cb);
			}

			break;
		case PT_BINARY:
			if (doNickname)
			{
				(void) parser->Get<LARGE_INTEGER>(); // union
			}

			if (doRuleProcessing || doNickname)
			{
				Value.bin.cb = parser->Get<DWORD>();
			}
			else
			{
				Value.bin.cb = parser->Get<WORD>();
			}

			// Note that we're not placing a restriction on how large a binary property we can parse. May need to revisit this.
			Value.bin.lpb = parser->GetBYTES(Value.bin.cb);
			break;
		case PT_UNICODE:
			if (doRuleProcessing)
			{
				Value.lpszW.str.parse(parser);
				Value.lpszW.cb.setData(static_cast<DWORD>(Value.lpszW.str.length()));
			}
			else
			{
				if (doNickname)
				{
					(void) parser->Get<LARGE_INTEGER>(); // union
					Value.lpszW.cb = parser->Get<DWORD>();
				}
				else
				{
					Value.lpszW.cb = parser->Get<WORD>();
				}

				Value.lpszW.str.parse(parser, Value.lpszW.cb / sizeof(WCHAR));
			}
			break;
		case PT_CLSID:
			if (doNickname) (void) parser->Get<LARGE_INTEGER>(); // union
			Value.lpguid = parser->Get<GUID>();
			break;
		case PT_MV_STRING8:
			if (doNickname)
			{
				(void) parser->Get<LARGE_INTEGER>(); // union
				Value.MVszA.cValues = parser->Get<DWORD>();
			}
			else
			{
				Value.MVszA.cValues = parser->Get<WORD>();
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
				(void) parser->Get<LARGE_INTEGER>(); // union
				Value.MVszW.cValues = parser->Get<DWORD>();
			}
			else
			{
				Value.MVszW.cValues = parser->Get<WORD>();
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
				(void) parser->Get<LARGE_INTEGER>(); // union
				Value.MVbin.cValues = parser->Get<DWORD>();
			}
			else
			{
				Value.MVbin.cValues = parser->Get<WORD>();
			}

			if (Value.MVbin.cValues && Value.MVbin.cValues < _MaxEntriesLarge)
			{
				for (ULONG j = 0; j < Value.MVbin.cValues; j++)
				{
					auto bin = SBinaryBlock{};
					bin.cb = parser->Get<DWORD>();
					// Note that we're not placing a restriction on how large a multivalued binary property we can parse. May need to revisit this.
					bin.lpb = parser->GetBYTES(bin.cb);
					Value.MVbin.lpbin.push_back(bin);
				}
			}
			break;
		default:
			break;
		}
	}
} // namespace smartview