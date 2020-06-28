#include <core/stdafx.h>
#include <core/smartview/TimeZoneDefinition.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	TZRule::TZRule(const std::shared_ptr<binaryParser>& parser)
	{
		bMajorVersion = blockT<BYTE>::parse(parser);
		bMinorVersion = blockT<BYTE>::parse(parser);
		wReserved = blockT<WORD>::parse(parser);
		wTZRuleFlags = blockT<WORD>::parse(parser);
		wYear = blockT<WORD>::parse(parser);
		X = blockBytes::parse(parser, 14);
		lBias = blockT<DWORD>::parse(parser);
		lStandardBias = blockT<DWORD>::parse(parser);
		lDaylightBias = blockT<DWORD>::parse(parser);
		stStandardDate = block::parse<SYSTEMTIMEBlock>(parser, false);
		stDaylightDate = block::parse<SYSTEMTIMEBlock>(parser, false);
	}

	void TimeZoneDefinition::parse()
	{
		m_bMajorVersion = blockT<BYTE>::parse(parser);
		m_bMinorVersion = blockT<BYTE>::parse(parser);
		m_cbHeader = blockT<WORD>::parse(parser);
		m_wReserved = blockT<WORD>::parse(parser);
		m_cchKeyName = blockT<WORD>::parse(parser);
		m_szKeyName = blockStringW::parse(parser, *m_cchKeyName);
		m_cRules = blockT<WORD>::parse(parser);

		if (*m_cRules && *m_cRules < _MaxEntriesSmall)
		{
			m_lpTZRule.reserve(*m_cRules);
			for (WORD i = 0; i < *m_cRules; i++)
			{
				m_lpTZRule.emplace_back(std::make_shared<TZRule>(parser));
			}
		}
	}

	void TimeZoneDefinition::parseBlocks()
	{
		setText(L"Time Zone Definition");
		addChild(m_bMajorVersion, L"bMajorVersion = 0x%1!02X! (%1!d!)", m_bMajorVersion->getData());
		addChild(m_bMinorVersion, L"bMinorVersion = 0x%1!02X! (%1!d!)", m_bMinorVersion->getData());
		addChild(m_cbHeader, L"cbHeader = 0x%1!04X! (%1!d!)", m_cbHeader->getData());
		addChild(m_wReserved, L"wReserved = 0x%1!04X! (%1!d!)", m_wReserved->getData());
		addChild(m_cchKeyName, L"cchKeyName = 0x%1!04X! (%1!d!)", m_cchKeyName->getData());
		addChild(m_szKeyName, L"szKeyName = %1!ws!", m_szKeyName->c_str());
		addChild(m_cRules, L"cRules = 0x%1!04X! (%1!d!)", m_cRules->getData());

		auto i = 0;
		for (const auto& rule : m_lpTZRule)
		{
			auto ruleBlock = create(L"TZRule[0x%1!X!]", i);
			addChild(ruleBlock);
			ruleBlock->addChild(
				rule->bMajorVersion, L"bMajorVersion = 0x%1!02X! (%1!d!)", rule->bMajorVersion->getData());
			ruleBlock->addChild(
				rule->bMinorVersion, L"bMinorVersion = 0x%1!02X! (%1!d!)", rule->bMinorVersion->getData());
			ruleBlock->addChild(rule->wReserved, L"wReserved = 0x%1!04X! (%1!d!)", rule->wReserved->getData());
			ruleBlock->addChild(
				rule->wTZRuleFlags,
				L"wTZRuleFlags = 0x%1!04X! = %2!ws!",
				rule->wTZRuleFlags->getData(),
				flags::InterpretFlags(flagTZRule, *rule->wTZRuleFlags).c_str());
			ruleBlock->addChild(rule->wYear, L"wYear = 0x%1!04X! (%1!d!)", rule->wYear->getData());
			ruleBlock->addLabeledChild(L"X =", rule->X);

			ruleBlock->addChild(rule->lBias, L"lBias = 0x%1!08X! (%1!d!)", rule->lBias->getData());
			ruleBlock->addChild(
				rule->lStandardBias, L"lStandardBias = 0x%1!08X! (%1!d!)", rule->lStandardBias->getData());
			ruleBlock->addChild(
				rule->lDaylightBias, L"lDaylightBias = 0x%1!08X! (%1!d!)", rule->lDaylightBias->getData());
			ruleBlock->addChild(rule->stStandardDate, L"stStandardDate");
			ruleBlock->addChild(rule->stDaylightDate, L"stDaylightDate");

			i++;
		}
	}
} // namespace smartview