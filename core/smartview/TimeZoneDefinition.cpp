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
		stStandardDate.wYear = blockT<WORD>::parse(parser);
		stStandardDate.wMonth = blockT<WORD>::parse(parser);
		stStandardDate.wDayOfWeek = blockT<WORD>::parse(parser);
		stStandardDate.wDay = blockT<WORD>::parse(parser);
		stStandardDate.wHour = blockT<WORD>::parse(parser);
		stStandardDate.wMinute = blockT<WORD>::parse(parser);
		stStandardDate.wSecond = blockT<WORD>::parse(parser);
		stStandardDate.wMilliseconds = blockT<WORD>::parse(parser);
		stDaylightDate.wYear = blockT<WORD>::parse(parser);
		stDaylightDate.wMonth = blockT<WORD>::parse(parser);
		stDaylightDate.wDayOfWeek = blockT<WORD>::parse(parser);
		stDaylightDate.wDay = blockT<WORD>::parse(parser);
		stDaylightDate.wHour = blockT<WORD>::parse(parser);
		stDaylightDate.wMinute = blockT<WORD>::parse(parser);
		stDaylightDate.wSecond = blockT<WORD>::parse(parser);
		stDaylightDate.wMilliseconds = blockT<WORD>::parse(parser);
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
		setText(L"Time Zone Definition: \r\n");
		addChild(m_bMajorVersion, L"bMajorVersion = 0x%1!02X! (%1!d!)\r\n", m_bMajorVersion->getData());
		addChild(m_bMinorVersion, L"bMinorVersion = 0x%1!02X! (%1!d!)\r\n", m_bMinorVersion->getData());
		addChild(m_cbHeader, L"cbHeader = 0x%1!04X! (%1!d!)\r\n", m_cbHeader->getData());
		addChild(m_wReserved, L"wReserved = 0x%1!04X! (%1!d!)\r\n", m_wReserved->getData());
		addChild(m_cchKeyName, L"cchKeyName = 0x%1!04X! (%1!d!)\r\n", m_cchKeyName->getData());
		addChild(m_szKeyName, L"szKeyName = %1!ws!\r\n", m_szKeyName->c_str());
		addChild(m_cRules, L"cRules = 0x%1!04X! (%1!d!)", m_cRules->getData());

		auto i = 0;
		for (const auto& rule : m_lpTZRule)
		{
			terminateBlock();
			addBlankLine();
			addChild(
				rule->bMajorVersion,
				L"TZRule[0x%1!X!].bMajorVersion = 0x%2!02X! (%2!d!)\r\n",
				i,
				rule->bMajorVersion->getData());
			addChild(
				rule->bMinorVersion,
				L"TZRule[0x%1!X!].bMinorVersion = 0x%2!02X! (%2!d!)\r\n",
				i,
				rule->bMinorVersion->getData());
			addChild(
				rule->wReserved, L"TZRule[0x%1!X!].wReserved = 0x%2!04X! (%2!d!)\r\n", i, rule->wReserved->getData());
			addChild(
				rule->wTZRuleFlags,
				L"TZRule[0x%1!X!].wTZRuleFlags = 0x%2!04X! = %3!ws!\r\n",
				i,
				rule->wTZRuleFlags->getData(),
				flags::InterpretFlags(flagTZRule, *rule->wTZRuleFlags).c_str());
			addChild(rule->wYear, L"TZRule[0x%1!X!].wYear = 0x%2!04X! (%2!d!)\r\n", i, rule->wYear->getData());
			addHeader(L"TZRule[0x%1!X!].X = ", i);
			addChild(rule->X);

			terminateBlock();
			addChild(rule->lBias, L"TZRule[0x%1!X!].lBias = 0x%2!08X! (%2!d!)\r\n", i, rule->lBias->getData());
			addChild(
				rule->lStandardBias,
				L"TZRule[0x%1!X!].lStandardBias = 0x%2!08X! (%2!d!)\r\n",
				i,
				rule->lStandardBias->getData());
			addChild(
				rule->lDaylightBias,
				L"TZRule[0x%1!X!].lDaylightBias = 0x%2!08X! (%2!d!)\r\n",
				i,
				rule->lDaylightBias->getData());
			addBlankLine();
			addChild(
				rule->stStandardDate.wYear,
				L"TZRule[0x%1!X!].stStandardDate.wYear = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stStandardDate.wYear->getData());
			addChild(
				rule->stStandardDate.wMonth,
				L"TZRule[0x%1!X!].stStandardDate.wMonth = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stStandardDate.wMonth->getData());
			addChild(
				rule->stStandardDate.wDayOfWeek,
				L"TZRule[0x%1!X!].stStandardDate.wDayOfWeek = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stStandardDate.wDayOfWeek->getData());
			addChild(
				rule->stStandardDate.wDay,
				L"TZRule[0x%1!X!].stStandardDate.wDay = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stStandardDate.wDay->getData());
			addChild(
				rule->stStandardDate.wHour,
				L"TZRule[0x%1!X!].stStandardDate.wHour = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stStandardDate.wHour->getData());
			addChild(
				rule->stStandardDate.wMinute,
				L"TZRule[0x%1!X!].stStandardDate.wMinute = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stStandardDate.wMinute->getData());
			addChild(
				rule->stStandardDate.wSecond,
				L"TZRule[0x%1!X!].stStandardDate.wSecond = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stStandardDate.wSecond->getData());
			addChild(
				rule->stStandardDate.wMilliseconds,
				L"TZRule[0x%1!X!].stStandardDate.wMilliseconds = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stStandardDate.wMilliseconds->getData());
			addBlankLine();
			addChild(
				rule->stDaylightDate.wYear,
				L"TZRule[0x%1!X!].stDaylightDate.wYear = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stDaylightDate.wYear->getData());
			addChild(
				rule->stDaylightDate.wMonth,
				L"TZRule[0x%1!X!].stDaylightDate.wMonth = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stDaylightDate.wMonth->getData());
			addChild(
				rule->stDaylightDate.wDayOfWeek,
				L"TZRule[0x%1!X!].stDaylightDate.wDayOfWeek = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stDaylightDate.wDayOfWeek->getData());
			addChild(
				rule->stDaylightDate.wDay,
				L"TZRule[0x%1!X!].stDaylightDate.wDay = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stDaylightDate.wDay->getData());
			addChild(
				rule->stDaylightDate.wHour,
				L"TZRule[0x%1!X!].stDaylightDate.wHour = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stDaylightDate.wHour->getData());
			addChild(
				rule->stDaylightDate.wMinute,
				L"TZRule[0x%1!X!].stDaylightDate.wMinute = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stDaylightDate.wMinute->getData());
			addChild(
				rule->stDaylightDate.wSecond,
				L"TZRule[0x%1!X!].stDaylightDate.wSecond = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stDaylightDate.wSecond->getData());
			addChild(
				rule->stDaylightDate.wMilliseconds,
				L"TZRule[0x%1!X!].stDaylightDate.wMilliseconds = 0x%2!X! (%2!d!)",
				i,
				rule->stDaylightDate.wMilliseconds->getData());

			i++;
		}
	}
} // namespace smartview