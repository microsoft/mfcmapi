#include <core/stdafx.h>
#include <core/smartview/TimeZoneDefinition.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	TZRule::TZRule(std::shared_ptr<binaryParser>& parser)
	{
		bMajorVersion.parse<BYTE>(parser);
		bMinorVersion.parse<BYTE>(parser);
		wReserved.parse<WORD>(parser);
		wTZRuleFlags.parse<WORD>(parser);
		wYear.parse<WORD>(parser);
		X.parse(parser, 14);
		lBias.parse<DWORD>(parser);
		lStandardBias.parse<DWORD>(parser);
		lDaylightBias.parse<DWORD>(parser);
		stStandardDate.wYear.parse<WORD>(parser);
		stStandardDate.wMonth.parse<WORD>(parser);
		stStandardDate.wDayOfWeek.parse<WORD>(parser);
		stStandardDate.wDay.parse<WORD>(parser);
		stStandardDate.wHour.parse<WORD>(parser);
		stStandardDate.wMinute.parse<WORD>(parser);
		stStandardDate.wSecond.parse<WORD>(parser);
		stStandardDate.wMilliseconds.parse<WORD>(parser);
		stDaylightDate.wYear.parse<WORD>(parser);
		stDaylightDate.wMonth.parse<WORD>(parser);
		stDaylightDate.wDayOfWeek.parse<WORD>(parser);
		stDaylightDate.wDay.parse<WORD>(parser);
		stDaylightDate.wHour.parse<WORD>(parser);
		stDaylightDate.wMinute.parse<WORD>(parser);
		stDaylightDate.wSecond.parse<WORD>(parser);
		stDaylightDate.wMilliseconds.parse<WORD>(parser);
	}

	void TimeZoneDefinition::Parse()
	{
		m_bMajorVersion.parse<BYTE>(m_Parser);
		m_bMinorVersion.parse<BYTE>(m_Parser);
		m_cbHeader.parse<WORD>(m_Parser);
		m_wReserved.parse<WORD>(m_Parser);
		m_cchKeyName.parse<WORD>(m_Parser);
		m_szKeyName.parse(m_Parser, m_cchKeyName);
		m_cRules.parse<WORD>(m_Parser);

		if (m_cRules && m_cRules < _MaxEntriesSmall)
		{
			m_lpTZRule.reserve(m_cRules);
			for (ULONG i = 0; i < m_cRules; i++)
			{
				m_lpTZRule.emplace_back(std::make_shared<TZRule>(m_Parser));
			}
		}
	}

	void TimeZoneDefinition::ParseBlocks()
	{
		setRoot(L"Time Zone Definition: \r\n");
		addBlock(m_bMajorVersion, L"bMajorVersion = 0x%1!02X! (%1!d!)\r\n", m_bMajorVersion.getData());
		addBlock(m_bMinorVersion, L"bMinorVersion = 0x%1!02X! (%1!d!)\r\n", m_bMinorVersion.getData());
		addBlock(m_cbHeader, L"cbHeader = 0x%1!04X! (%1!d!)\r\n", m_cbHeader.getData());
		addBlock(m_wReserved, L"wReserved = 0x%1!04X! (%1!d!)\r\n", m_wReserved.getData());
		addBlock(m_cchKeyName, L"cchKeyName = 0x%1!04X! (%1!d!)\r\n", m_cchKeyName.getData());
		addBlock(m_szKeyName, L"szKeyName = %1!ws!\r\n", m_szKeyName.c_str());
		addBlock(m_cRules, L"cRules = 0x%1!04X! (%1!d!)", m_cRules.getData());

		auto i = 0;
		for (const auto& rule : m_lpTZRule)
		{
			terminateBlock();
			addBlankLine();
			addBlock(
				rule->bMajorVersion,
				L"TZRule[0x%1!X!].bMajorVersion = 0x%2!02X! (%2!d!)\r\n",
				i,
				rule->bMajorVersion.getData());
			addBlock(
				rule->bMinorVersion,
				L"TZRule[0x%1!X!].bMinorVersion = 0x%2!02X! (%2!d!)\r\n",
				i,
				rule->bMinorVersion.getData());
			addBlock(
				rule->wReserved, L"TZRule[0x%1!X!].wReserved = 0x%2!04X! (%2!d!)\r\n", i, rule->wReserved.getData());
			addBlock(
				rule->wTZRuleFlags,
				L"TZRule[0x%1!X!].wTZRuleFlags = 0x%2!04X! = %3!ws!\r\n",
				i,
				rule->wTZRuleFlags.getData(),
				flags::InterpretFlags(flagTZRule, rule->wTZRuleFlags).c_str());
			addBlock(rule->wYear, L"TZRule[0x%1!X!].wYear = 0x%2!04X! (%2!d!)\r\n", i, rule->wYear.getData());
			addHeader(L"TZRule[0x%1!X!].X = ", i);
			addBlock(rule->X);

			terminateBlock();
			addBlock(rule->lBias, L"TZRule[0x%1!X!].lBias = 0x%2!08X! (%2!d!)\r\n", i, rule->lBias.getData());
			addBlock(
				rule->lStandardBias,
				L"TZRule[0x%1!X!].lStandardBias = 0x%2!08X! (%2!d!)\r\n",
				i,
				rule->lStandardBias.getData());
			addBlock(
				rule->lDaylightBias,
				L"TZRule[0x%1!X!].lDaylightBias = 0x%2!08X! (%2!d!)\r\n",
				i,
				rule->lDaylightBias.getData());
			addBlankLine();
			addBlock(
				rule->stStandardDate.wYear,
				L"TZRule[0x%1!X!].stStandardDate.wYear = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stStandardDate.wYear.getData());
			addBlock(
				rule->stStandardDate.wMonth,
				L"TZRule[0x%1!X!].stStandardDate.wMonth = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stStandardDate.wMonth.getData());
			addBlock(
				rule->stStandardDate.wDayOfWeek,
				L"TZRule[0x%1!X!].stStandardDate.wDayOfWeek = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stStandardDate.wDayOfWeek.getData());
			addBlock(
				rule->stStandardDate.wDay,
				L"TZRule[0x%1!X!].stStandardDate.wDay = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stStandardDate.wDay.getData());
			addBlock(
				rule->stStandardDate.wHour,
				L"TZRule[0x%1!X!].stStandardDate.wHour = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stStandardDate.wHour.getData());
			addBlock(
				rule->stStandardDate.wMinute,
				L"TZRule[0x%1!X!].stStandardDate.wMinute = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stStandardDate.wMinute.getData());
			addBlock(
				rule->stStandardDate.wSecond,
				L"TZRule[0x%1!X!].stStandardDate.wSecond = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stStandardDate.wSecond.getData());
			addBlock(
				rule->stStandardDate.wMilliseconds,
				L"TZRule[0x%1!X!].stStandardDate.wMilliseconds = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stStandardDate.wMilliseconds.getData());
			addBlankLine();
			addBlock(
				rule->stDaylightDate.wYear,
				L"TZRule[0x%1!X!].stDaylightDate.wYear = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stDaylightDate.wYear.getData());
			addBlock(
				rule->stDaylightDate.wMonth,
				L"TZRule[0x%1!X!].stDaylightDate.wMonth = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stDaylightDate.wMonth.getData());
			addBlock(
				rule->stDaylightDate.wDayOfWeek,
				L"TZRule[0x%1!X!].stDaylightDate.wDayOfWeek = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stDaylightDate.wDayOfWeek.getData());
			addBlock(
				rule->stDaylightDate.wDay,
				L"TZRule[0x%1!X!].stDaylightDate.wDay = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stDaylightDate.wDay.getData());
			addBlock(
				rule->stDaylightDate.wHour,
				L"TZRule[0x%1!X!].stDaylightDate.wHour = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stDaylightDate.wHour.getData());
			addBlock(
				rule->stDaylightDate.wMinute,
				L"TZRule[0x%1!X!].stDaylightDate.wMinute = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stDaylightDate.wMinute.getData());
			addBlock(
				rule->stDaylightDate.wSecond,
				L"TZRule[0x%1!X!].stDaylightDate.wSecond = 0x%2!X! (%2!d!)\r\n",
				i,
				rule->stDaylightDate.wSecond.getData());
			addBlock(
				rule->stDaylightDate.wMilliseconds,
				L"TZRule[0x%1!X!].stDaylightDate.wMilliseconds = 0x%2!X! (%2!d!)",
				i,
				rule->stDaylightDate.wMilliseconds.getData());

			i++;
		}
	}
} // namespace smartview