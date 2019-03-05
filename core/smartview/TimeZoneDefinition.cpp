#include <core/stdafx.h>
#include <core/smartview/TimeZoneDefinition.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	void TimeZoneDefinition::Parse()
	{
		m_bMajorVersion = m_Parser->Get<BYTE>();
		m_bMinorVersion = m_Parser->Get<BYTE>();
		m_cbHeader = m_Parser->Get<WORD>();
		m_wReserved = m_Parser->Get<WORD>();
		m_cchKeyName = m_Parser->Get<WORD>();
		m_szKeyName = m_Parser->GetStringW(m_cchKeyName);
		m_cRules = m_Parser->Get<WORD>();

		if (m_cRules && m_cRules < _MaxEntriesSmall)
		{
			m_lpTZRule.reserve(m_cRules);
			for (ULONG i = 0; i < m_cRules; i++)
			{
				TZRule tzRule;
				tzRule.bMajorVersion = m_Parser->Get<BYTE>();
				tzRule.bMinorVersion = m_Parser->Get<BYTE>();
				tzRule.wReserved = m_Parser->Get<WORD>();
				tzRule.wTZRuleFlags = m_Parser->Get<WORD>();
				tzRule.wYear = m_Parser->Get<WORD>();
				tzRule.X = m_Parser->GetBYTES(14);
				tzRule.lBias = m_Parser->Get<DWORD>();
				tzRule.lStandardBias = m_Parser->Get<DWORD>();
				tzRule.lDaylightBias = m_Parser->Get<DWORD>();
				tzRule.stStandardDate.wYear = m_Parser->Get<WORD>();
				tzRule.stStandardDate.wMonth = m_Parser->Get<WORD>();
				tzRule.stStandardDate.wDayOfWeek = m_Parser->Get<WORD>();
				tzRule.stStandardDate.wDay = m_Parser->Get<WORD>();
				tzRule.stStandardDate.wHour = m_Parser->Get<WORD>();
				tzRule.stStandardDate.wMinute = m_Parser->Get<WORD>();
				tzRule.stStandardDate.wSecond = m_Parser->Get<WORD>();
				tzRule.stStandardDate.wMilliseconds = m_Parser->Get<WORD>();
				tzRule.stDaylightDate.wYear = m_Parser->Get<WORD>();
				tzRule.stDaylightDate.wMonth = m_Parser->Get<WORD>();
				tzRule.stDaylightDate.wDayOfWeek = m_Parser->Get<WORD>();
				tzRule.stDaylightDate.wDay = m_Parser->Get<WORD>();
				tzRule.stDaylightDate.wHour = m_Parser->Get<WORD>();
				tzRule.stDaylightDate.wMinute = m_Parser->Get<WORD>();
				tzRule.stDaylightDate.wSecond = m_Parser->Get<WORD>();
				tzRule.stDaylightDate.wMilliseconds = m_Parser->Get<WORD>();
				m_lpTZRule.push_back(tzRule);
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

		for (WORD i = 0; i < m_lpTZRule.size(); i++)
		{
			terminateBlock();
			addBlankLine();
			addBlock(
				m_lpTZRule[i].bMajorVersion,
				L"TZRule[0x%1!X!].bMajorVersion = 0x%2!02X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].bMajorVersion.getData());
			addBlock(
				m_lpTZRule[i].bMinorVersion,
				L"TZRule[0x%1!X!].bMinorVersion = 0x%2!02X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].bMinorVersion.getData());
			addBlock(
				m_lpTZRule[i].wReserved,
				L"TZRule[0x%1!X!].wReserved = 0x%2!04X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].wReserved.getData());
			addBlock(
				m_lpTZRule[i].wTZRuleFlags,
				L"TZRule[0x%1!X!].wTZRuleFlags = 0x%2!04X! = %3!ws!\r\n",
				i,
				m_lpTZRule[i].wTZRuleFlags.getData(),
				flags::InterpretFlags(flagTZRule, m_lpTZRule[i].wTZRuleFlags).c_str());
			addBlock(
				m_lpTZRule[i].wYear,
				L"TZRule[0x%1!X!].wYear = 0x%2!04X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].wYear.getData());
			addHeader(L"TZRule[0x%1!X!].X = ", i);
			addBlock(m_lpTZRule[i].X);

			terminateBlock();
			addBlock(
				m_lpTZRule[i].lBias,
				L"TZRule[0x%1!X!].lBias = 0x%2!08X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].lBias.getData());
			addBlock(
				m_lpTZRule[i].lStandardBias,
				L"TZRule[0x%1!X!].lStandardBias = 0x%2!08X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].lStandardBias.getData());
			addBlock(
				m_lpTZRule[i].lDaylightBias,
				L"TZRule[0x%1!X!].lDaylightBias = 0x%2!08X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].lDaylightBias.getData());
			addBlankLine();
			addBlock(
				m_lpTZRule[i].stStandardDate.wYear,
				L"TZRule[0x%1!X!].stStandardDate.wYear = 0x%2!X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].stStandardDate.wYear.getData());
			addBlock(
				m_lpTZRule[i].stStandardDate.wMonth,
				L"TZRule[0x%1!X!].stStandardDate.wMonth = 0x%2!X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].stStandardDate.wMonth.getData());
			addBlock(
				m_lpTZRule[i].stStandardDate.wDayOfWeek,
				L"TZRule[0x%1!X!].stStandardDate.wDayOfWeek = 0x%2!X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].stStandardDate.wDayOfWeek.getData());
			addBlock(
				m_lpTZRule[i].stStandardDate.wDay,
				L"TZRule[0x%1!X!].stStandardDate.wDay = 0x%2!X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].stStandardDate.wDay.getData());
			addBlock(
				m_lpTZRule[i].stStandardDate.wHour,
				L"TZRule[0x%1!X!].stStandardDate.wHour = 0x%2!X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].stStandardDate.wHour.getData());
			addBlock(
				m_lpTZRule[i].stStandardDate.wMinute,
				L"TZRule[0x%1!X!].stStandardDate.wMinute = 0x%2!X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].stStandardDate.wMinute.getData());
			addBlock(
				m_lpTZRule[i].stStandardDate.wSecond,
				L"TZRule[0x%1!X!].stStandardDate.wSecond = 0x%2!X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].stStandardDate.wSecond.getData());
			addBlock(
				m_lpTZRule[i].stStandardDate.wMilliseconds,
				L"TZRule[0x%1!X!].stStandardDate.wMilliseconds = 0x%2!X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].stStandardDate.wMilliseconds.getData());
			addBlankLine();
			addBlock(
				m_lpTZRule[i].stDaylightDate.wYear,
				L"TZRule[0x%1!X!].stDaylightDate.wYear = 0x%2!X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].stDaylightDate.wYear.getData());
			addBlock(
				m_lpTZRule[i].stDaylightDate.wMonth,
				L"TZRule[0x%1!X!].stDaylightDate.wMonth = 0x%2!X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].stDaylightDate.wMonth.getData());
			addBlock(
				m_lpTZRule[i].stDaylightDate.wDayOfWeek,
				L"TZRule[0x%1!X!].stDaylightDate.wDayOfWeek = 0x%2!X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].stDaylightDate.wDayOfWeek.getData());
			addBlock(
				m_lpTZRule[i].stDaylightDate.wDay,
				L"TZRule[0x%1!X!].stDaylightDate.wDay = 0x%2!X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].stDaylightDate.wDay.getData());
			addBlock(
				m_lpTZRule[i].stDaylightDate.wHour,
				L"TZRule[0x%1!X!].stDaylightDate.wHour = 0x%2!X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].stDaylightDate.wHour.getData());
			addBlock(
				m_lpTZRule[i].stDaylightDate.wMinute,
				L"TZRule[0x%1!X!].stDaylightDate.wMinute = 0x%2!X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].stDaylightDate.wMinute.getData());
			addBlock(
				m_lpTZRule[i].stDaylightDate.wSecond,
				L"TZRule[0x%1!X!].stDaylightDate.wSecond = 0x%2!X! (%2!d!)\r\n",
				i,
				m_lpTZRule[i].stDaylightDate.wSecond.getData());
			addBlock(
				m_lpTZRule[i].stDaylightDate.wMilliseconds,
				L"TZRule[0x%1!X!].stDaylightDate.wMilliseconds = 0x%2!X! (%2!d!)",
				i,
				m_lpTZRule[i].stDaylightDate.wMilliseconds.getData());
		}
	}
} // namespace smartview