#include "stdafx.h"
#include "TimeZoneDefinition.h"
#include <Interpret/InterpretProp2.h>
#include <Interpret/ExtraPropTags.h>

TimeZoneDefinition::TimeZoneDefinition()
{
	m_bMajorVersion = 0;
	m_bMinorVersion = 0;
	m_cbHeader = 0;
	m_wReserved = 0;
	m_cchKeyName = 0;
	m_cRules = 0;
}

void TimeZoneDefinition::Parse()
{
	m_bMajorVersion = m_Parser.Get<BYTE>();
	m_bMinorVersion = m_Parser.Get<BYTE>();
	m_cbHeader = m_Parser.Get<WORD>();
	m_wReserved = m_Parser.Get<WORD>();
	m_cchKeyName = m_Parser.Get<WORD>();
	m_szKeyName = m_Parser.GetStringW(m_cchKeyName);
	m_cRules = m_Parser.Get<WORD>();

	if (m_cRules && m_cRules < _MaxEntriesSmall)
	{
		for (ULONG i = 0; i < m_cRules; i++)
		{
			TZRule tzRule;
			tzRule.bMajorVersion = m_Parser.Get<BYTE>();
			tzRule.bMinorVersion = m_Parser.Get<BYTE>();
			tzRule.wReserved = m_Parser.Get<WORD>();
			tzRule.wTZRuleFlags = m_Parser.Get<WORD>();
			tzRule.wYear = m_Parser.Get<WORD>();
			tzRule.X = m_Parser.GetBYTES(14);
			tzRule.lBias = m_Parser.Get<DWORD>();
			tzRule.lStandardBias = m_Parser.Get<DWORD>();
			tzRule.lDaylightBias = m_Parser.Get<DWORD>();
			tzRule.stStandardDate.wYear = m_Parser.Get<WORD>();
			tzRule.stStandardDate.wMonth = m_Parser.Get<WORD>();
			tzRule.stStandardDate.wDayOfWeek = m_Parser.Get<WORD>();
			tzRule.stStandardDate.wDay = m_Parser.Get<WORD>();
			tzRule.stStandardDate.wHour = m_Parser.Get<WORD>();
			tzRule.stStandardDate.wMinute = m_Parser.Get<WORD>();
			tzRule.stStandardDate.wSecond = m_Parser.Get<WORD>();
			tzRule.stStandardDate.wMilliseconds = m_Parser.Get<WORD>();
			tzRule.stDaylightDate.wYear = m_Parser.Get<WORD>();
			tzRule.stDaylightDate.wMonth = m_Parser.Get<WORD>();
			tzRule.stDaylightDate.wDayOfWeek = m_Parser.Get<WORD>();
			tzRule.stDaylightDate.wDay = m_Parser.Get<WORD>();
			tzRule.stDaylightDate.wHour = m_Parser.Get<WORD>();
			tzRule.stDaylightDate.wMinute = m_Parser.Get<WORD>();
			tzRule.stDaylightDate.wSecond = m_Parser.Get<WORD>();
			tzRule.stDaylightDate.wMilliseconds = m_Parser.Get<WORD>();
			m_lpTZRule.push_back(tzRule);
		}
	}
}

_Check_return_ std::wstring TimeZoneDefinition::ToStringInternal()
{
	auto szTimeZoneDefinition = strings::formatmessage(IDS_TIMEZONEDEFINITION,
		m_bMajorVersion,
		m_bMinorVersion,
		m_cbHeader,
		m_wReserved,
		m_cchKeyName,
		m_szKeyName.c_str(),
		m_cRules);

	{
		for (WORD i = 0; i < m_lpTZRule.size(); i++)
		{
			auto szFlags = InterpretFlags(flagTZRule, m_lpTZRule[i].wTZRuleFlags);
			szTimeZoneDefinition += strings::formatmessage(IDS_TZRULEHEADER,
				i,
				m_lpTZRule[i].bMajorVersion,
				m_lpTZRule[i].bMinorVersion,
				m_lpTZRule[i].wReserved,
				m_lpTZRule[i].wTZRuleFlags,
				szFlags.c_str(),
				m_lpTZRule[i].wYear);

			szTimeZoneDefinition += strings::BinToHexString(m_lpTZRule[i].X, true);

			szTimeZoneDefinition += strings::formatmessage(IDS_TZRULEFOOTER,
				i,
				m_lpTZRule[i].lBias,
				m_lpTZRule[i].lStandardBias,
				m_lpTZRule[i].lDaylightBias,
				m_lpTZRule[i].stStandardDate.wYear,
				m_lpTZRule[i].stStandardDate.wMonth,
				m_lpTZRule[i].stStandardDate.wDayOfWeek,
				m_lpTZRule[i].stStandardDate.wDay,
				m_lpTZRule[i].stStandardDate.wHour,
				m_lpTZRule[i].stStandardDate.wMinute,
				m_lpTZRule[i].stStandardDate.wSecond,
				m_lpTZRule[i].stStandardDate.wMilliseconds,
				m_lpTZRule[i].stDaylightDate.wYear,
				m_lpTZRule[i].stDaylightDate.wMonth,
				m_lpTZRule[i].stDaylightDate.wDayOfWeek,
				m_lpTZRule[i].stDaylightDate.wDay,
				m_lpTZRule[i].stDaylightDate.wHour,
				m_lpTZRule[i].stDaylightDate.wMinute,
				m_lpTZRule[i].stDaylightDate.wSecond,
				m_lpTZRule[i].stDaylightDate.wMilliseconds);
		}
	}

	return szTimeZoneDefinition;
}