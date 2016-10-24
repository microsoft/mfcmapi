#include "stdafx.h"
#include "TimeZoneDefinition.h"
#include "InterpretProp2.h"
#include "ExtraPropTags.h"

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
	m_Parser.GetBYTE(&m_bMajorVersion);
	m_Parser.GetBYTE(&m_bMinorVersion);
	m_Parser.GetWORD(&m_cbHeader);
	m_Parser.GetWORD(&m_wReserved);
	m_Parser.GetWORD(&m_cchKeyName);
	m_szKeyName = m_Parser.GetStringW(m_cchKeyName);
	m_Parser.GetWORD(&m_cRules);

	if (m_cRules && m_cRules < _MaxEntriesSmall)
	{
		for (ULONG i = 0; i < m_cRules; i++)
		{
			TZRule tzRule;
			m_Parser.GetBYTE(&tzRule.bMajorVersion);
			m_Parser.GetBYTE(&tzRule.bMinorVersion);
			m_Parser.GetWORD(&tzRule.wReserved);
			m_Parser.GetWORD(&tzRule.wTZRuleFlags);
			m_Parser.GetWORD(&tzRule.wYear);
			tzRule.X = m_Parser.GetBYTES(14);
			m_Parser.GetDWORD(&tzRule.lBias);
			m_Parser.GetDWORD(&tzRule.lStandardBias);
			m_Parser.GetDWORD(&tzRule.lDaylightBias);
			m_Parser.GetWORD(&tzRule.stStandardDate.wYear);
			m_Parser.GetWORD(&tzRule.stStandardDate.wMonth);
			m_Parser.GetWORD(&tzRule.stStandardDate.wDayOfWeek);
			m_Parser.GetWORD(&tzRule.stStandardDate.wDay);
			m_Parser.GetWORD(&tzRule.stStandardDate.wHour);
			m_Parser.GetWORD(&tzRule.stStandardDate.wMinute);
			m_Parser.GetWORD(&tzRule.stStandardDate.wSecond);
			m_Parser.GetWORD(&tzRule.stStandardDate.wMilliseconds);
			m_Parser.GetWORD(&tzRule.stDaylightDate.wYear);
			m_Parser.GetWORD(&tzRule.stDaylightDate.wMonth);
			m_Parser.GetWORD(&tzRule.stDaylightDate.wDayOfWeek);
			m_Parser.GetWORD(&tzRule.stDaylightDate.wDay);
			m_Parser.GetWORD(&tzRule.stDaylightDate.wHour);
			m_Parser.GetWORD(&tzRule.stDaylightDate.wMinute);
			m_Parser.GetWORD(&tzRule.stDaylightDate.wSecond);
			m_Parser.GetWORD(&tzRule.stDaylightDate.wMilliseconds);
			m_lpTZRule.push_back(tzRule);
		}
	}
}

_Check_return_ wstring TimeZoneDefinition::ToStringInternal()
{
	auto szTimeZoneDefinition = formatmessage(IDS_TIMEZONEDEFINITION,
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
			szTimeZoneDefinition += formatmessage(IDS_TZRULEHEADER,
				i,
				m_lpTZRule[i].bMajorVersion,
				m_lpTZRule[i].bMinorVersion,
				m_lpTZRule[i].wReserved,
				m_lpTZRule[i].wTZRuleFlags,
				szFlags.c_str(),
				m_lpTZRule[i].wYear);

			szTimeZoneDefinition += BinToHexString(m_lpTZRule[i].X, true);

			szTimeZoneDefinition += formatmessage(IDS_TZRULEFOOTER,
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