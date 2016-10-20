#include "stdafx.h"
#include "TimeZoneDefinition.h"
#include "String.h"
#include "InterpretProp2.h"
#include "ExtraPropTags.h"

TimeZoneDefinition::TimeZoneDefinition(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin)
{
	Init(cbBin, lpBin);
	m_bMajorVersion = 0;
	m_bMinorVersion = 0;
	m_cbHeader = 0;
	m_wReserved = 0;
	m_cchKeyName = 0;
	m_szKeyName = nullptr;
	m_cRules = 0;
	m_lpTZRule = nullptr;
}

TimeZoneDefinition::~TimeZoneDefinition()
{
	delete[] m_szKeyName;
	delete[] m_lpTZRule;
}

void TimeZoneDefinition::Parse()
{
	m_Parser.GetBYTE(&m_bMajorVersion);
	m_Parser.GetBYTE(&m_bMinorVersion);
	m_Parser.GetWORD(&m_cbHeader);
	m_Parser.GetWORD(&m_wReserved);
	m_Parser.GetWORD(&m_cchKeyName);
	m_Parser.GetStringW(m_cchKeyName, &m_szKeyName);
	m_Parser.GetWORD(&m_cRules);

	if (m_cRules && m_cRules < _MaxEntriesSmall)
		m_lpTZRule = new TZRule[m_cRules];

	if (m_lpTZRule)
	{
		memset(m_lpTZRule, 0, sizeof(TZRule)* m_cRules);
		for (ULONG i = 0; i < m_cRules; i++)
		{
			m_Parser.GetBYTE(&m_lpTZRule[i].bMajorVersion);
			m_Parser.GetBYTE(&m_lpTZRule[i].bMinorVersion);
			m_Parser.GetWORD(&m_lpTZRule[i].wReserved);
			m_Parser.GetWORD(&m_lpTZRule[i].wTZRuleFlags);
			m_Parser.GetWORD(&m_lpTZRule[i].wYear);
			m_Parser.GetBYTESNoAlloc(sizeof m_lpTZRule[i].X, sizeof m_lpTZRule[i].X, m_lpTZRule[i].X);
			m_Parser.GetDWORD(&m_lpTZRule[i].lBias);
			m_Parser.GetDWORD(&m_lpTZRule[i].lStandardBias);
			m_Parser.GetDWORD(&m_lpTZRule[i].lDaylightBias);
			m_Parser.GetWORD(&m_lpTZRule[i].stStandardDate.wYear);
			m_Parser.GetWORD(&m_lpTZRule[i].stStandardDate.wMonth);
			m_Parser.GetWORD(&m_lpTZRule[i].stStandardDate.wDayOfWeek);
			m_Parser.GetWORD(&m_lpTZRule[i].stStandardDate.wDay);
			m_Parser.GetWORD(&m_lpTZRule[i].stStandardDate.wHour);
			m_Parser.GetWORD(&m_lpTZRule[i].stStandardDate.wMinute);
			m_Parser.GetWORD(&m_lpTZRule[i].stStandardDate.wSecond);
			m_Parser.GetWORD(&m_lpTZRule[i].stStandardDate.wMilliseconds);
			m_Parser.GetWORD(&m_lpTZRule[i].stDaylightDate.wYear);
			m_Parser.GetWORD(&m_lpTZRule[i].stDaylightDate.wMonth);
			m_Parser.GetWORD(&m_lpTZRule[i].stDaylightDate.wDayOfWeek);
			m_Parser.GetWORD(&m_lpTZRule[i].stDaylightDate.wDay);
			m_Parser.GetWORD(&m_lpTZRule[i].stDaylightDate.wHour);
			m_Parser.GetWORD(&m_lpTZRule[i].stDaylightDate.wMinute);
			m_Parser.GetWORD(&m_lpTZRule[i].stDaylightDate.wSecond);
			m_Parser.GetWORD(&m_lpTZRule[i].stDaylightDate.wMilliseconds);
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
		m_szKeyName,
		m_cRules);

	if (m_cRules && m_lpTZRule)
	{
		for (WORD i = 0; i < m_cRules; i++)
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

			SBinary sBin = { 0 };
			sBin.cb = 14;
			sBin.lpb = m_lpTZRule[i].X;
			szTimeZoneDefinition += BinToHexString(&sBin, true);

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