#include <StdAfx.h>
#include <Interpret/SmartView/TimeZoneDefinition.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>

namespace smartview
{
	TimeZoneDefinition::TimeZoneDefinition() {}

	void TimeZoneDefinition::Parse()
	{
		m_bMajorVersion = m_Parser.GetBlock<BYTE>();
		m_bMinorVersion = m_Parser.GetBlock<BYTE>();
		m_cbHeader = m_Parser.GetBlock<WORD>();
		m_wReserved = m_Parser.GetBlock<WORD>();
		m_cchKeyName = m_Parser.GetBlock<WORD>();
		m_szKeyName = m_Parser.GetBlockStringW(m_cchKeyName);
		m_cRules = m_Parser.GetBlock<WORD>();

		if (m_cRules && m_cRules < _MaxEntriesSmall)
		{
			for (ULONG i = 0; i < m_cRules; i++)
			{
				TZRule tzRule;
				tzRule.bMajorVersion = m_Parser.GetBlock<BYTE>();
				tzRule.bMinorVersion = m_Parser.GetBlock<BYTE>();
				tzRule.wReserved = m_Parser.GetBlock<WORD>();
				tzRule.wTZRuleFlags = m_Parser.GetBlock<WORD>();
				tzRule.wYear = m_Parser.GetBlock<WORD>();
				tzRule.X = m_Parser.GetBlockBYTES(14);
				tzRule.lBias = m_Parser.GetBlock<DWORD>();
				tzRule.lStandardBias = m_Parser.GetBlock<DWORD>();
				tzRule.lDaylightBias = m_Parser.GetBlock<DWORD>();
				tzRule.stStandardDate.wYear = m_Parser.GetBlock<WORD>();
				tzRule.stStandardDate.wMonth = m_Parser.GetBlock<WORD>();
				tzRule.stStandardDate.wDayOfWeek = m_Parser.GetBlock<WORD>();
				tzRule.stStandardDate.wDay = m_Parser.GetBlock<WORD>();
				tzRule.stStandardDate.wHour = m_Parser.GetBlock<WORD>();
				tzRule.stStandardDate.wMinute = m_Parser.GetBlock<WORD>();
				tzRule.stStandardDate.wSecond = m_Parser.GetBlock<WORD>();
				tzRule.stStandardDate.wMilliseconds = m_Parser.GetBlock<WORD>();
				tzRule.stDaylightDate.wYear = m_Parser.GetBlock<WORD>();
				tzRule.stDaylightDate.wMonth = m_Parser.GetBlock<WORD>();
				tzRule.stDaylightDate.wDayOfWeek = m_Parser.GetBlock<WORD>();
				tzRule.stDaylightDate.wDay = m_Parser.GetBlock<WORD>();
				tzRule.stDaylightDate.wHour = m_Parser.GetBlock<WORD>();
				tzRule.stDaylightDate.wMinute = m_Parser.GetBlock<WORD>();
				tzRule.stDaylightDate.wSecond = m_Parser.GetBlock<WORD>();
				tzRule.stDaylightDate.wMilliseconds = m_Parser.GetBlock<WORD>();
				m_lpTZRule.push_back(tzRule);
			}
		}
	}

	_Check_return_ std::wstring TimeZoneDefinition::ToStringInternal()
	{
		auto szTimeZoneDefinition = strings::formatmessage(
			IDS_TIMEZONEDEFINITION,
			m_bMajorVersion.getData(),
			m_bMinorVersion.getData(),
			m_cbHeader.getData(),
			m_wReserved.getData(),
			m_cchKeyName.getData(),
			m_szKeyName.c_str(),
			m_cRules.getData());

		for (WORD i = 0; i < m_lpTZRule.size(); i++)
		{
			auto szFlags = interpretprop::InterpretFlags(flagTZRule, m_lpTZRule[i].wTZRuleFlags);
			szTimeZoneDefinition += strings::formatmessage(
				IDS_TZRULEHEADER,
				i,
				m_lpTZRule[i].bMajorVersion.getData(),
				m_lpTZRule[i].bMinorVersion.getData(),
				m_lpTZRule[i].wReserved.getData(),
				m_lpTZRule[i].wTZRuleFlags.getData(),
				szFlags.c_str(),
				m_lpTZRule[i].wYear.getData());

			szTimeZoneDefinition += strings::BinToHexString(m_lpTZRule[i].X, true);

			szTimeZoneDefinition += strings::formatmessage(
				IDS_TZRULEFOOTER,
				i,
				m_lpTZRule[i].lBias.getData(),
				m_lpTZRule[i].lStandardBias.getData(),
				m_lpTZRule[i].lDaylightBias.getData(),
				m_lpTZRule[i].stStandardDate.wYear.getData(),
				m_lpTZRule[i].stStandardDate.wMonth.getData(),
				m_lpTZRule[i].stStandardDate.wDayOfWeek.getData(),
				m_lpTZRule[i].stStandardDate.wDay.getData(),
				m_lpTZRule[i].stStandardDate.wHour.getData(),
				m_lpTZRule[i].stStandardDate.wMinute.getData(),
				m_lpTZRule[i].stStandardDate.wSecond.getData(),
				m_lpTZRule[i].stStandardDate.wMilliseconds.getData(),
				m_lpTZRule[i].stDaylightDate.wYear.getData(),
				m_lpTZRule[i].stDaylightDate.wMonth.getData(),
				m_lpTZRule[i].stDaylightDate.wDayOfWeek.getData(),
				m_lpTZRule[i].stDaylightDate.wDay.getData(),
				m_lpTZRule[i].stDaylightDate.wHour.getData(),
				m_lpTZRule[i].stDaylightDate.wMinute.getData(),
				m_lpTZRule[i].stDaylightDate.wSecond.getData(),
				m_lpTZRule[i].stDaylightDate.wMilliseconds.getData());
		}

		return szTimeZoneDefinition;
	}
}