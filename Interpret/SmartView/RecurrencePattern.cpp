#include "StdAfx.h"
#include "RecurrencePattern.h"
#include <Interpret/SmartView/SmartView.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>

namespace smartview
{
	RecurrencePattern::RecurrencePattern()
	{
		m_ReaderVersion = 0;
		m_WriterVersion = 0;
		m_RecurFrequency = 0;
		m_PatternType = 0;
		m_CalendarType = 0;
		m_FirstDateTime = 0;
		m_Period = 0;
		m_SlidingFlag = 0;
		m_PatternTypeSpecific = { 0 };
		m_EndType = 0;
		m_OccurrenceCount = 0;
		m_FirstDOW = 0;
		m_DeletedInstanceCount = 0;
		m_ModifiedInstanceCount = 0;
		m_StartDate = 0;
		m_EndDate = 0;

	}

	void RecurrencePattern::Parse()
	{
		m_ReaderVersion = m_Parser.Get<WORD>();
		m_WriterVersion = m_Parser.Get<WORD>();
		m_RecurFrequency = m_Parser.Get<WORD>();
		m_PatternType = m_Parser.Get<WORD>();
		m_CalendarType = m_Parser.Get<WORD>();
		m_FirstDateTime = m_Parser.Get<DWORD>();
		m_Period = m_Parser.Get<DWORD>();
		m_SlidingFlag = m_Parser.Get<DWORD>();

		switch (m_PatternType)
		{
		case rptMinute:
			break;
		case rptWeek:
			m_PatternTypeSpecific.WeekRecurrencePattern = m_Parser.Get<DWORD>();
			break;
		case rptMonth:
		case rptMonthEnd:
		case rptHjMonth:
		case rptHjMonthEnd:
			m_PatternTypeSpecific.MonthRecurrencePattern = m_Parser.Get<DWORD>();
			break;
		case rptMonthNth:
		case rptHjMonthNth:
			m_PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek = m_Parser.Get<DWORD>();
			m_PatternTypeSpecific.MonthNthRecurrencePattern.N = m_Parser.Get<DWORD>();
			break;
		}

		m_EndType = m_Parser.Get<DWORD>();
		m_OccurrenceCount = m_Parser.Get<DWORD>();
		m_FirstDOW = m_Parser.Get<DWORD>();
		m_DeletedInstanceCount = m_Parser.Get<DWORD>();

		if (m_DeletedInstanceCount && m_DeletedInstanceCount < _MaxEntriesSmall)
		{
			for (DWORD i = 0; i < m_DeletedInstanceCount; i++)
			{
				DWORD deletedInstanceDate = 0;
				deletedInstanceDate = m_Parser.Get<DWORD>();
				m_DeletedInstanceDates.push_back(deletedInstanceDate);
			}
		}

		m_ModifiedInstanceCount = m_Parser.Get<DWORD>();

		if (m_ModifiedInstanceCount &&
			m_ModifiedInstanceCount <= m_DeletedInstanceCount &&
			m_ModifiedInstanceCount < _MaxEntriesSmall)
		{
			for (DWORD i = 0; i < m_ModifiedInstanceCount; i++)
			{
				DWORD modifiedInstanceDate = 0;
				modifiedInstanceDate = m_Parser.Get<DWORD>();
				m_ModifiedInstanceDates.push_back(modifiedInstanceDate);
			}
		}

		m_StartDate = m_Parser.Get<DWORD>();
		m_EndDate = m_Parser.Get<DWORD>();
	}

	_Check_return_ std::wstring RecurrencePattern::ToStringInternal()
	{
		auto szRecurFrequency = InterpretFlags(flagRecurFrequency, m_RecurFrequency);
		auto szPatternType = InterpretFlags(flagPatternType, m_PatternType);
		auto szCalendarType = InterpretFlags(flagCalendarType, m_CalendarType);
		auto szRP = strings::formatmessage(IDS_RPHEADER,
			m_ReaderVersion,
			m_WriterVersion,
			m_RecurFrequency, szRecurFrequency.c_str(),
			m_PatternType, szPatternType.c_str(),
			m_CalendarType, szCalendarType.c_str(),
			m_FirstDateTime,
			m_Period,
			m_SlidingFlag);

		std::wstring szDOW;
		std::wstring szN;
		switch (m_PatternType)
		{
		case rptMinute:
			break;
		case rptWeek:
			szDOW = InterpretFlags(flagDOW, m_PatternTypeSpecific.WeekRecurrencePattern);
			szRP += strings::formatmessage(IDS_RPPATTERNWEEK,
				m_PatternTypeSpecific.WeekRecurrencePattern, szDOW.c_str());
			break;
		case rptMonth:
		case rptMonthEnd:
		case rptHjMonth:
		case rptHjMonthEnd:
			szRP += strings::formatmessage(IDS_RPPATTERNMONTH,
				m_PatternTypeSpecific.MonthRecurrencePattern);
			break;
		case rptMonthNth:
		case rptHjMonthNth:
			szDOW = InterpretFlags(flagDOW, m_PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek);
			szN = InterpretFlags(flagN, m_PatternTypeSpecific.MonthNthRecurrencePattern.N);
			szRP += strings::formatmessage(IDS_RPPATTERNMONTHNTH,
				m_PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek, szDOW.c_str(),
				m_PatternTypeSpecific.MonthNthRecurrencePattern.N, szN.c_str());
			break;
		}

		auto szEndType = InterpretFlags(flagEndType, m_EndType);
		auto szFirstDOW = InterpretFlags(flagFirstDOW, m_FirstDOW);

		szRP += strings::formatmessage(IDS_RPHEADER2,
			m_EndType, szEndType.c_str(),
			m_OccurrenceCount,
			m_FirstDOW, szFirstDOW.c_str(),
			m_DeletedInstanceCount);

		if (m_DeletedInstanceDates.size())
		{
			for (DWORD i = 0; i < m_DeletedInstanceDates.size(); i++)
			{
				szRP += strings::formatmessage(IDS_RPDELETEDINSTANCEDATES,
					i, m_DeletedInstanceDates[i], RTimeToString(m_DeletedInstanceDates[i]).c_str());
			}
		}

		szRP += strings::formatmessage(IDS_RPMODIFIEDINSTANCECOUNT,
			m_ModifiedInstanceCount);

		if (m_ModifiedInstanceDates.size())
		{
			for (DWORD i = 0; i < m_ModifiedInstanceDates.size(); i++)
			{
				szRP += strings::formatmessage(IDS_RPMODIFIEDINSTANCEDATES,
					i, m_ModifiedInstanceDates[i], RTimeToString(m_ModifiedInstanceDates[i]).c_str());
			}
		}

		szRP += strings::formatmessage(IDS_RPDATE,
			m_StartDate, RTimeToString(m_StartDate).c_str(),
			m_EndDate, RTimeToString(m_EndDate).c_str());

		return szRP;
	}
}