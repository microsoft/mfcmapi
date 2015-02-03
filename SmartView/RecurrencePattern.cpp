#include "stdafx.h"
#include "..\stdafx.h"
#include "RecurrencePattern.h"
#include "..\String.h"
#include "..\ParseProperty.h"
#include "..\InterpretProp.h"
#include "..\InterpretProp2.h"
#include "..\ExtraPropTags.h"

RecurrencePattern::RecurrencePattern(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
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
	m_DeletedInstanceDates = 0;
	m_ModifiedInstanceCount = 0;
	m_ModifiedInstanceDates = 0;
	m_StartDate = 0;
	m_EndDate = 0;

}

RecurrencePattern::~RecurrencePattern()
{
	delete[] m_DeletedInstanceDates;
	delete[] m_ModifiedInstanceDates;
}

void RecurrencePattern::Parse()
{
	m_Parser.GetWORD(&m_ReaderVersion);
	m_Parser.GetWORD(&m_WriterVersion);
	m_Parser.GetWORD(&m_RecurFrequency);
	m_Parser.GetWORD(&m_PatternType);
	m_Parser.GetWORD(&m_CalendarType);
	m_Parser.GetDWORD(&m_FirstDateTime);
	m_Parser.GetDWORD(&m_Period);
	m_Parser.GetDWORD(&m_SlidingFlag);

	switch (m_PatternType)
	{
	case rptMinute:
		break;
	case rptWeek:
		m_Parser.GetDWORD(&m_PatternTypeSpecific.WeekRecurrencePattern);
		break;
	case rptMonth:
	case rptMonthEnd:
	case rptHjMonth:
	case rptHjMonthEnd:
		m_Parser.GetDWORD(&m_PatternTypeSpecific.MonthRecurrencePattern);
		break;
	case rptMonthNth:
	case rptHjMonthNth:
		m_Parser.GetDWORD(&m_PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek);
		m_Parser.GetDWORD(&m_PatternTypeSpecific.MonthNthRecurrencePattern.N);
		break;
	}

	m_Parser.GetDWORD(&m_EndType);
	m_Parser.GetDWORD(&m_OccurrenceCount);
	m_Parser.GetDWORD(&m_FirstDOW);
	m_Parser.GetDWORD(&m_DeletedInstanceCount);

	if (m_DeletedInstanceCount && m_DeletedInstanceCount < _MaxEntriesSmall)
	{
		m_DeletedInstanceDates = new DWORD[m_DeletedInstanceCount];
		if (m_DeletedInstanceDates)
		{
			memset(m_DeletedInstanceDates, 0, sizeof(DWORD)* m_DeletedInstanceCount);
			DWORD i = 0;
			for (i = 0; i < m_DeletedInstanceCount; i++)
			{
				m_Parser.GetDWORD(&m_DeletedInstanceDates[i]);
			}
		}
	}

	m_Parser.GetDWORD(&m_ModifiedInstanceCount);

	if (m_ModifiedInstanceCount &&
		m_ModifiedInstanceCount <= m_DeletedInstanceCount &&
		m_ModifiedInstanceCount < _MaxEntriesSmall)
	{
		m_ModifiedInstanceDates = new DWORD[m_ModifiedInstanceCount];
		if (m_ModifiedInstanceDates)
		{
			memset(m_ModifiedInstanceDates, 0, sizeof(DWORD)* m_ModifiedInstanceCount);
			DWORD i = 0;
			for (i = 0; i < m_ModifiedInstanceCount; i++)
			{
				m_Parser.GetDWORD(&m_ModifiedInstanceDates[i]);
			}
		}
	}
	m_Parser.GetDWORD(&m_StartDate);
	m_Parser.GetDWORD(&m_EndDate);
}

_Check_return_ wstring RecurrencePattern::ToStringInternal()
{
	wstring szRP;

	wstring szRecurFrequency = InterpretFlags(flagRecurFrequency, m_RecurFrequency);
	wstring szPatternType = InterpretFlags(flagPatternType, m_PatternType);
	wstring szCalendarType = InterpretFlags(flagCalendarType, m_CalendarType);
	szRP = formatmessage(IDS_RPHEADER,
		m_ReaderVersion,
		m_WriterVersion,
		m_RecurFrequency, szRecurFrequency.c_str(),
		m_PatternType, szPatternType.c_str(),
		m_CalendarType, szCalendarType.c_str(),
		m_FirstDateTime,
		m_Period,
		m_SlidingFlag);

	wstring szDOW;
	wstring szN;
	switch (m_PatternType)
	{
	case rptMinute:
		break;
	case rptWeek:
		szDOW = InterpretFlags(flagDOW, m_PatternTypeSpecific.WeekRecurrencePattern);
		szRP += formatmessage(IDS_RPPATTERNWEEK,
			m_PatternTypeSpecific.WeekRecurrencePattern, szDOW.c_str());
		break;
	case rptMonth:
	case rptMonthEnd:
	case rptHjMonth:
	case rptHjMonthEnd:
		szRP += formatmessage(IDS_RPPATTERNMONTH,
			m_PatternTypeSpecific.MonthRecurrencePattern);
		break;
	case rptMonthNth:
	case rptHjMonthNth:
		szDOW = InterpretFlags(flagDOW, m_PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek);
		szN = InterpretFlags(flagN, m_PatternTypeSpecific.MonthNthRecurrencePattern.N);
		szRP += formatmessage(IDS_RPPATTERNMONTHNTH,
			m_PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek, szDOW.c_str(),
			m_PatternTypeSpecific.MonthNthRecurrencePattern.N, szN.c_str());
		break;
	}


	wstring szEndType = InterpretFlags(flagEndType, m_EndType);
	wstring szFirstDOW = InterpretFlags(flagFirstDOW, m_FirstDOW);

	szRP += formatmessage(IDS_RPHEADER2,
		m_EndType, szEndType.c_str(),
		m_OccurrenceCount,
		m_FirstDOW, szFirstDOW.c_str(),
		m_DeletedInstanceCount);

	if (m_DeletedInstanceCount && m_DeletedInstanceDates)
	{
		DWORD i = 0;
		for (i = 0; i < m_DeletedInstanceCount; i++)
		{
			szRP += formatmessage(IDS_RPDELETEDINSTANCEDATES,
				i, m_DeletedInstanceDates[i], RTimeToString(m_DeletedInstanceDates[i]).c_str());
		}
	}

	szRP += formatmessage(IDS_RPMODIFIEDINSTANCECOUNT,
		m_ModifiedInstanceCount);

	if (m_ModifiedInstanceCount && m_ModifiedInstanceDates)
	{
		DWORD i = 0;
		for (i = 0; i < m_ModifiedInstanceCount; i++)
		{
			szRP += formatmessage(IDS_RPMODIFIEDINSTANCEDATES,
				i, m_ModifiedInstanceDates[i], RTimeToString(m_ModifiedInstanceDates[i]).c_str());
		}
	}

	szRP += formatmessage(IDS_RPDATE,
		m_StartDate, RTimeToString(m_StartDate).c_str(),
		m_EndDate, RTimeToString(m_EndDate).c_str());

	return szRP;
}