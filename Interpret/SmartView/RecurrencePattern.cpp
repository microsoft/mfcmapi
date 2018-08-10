#include <StdAfx.h>
#include <Interpret/SmartView/RecurrencePattern.h>
#include <Interpret/SmartView/SmartView.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>

namespace smartview
{
	RecurrencePattern::RecurrencePattern() {}

	void RecurrencePattern::Parse()
	{
		m_ReaderVersion = m_Parser.GetBlock<WORD>();
		m_WriterVersion = m_Parser.GetBlock<WORD>();
		m_RecurFrequency = m_Parser.GetBlock<WORD>();
		m_PatternType = m_Parser.GetBlock<WORD>();
		m_CalendarType = m_Parser.GetBlock<WORD>();
		m_FirstDateTime = m_Parser.GetBlock<DWORD>();
		m_Period = m_Parser.GetBlock<DWORD>();
		m_SlidingFlag = m_Parser.GetBlock<DWORD>();

		switch (m_PatternType.getData())
		{
		case rptMinute:
			break;
		case rptWeek:
			m_PatternTypeSpecific.WeekRecurrencePattern = m_Parser.GetBlock<DWORD>();
			break;
		case rptMonth:
		case rptMonthEnd:
		case rptHjMonth:
		case rptHjMonthEnd:
			m_PatternTypeSpecific.MonthRecurrencePattern = m_Parser.GetBlock<DWORD>();
			break;
		case rptMonthNth:
		case rptHjMonthNth:
			m_PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek = m_Parser.GetBlock<DWORD>();
			m_PatternTypeSpecific.MonthNthRecurrencePattern.N = m_Parser.GetBlock<DWORD>();
			break;
		}

		m_EndType = m_Parser.GetBlock<DWORD>();
		m_OccurrenceCount = m_Parser.GetBlock<DWORD>();
		m_FirstDOW = m_Parser.GetBlock<DWORD>();
		m_DeletedInstanceCount = m_Parser.GetBlock<DWORD>();

		if (m_DeletedInstanceCount.getData() && m_DeletedInstanceCount.getData() < _MaxEntriesSmall)
		{
			for (DWORD i = 0; i < m_DeletedInstanceCount.getData(); i++)
			{
				m_DeletedInstanceDates.push_back(m_Parser.GetBlock<DWORD>());
			}
		}

		m_ModifiedInstanceCount = m_Parser.GetBlock<DWORD>();

		if (m_ModifiedInstanceCount.getData() &&
			m_ModifiedInstanceCount.getData() <= m_DeletedInstanceCount.getData() &&
			m_ModifiedInstanceCount.getData() < _MaxEntriesSmall)
		{
			for (DWORD i = 0; i < m_ModifiedInstanceCount.getData(); i++)
			{
				m_ModifiedInstanceDates.push_back(m_Parser.GetBlock<DWORD>());
			}
		}

		m_StartDate = m_Parser.GetBlock<DWORD>();
		m_EndDate = m_Parser.GetBlock<DWORD>();
	}

	_Check_return_ std::wstring RecurrencePattern::ToStringInternal()
	{
		auto szRecurFrequency = interpretprop::InterpretFlags(flagRecurFrequency, m_RecurFrequency.getData());
		auto szPatternType = interpretprop::InterpretFlags(flagPatternType, m_PatternType.getData());
		auto szCalendarType = interpretprop::InterpretFlags(flagCalendarType, m_CalendarType.getData());
		auto szRP = strings::formatmessage(
			IDS_RPHEADER,
			m_ReaderVersion.getData(),
			m_WriterVersion.getData(),
			m_RecurFrequency.getData(),
			szRecurFrequency.c_str(),
			m_PatternType.getData(),
			szPatternType.c_str(),
			m_CalendarType.getData(),
			szCalendarType.c_str(),
			m_FirstDateTime.getData(),
			m_Period.getData(),
			m_SlidingFlag.getData());

		std::wstring szDOW;
		std::wstring szN;
		switch (m_PatternType.getData())
		{
		case rptMinute:
			break;
		case rptWeek:
			szDOW = interpretprop::InterpretFlags(flagDOW, m_PatternTypeSpecific.WeekRecurrencePattern.getData());
			szRP += strings::formatmessage(
				IDS_RPPATTERNWEEK, m_PatternTypeSpecific.WeekRecurrencePattern.getData(), szDOW.c_str());
			break;
		case rptMonth:
		case rptMonthEnd:
		case rptHjMonth:
		case rptHjMonthEnd:
			szRP += strings::formatmessage(IDS_RPPATTERNMONTH, m_PatternTypeSpecific.MonthRecurrencePattern.getData());
			break;
		case rptMonthNth:
		case rptHjMonthNth:
			szDOW = interpretprop::InterpretFlags(
				flagDOW, m_PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek.getData());
			szN = interpretprop::InterpretFlags(flagN, m_PatternTypeSpecific.MonthNthRecurrencePattern.N.getData());
			szRP += strings::formatmessage(
				IDS_RPPATTERNMONTHNTH,
				m_PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek.getData(),
				szDOW.c_str(),
				m_PatternTypeSpecific.MonthNthRecurrencePattern.N.getData(),
				szN.c_str());
			break;
		}

		auto szEndType = interpretprop::InterpretFlags(flagEndType, m_EndType.getData());
		auto szFirstDOW = interpretprop::InterpretFlags(flagFirstDOW, m_FirstDOW.getData());

		szRP += strings::formatmessage(
			IDS_RPHEADER2,
			m_EndType.getData(),
			szEndType.c_str(),
			m_OccurrenceCount.getData(),
			m_FirstDOW.getData(),
			szFirstDOW.c_str(),
			m_DeletedInstanceCount.getData());

		if (m_DeletedInstanceDates.size())
		{
			for (DWORD i = 0; i < m_DeletedInstanceDates.size(); i++)
			{
				szRP += strings::formatmessage(
					IDS_RPDELETEDINSTANCEDATES,
					i,
					m_DeletedInstanceDates[i].getData(),
					RTimeToString(m_DeletedInstanceDates[i].getData()).c_str());
			}
		}

		szRP += strings::formatmessage(IDS_RPMODIFIEDINSTANCECOUNT, m_ModifiedInstanceCount.getData());

		if (m_ModifiedInstanceDates.size())
		{
			for (DWORD i = 0; i < m_ModifiedInstanceDates.size(); i++)
			{
				szRP += strings::formatmessage(
					IDS_RPMODIFIEDINSTANCEDATES,
					i,
					m_ModifiedInstanceDates[i].getData(),
					RTimeToString(m_ModifiedInstanceDates[i].getData()).c_str());
			}
		}

		szRP += strings::formatmessage(
			IDS_RPDATE,
			m_StartDate.getData(),
			RTimeToString(m_StartDate.getData()).c_str(),
			m_EndDate.getData(),
			RTimeToString(m_EndDate.getData()).c_str());

		return szRP;
	}
}