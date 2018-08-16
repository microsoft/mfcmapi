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

		switch (m_PatternType)
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

		if (m_DeletedInstanceCount && m_DeletedInstanceCount < _MaxEntriesSmall)
		{
			for (DWORD i = 0; i < m_DeletedInstanceCount; i++)
			{
				m_DeletedInstanceDates.push_back(m_Parser.GetBlock<DWORD>());
			}
		}

		m_ModifiedInstanceCount = m_Parser.GetBlock<DWORD>();

		if (m_ModifiedInstanceCount && m_ModifiedInstanceCount <= m_DeletedInstanceCount &&
			m_ModifiedInstanceCount < _MaxEntriesSmall)
		{
			for (DWORD i = 0; i < m_ModifiedInstanceCount; i++)
			{
				m_ModifiedInstanceDates.push_back(m_Parser.GetBlock<DWORD>());
			}
		}

		m_StartDate = m_Parser.GetBlock<DWORD>();
		m_EndDate = m_Parser.GetBlock<DWORD>();
	}

	void RecurrencePattern::ParseBlocks()
	{
		addHeader(L"Recurrence Pattern: \r\n");
		addBlock(m_ReaderVersion, L"ReaderVersion: 0x%1!04X!\r\n", m_ReaderVersion.getData());
		addBlock(m_WriterVersion, L"WriterVersion: 0x%1!04X!\r\n", m_WriterVersion.getData());
		auto szRecurFrequency = interpretprop::InterpretFlags(flagRecurFrequency, m_RecurFrequency);
		addBlock(
			m_RecurFrequency,
			L"RecurFrequency: 0x%1!04X! = %2!ws!\r\n",
			m_RecurFrequency.getData(),
			szRecurFrequency.c_str());
		auto szPatternType = interpretprop::InterpretFlags(flagPatternType, m_PatternType);
		addBlock(m_PatternType, L"PatternType: 0x%1!04X! = %2!ws!\r\n", m_PatternType.getData(), szPatternType.c_str());
		auto szCalendarType = interpretprop::InterpretFlags(flagCalendarType, m_CalendarType);
		addBlock(
			m_CalendarType, L"CalendarType: 0x%1!04X! = %2!ws!\r\n", m_CalendarType.getData(), szCalendarType.c_str());
		addBlock(m_FirstDateTime, L"FirstDateTime: 0x%1!08X! = %1!d!\r\n", m_FirstDateTime.getData());
		addBlock(m_Period, L"Period: 0x%1!08X! = %1!d!\r\n", m_Period.getData());
		addBlock(m_SlidingFlag, L"SlidingFlag: 0x%1!08X!\r\n", m_SlidingFlag.getData());

		switch (m_PatternType)
		{
		case rptMinute:
			break;
		case rptWeek:
			addBlock(
				m_PatternTypeSpecific.WeekRecurrencePattern,
				L"PatternTypeSpecific.WeekRecurrencePattern: 0x%1!08X! = %2!ws!\r\n",
				m_PatternTypeSpecific.WeekRecurrencePattern.getData(),
				interpretprop::InterpretFlags(flagDOW, m_PatternTypeSpecific.WeekRecurrencePattern).c_str());
			break;
		case rptMonth:
		case rptMonthEnd:
		case rptHjMonth:
		case rptHjMonthEnd:
			addBlock(
				m_PatternTypeSpecific.MonthRecurrencePattern,
				L"PatternTypeSpecific.MonthRecurrencePattern: 0x%1!08X! = %1!d!\r\n",
				m_PatternTypeSpecific.MonthRecurrencePattern.getData());
			break;
		case rptMonthNth:
		case rptHjMonthNth:
			addBlock(
				m_PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek,
				L"PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek: 0x%1!08X! = %2!ws!\r\n",
				m_PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek.getData(),
				interpretprop::InterpretFlags(flagDOW, m_PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek)
					.c_str());
			addBlock(
				m_PatternTypeSpecific.MonthNthRecurrencePattern.N,
				L"PatternTypeSpecific.MonthNthRecurrencePattern.N: 0x%1!08X! = %2!ws!\r\n",
				m_PatternTypeSpecific.MonthNthRecurrencePattern.N.getData(),
				interpretprop::InterpretFlags(flagN, m_PatternTypeSpecific.MonthNthRecurrencePattern.N).c_str());
			break;
		}

		addBlock(
			m_EndType,
			L"EndType: 0x%1!08X! = %2!ws!\r\n",
			m_EndType.getData(),
			interpretprop::InterpretFlags(flagEndType, m_EndType).c_str());
		addBlock(m_OccurrenceCount, L"OccurrenceCount: 0x%1!08X! = %1!d!\r\n", m_OccurrenceCount.getData());
		addBlock(
			m_FirstDOW,
			L"FirstDOW: 0x%1!08X! = %2!ws!\r\n",
			m_FirstDOW.getData(),
			interpretprop::InterpretFlags(flagFirstDOW, m_FirstDOW).c_str());
		addBlock(
			m_DeletedInstanceCount, L"DeletedInstanceCount: 0x%1!08X! = %1!d!\r\n", m_DeletedInstanceCount.getData());

		if (m_DeletedInstanceDates.size())
		{
			for (DWORD i = 0; i < m_DeletedInstanceDates.size(); i++)
			{
				addBlock(
					m_DeletedInstanceDates[i],
					L"DeletedInstanceDates[%1!d!]: 0x%2!08X! = %3!ws!\r\n",
					i,
					m_DeletedInstanceDates[i].getData(),
					RTimeToString(m_DeletedInstanceDates[i]).c_str());
			}
		}

		addBlock(
			m_ModifiedInstanceCount,
			L"ModifiedInstanceCount: 0x%1!08X! = %1!d!\r\n",
			m_ModifiedInstanceCount.getData());

		if (m_ModifiedInstanceDates.size())
		{
			for (DWORD i = 0; i < m_ModifiedInstanceDates.size(); i++)
			{
				addBlock(
					m_ModifiedInstanceDates[i],
					L"ModifiedInstanceDates[%1!d!]: 0x%2!08X! = %3!ws!\r\n",
					i,
					m_ModifiedInstanceDates[i].getData(),
					RTimeToString(m_ModifiedInstanceDates[i]).c_str());
			}
		}

		addBlock(
			m_StartDate,
			L"StartDate: 0x%1!08X! = %1!d! = %2!ws!\r\n",
			m_StartDate.getData(),
			RTimeToString(m_StartDate).c_str());
		addBlock(
			m_EndDate,
			L"EndDate: 0x%1!08X! = %1!d! = %2!ws!",
			m_EndDate.getData(),
			RTimeToString(m_EndDate).c_str());
	}
}