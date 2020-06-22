#include <core/stdafx.h>
#include <core/smartview/RecurrencePattern.h>
#include <core/smartview/SmartView.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	void RecurrencePattern::parse()
	{
		m_ReaderVersion = blockT<WORD>::parse(parser);
		m_WriterVersion = blockT<WORD>::parse(parser);
		m_RecurFrequency = blockT<WORD>::parse(parser);
		m_PatternType = blockT<WORD>::parse(parser);
		m_CalendarType = blockT<WORD>::parse(parser);
		m_FirstDateTime = blockT<DWORD>::parse(parser);
		m_Period = blockT<DWORD>::parse(parser);
		m_SlidingFlag = blockT<DWORD>::parse(parser);

		switch (*m_PatternType)
		{
		case rptMinute:
			break;
		case rptWeek:
			m_PatternTypeSpecific.WeekRecurrencePattern = blockT<DWORD>::parse(parser);
			break;
		case rptMonth:
		case rptMonthEnd:
		case rptHjMonth:
		case rptHjMonthEnd:
			m_PatternTypeSpecific.MonthRecurrencePattern = blockT<DWORD>::parse(parser);
			break;
		case rptMonthNth:
		case rptHjMonthNth:
			m_PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek = blockT<DWORD>::parse(parser);
			m_PatternTypeSpecific.MonthNthRecurrencePattern.N = blockT<DWORD>::parse(parser);
			break;
		}

		m_EndType = blockT<DWORD>::parse(parser);
		m_OccurrenceCount = blockT<DWORD>::parse(parser);
		m_FirstDOW = blockT<DWORD>::parse(parser);
		m_DeletedInstanceCount = blockT<DWORD>::parse(parser);

		if (*m_DeletedInstanceCount && *m_DeletedInstanceCount < _MaxEntriesSmall)
		{
			m_DeletedInstanceDates.reserve(*m_DeletedInstanceCount);
			for (DWORD i = 0; i < *m_DeletedInstanceCount; i++)
			{
				m_DeletedInstanceDates.emplace_back(blockT<DWORD>::parse(parser));
			}
		}

		m_ModifiedInstanceCount = blockT<DWORD>::parse(parser);

		if (*m_ModifiedInstanceCount && *m_ModifiedInstanceCount <= *m_DeletedInstanceCount &&
			*m_ModifiedInstanceCount < _MaxEntriesSmall)
		{
			m_ModifiedInstanceDates.reserve(*m_ModifiedInstanceCount);
			for (DWORD i = 0; i < *m_ModifiedInstanceCount; i++)
			{
				m_ModifiedInstanceDates.emplace_back(blockT<DWORD>::parse(parser));
			}
		}

		m_StartDate = blockT<DWORD>::parse(parser);
		m_EndDate = blockT<DWORD>::parse(parser);
	}

	void RecurrencePattern::parseBlocks()
	{
		setText(L"Recurrence Pattern:");
		addChild(m_ReaderVersion, L"ReaderVersion: 0x%1!04X!", m_ReaderVersion->getData());
		addChild(m_WriterVersion, L"WriterVersion: 0x%1!04X!", m_WriterVersion->getData());
		auto szRecurFrequency = flags::InterpretFlags(flagRecurFrequency, *m_RecurFrequency);
		addChild(
			m_RecurFrequency,
			L"RecurFrequency: 0x%1!04X! = %2!ws!",
			m_RecurFrequency->getData(),
			szRecurFrequency.c_str());
		auto szPatternType = flags::InterpretFlags(flagPatternType, *m_PatternType);
		addChild(m_PatternType, L"PatternType: 0x%1!04X! = %2!ws!", m_PatternType->getData(), szPatternType.c_str());
		auto szCalendarType = flags::InterpretFlags(flagCalendarType, *m_CalendarType);
		addChild(
			m_CalendarType, L"CalendarType: 0x%1!04X! = %2!ws!", m_CalendarType->getData(), szCalendarType.c_str());
		addChild(m_FirstDateTime, L"FirstDateTime: 0x%1!08X! = %1!d!", m_FirstDateTime->getData());
		addChild(m_Period, L"Period: 0x%1!08X! = %1!d!", m_Period->getData());
		addChild(m_SlidingFlag, L"SlidingFlag: 0x%1!08X!", m_SlidingFlag->getData());

		switch (*m_PatternType)
		{
		case rptMinute:
			break;
		case rptWeek:
			addChild(
				m_PatternTypeSpecific.WeekRecurrencePattern,
				L"PatternTypeSpecific.WeekRecurrencePattern: 0x%1!08X! = %2!ws!",
				m_PatternTypeSpecific.WeekRecurrencePattern->getData(),
				flags::InterpretFlags(flagDOW, *m_PatternTypeSpecific.WeekRecurrencePattern).c_str());
			break;
		case rptMonth:
		case rptMonthEnd:
		case rptHjMonth:
		case rptHjMonthEnd:
			addChild(
				m_PatternTypeSpecific.MonthRecurrencePattern,
				L"PatternTypeSpecific.MonthRecurrencePattern: 0x%1!08X! = %1!d!",
				m_PatternTypeSpecific.MonthRecurrencePattern->getData());
			break;
		case rptMonthNth:
		case rptHjMonthNth:
			addChild(
				m_PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek,
				L"PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek: 0x%1!08X! = %2!ws!",
				m_PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek->getData(),
				flags::InterpretFlags(flagDOW, *m_PatternTypeSpecific.MonthNthRecurrencePattern.DayOfWeek).c_str());
			addChild(
				m_PatternTypeSpecific.MonthNthRecurrencePattern.N,
				L"PatternTypeSpecific.MonthNthRecurrencePattern.N: 0x%1!08X! = %2!ws!",
				m_PatternTypeSpecific.MonthNthRecurrencePattern.N->getData(),
				flags::InterpretFlags(flagN, *m_PatternTypeSpecific.MonthNthRecurrencePattern.N).c_str());
			break;
		}

		addChild(
			m_EndType,
			L"EndType: 0x%1!08X! = %2!ws!",
			m_EndType->getData(),
			flags::InterpretFlags(flagEndType, *m_EndType).c_str());
		addChild(m_OccurrenceCount, L"OccurrenceCount: 0x%1!08X! = %1!d!", m_OccurrenceCount->getData());
		addChild(
			m_FirstDOW,
			L"FirstDOW: 0x%1!08X! = %2!ws!",
			m_FirstDOW->getData(),
			flags::InterpretFlags(flagFirstDOW, *m_FirstDOW).c_str());

		m_DeletedInstanceCount->setText(L"DeletedInstanceCount: 0x%1!08X! = %1!d!", m_DeletedInstanceCount->getData());
		addChild(m_DeletedInstanceCount);

		if (m_DeletedInstanceDates.size())
		{
			auto i = 0;
			for (const auto& date : m_DeletedInstanceDates)
			{
				auto exception = create(L"DeletedInstanceDates[%1!d!]", i);
				m_DeletedInstanceCount->addChild(exception);
				exception->addChild(date, L"0x%1!08X! = %2!ws!", date->getData(), RTimeToString(*date).c_str());
				i++;
			}
		}

		addChild(
			m_ModifiedInstanceCount, L"ModifiedInstanceCount: 0x%1!08X! = %1!d!", m_ModifiedInstanceCount->getData());

		if (m_ModifiedInstanceDates.size())
		{
			auto i = 0;
			for (const auto& date : m_ModifiedInstanceDates)
			{
				auto modified = create(L"ModifiedInstanceDates[%1!d!]", i);
				m_ModifiedInstanceCount->addChild(modified);
				modified->addChild(
					date,
					L"0x%1!08X! = %2!ws!",
					date->getData(),
					RTimeToString(*date).c_str());
				i++;
			}
		}

		addChild(
			m_StartDate,
			L"StartDate: 0x%1!08X! = %1!d! = %2!ws!",
			m_StartDate->getData(),
			RTimeToString(*m_StartDate).c_str());
		addChild(
			m_EndDate, L"EndDate: 0x%1!08X! = %1!d! = %2!ws!", m_EndDate->getData(), RTimeToString(*m_EndDate).c_str());
	}
} // namespace smartview