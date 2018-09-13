#pragma once
#include <Interpret/SmartView/SmartViewParser.h>

namespace smartview
{
	// [MS-OXOCAL].pdf
	// PatternTypeSpecific
	// =====================
	//   This structure specifies the details of the recurrence type
	//
	struct PatternTypeSpecific
	{
		blockT<DWORD> WeekRecurrencePattern;
		blockT<DWORD> MonthRecurrencePattern;
		struct
		{
			blockT<DWORD> DayOfWeek;
			blockT<DWORD> N;
		} MonthNthRecurrencePattern;
	};

	class RecurrencePattern : public SmartViewParser
	{
	public:
		blockT<DWORD> m_ModifiedInstanceCount;

	private:
		void Parse() override;
		void ParseBlocks() override;

		blockT<WORD> m_ReaderVersion;
		blockT<WORD> m_WriterVersion;
		blockT<WORD> m_RecurFrequency;
		blockT<WORD> m_PatternType;
		blockT<WORD> m_CalendarType;
		blockT<DWORD> m_FirstDateTime;
		blockT<DWORD> m_Period;
		blockT<DWORD> m_SlidingFlag;
		PatternTypeSpecific m_PatternTypeSpecific;
		blockT<DWORD> m_EndType;
		blockT<DWORD> m_OccurrenceCount;
		blockT<DWORD> m_FirstDOW;
		blockT<DWORD> m_DeletedInstanceCount;
		std::vector<blockT<DWORD>> m_DeletedInstanceDates;
		std::vector<blockT<DWORD>> m_ModifiedInstanceDates;
		blockT<DWORD> m_StartDate;
		blockT<DWORD> m_EndDate;
	};
} // namespace smartview