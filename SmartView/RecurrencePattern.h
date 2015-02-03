#pragma once
#include "SmartViewParser.h"

// [MS-OXOCAL].pdf
// PatternTypeSpecific
// =====================
//   This structure specifies the details of the recurrence type
//
union PatternTypeSpecific
{
	DWORD WeekRecurrencePattern;
	DWORD MonthRecurrencePattern;
	struct
	{
		DWORD DayOfWeek;
		DWORD N;
	} MonthNthRecurrencePattern;
};


class RecurrencePattern : public SmartViewParser
{
public:
	RecurrencePattern(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
	~RecurrencePattern();

	DWORD m_ModifiedInstanceCount;

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	WORD m_ReaderVersion;
	WORD m_WriterVersion;
	WORD m_RecurFrequency;
	WORD m_PatternType;
	WORD m_CalendarType;
	DWORD m_FirstDateTime;
	DWORD m_Period;
	DWORD m_SlidingFlag;
	PatternTypeSpecific m_PatternTypeSpecific;
	DWORD m_EndType;
	DWORD m_OccurrenceCount;
	DWORD m_FirstDOW;
	DWORD m_DeletedInstanceCount;
	DWORD* m_DeletedInstanceDates;
	DWORD* m_ModifiedInstanceDates;
	DWORD m_StartDate;
	DWORD m_EndDate;
};