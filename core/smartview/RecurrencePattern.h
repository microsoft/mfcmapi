#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	// [MS-OXOCAL].pdf
	// PatternTypeSpecific
	// =====================
	//   This structure specifies the details of the recurrence type
	//
	struct PatternTypeSpecific
	{
		std::shared_ptr<blockT<DWORD>> WeekRecurrencePattern = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> MonthRecurrencePattern = emptyT<DWORD>();
		struct
		{
			std::shared_ptr<blockT<DWORD>> DayOfWeek = emptyT<DWORD>();
			std::shared_ptr<blockT<DWORD>> N = emptyT<DWORD>();
		} MonthNthRecurrencePattern;
	};

	class RecurrencePattern : public block
	{
	public:
		std::shared_ptr<blockT<DWORD>> m_ModifiedInstanceCount;

	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<WORD>> m_ReaderVersion = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> m_WriterVersion = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> m_RecurFrequency = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> m_PatternType = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> m_CalendarType = emptyT<WORD>();
		std::shared_ptr<blockT<DWORD>> m_FirstDateTime = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_Period = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_SlidingFlag = emptyT<DWORD>();
		PatternTypeSpecific m_PatternTypeSpecific;
		std::shared_ptr<blockT<DWORD>> m_EndType = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_OccurrenceCount = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_FirstDOW = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_DeletedInstanceCount = emptyT<DWORD>();
		std::vector<std::shared_ptr<blockT<DWORD>>> m_DeletedInstanceDates;
		std::vector<std::shared_ptr<blockT<DWORD>>> m_ModifiedInstanceDates;
		std::shared_ptr<blockT<DWORD>> m_StartDate = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_EndDate = emptyT<DWORD>();
	};
} // namespace smartview