#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/RecurrencePattern.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	// ExceptionInfo
	// =====================
	//   This structure specifies an exception
	class ExceptionInfo : public block
	{
	public:
		std::shared_ptr<blockT<WORD>> OverrideFlags = emptyT<WORD>();

	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<DWORD>> StartDateTime = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> EndDateTime = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> OriginalStartDate = emptyT<DWORD>();
		std::shared_ptr<blockT<WORD>> SubjectLength = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> SubjectLength2 = emptyT<WORD>();
		std::shared_ptr<blockStringA> Subject = emptySA();
		std::shared_ptr<blockT<DWORD>> MeetingType = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ReminderDelta = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ReminderSet = emptyT<DWORD>();
		std::shared_ptr<blockT<WORD>> LocationLength = emptyT<WORD>();
		std::shared_ptr<blockT<WORD>> LocationLength2 = emptyT<WORD>();
		std::shared_ptr<blockStringA> Location = emptySA();
		std::shared_ptr<blockT<DWORD>> BusyStatus = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> Attachment = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> SubType = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> AppointmentColor = emptyT<DWORD>();
	};

	class ChangeHighlight : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<DWORD>> ChangeHighlightSize = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> ChangeHighlightValue = emptyT<DWORD>();
		std::shared_ptr<blockBytes> Reserved = emptyBB();
	};

	// ExtendedException
	// =====================
	//   This structure specifies additional information about an exception
	class ExtendedException : public block
	{
	public:
		ExtendedException(DWORD _writerVersion2, WORD _flags) : writerVersion2(_writerVersion2), flags(_flags) {}

	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<ChangeHighlight> ChangeHighlight;
		std::shared_ptr<blockT<DWORD>> ReservedBlockEE1Size = emptyT<DWORD>();
		std::shared_ptr<blockBytes> ReservedBlockEE1 = emptyBB();
		std::shared_ptr<blockT<DWORD>> StartDateTime = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> EndDateTime = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> OriginalStartDate = emptyT<DWORD>();
		std::shared_ptr<blockT<WORD>> WideCharSubjectLength = emptyT<WORD>();
		std::shared_ptr<blockStringW> WideCharSubject = emptySW();
		std::shared_ptr<blockT<WORD>> WideCharLocationLength = emptyT<WORD>();
		std::shared_ptr<blockStringW> WideCharLocation = emptySW();
		std::shared_ptr<blockT<DWORD>> ReservedBlockEE2Size = emptyT<DWORD>();
		std::shared_ptr<blockBytes> ReservedBlockEE2 = emptyBB();

		DWORD writerVersion2;
		WORD flags;
	};

	// AppointmentRecurrencePattern
	// =====================
	//   This structure specifies a recurrence pattern for a calendar object
	//   including information about exception property values.
	class AppointmentRecurrencePattern : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<RecurrencePattern> m_RecurrencePattern;
		std::shared_ptr<blockT<DWORD>> m_ReaderVersion2 = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_WriterVersion2 = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_StartTimeOffset = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> m_EndTimeOffset = emptyT<DWORD>();
		std::shared_ptr<blockT<WORD>> m_ExceptionCount = emptyT<WORD>();
		std::vector<std::shared_ptr<ExceptionInfo>> m_ExceptionInfo;
		std::shared_ptr<blockT<DWORD>> m_ReservedBlock1Size = emptyT<DWORD>();
		std::shared_ptr<blockBytes> m_ReservedBlock1 = emptyBB();
		std::vector<std::shared_ptr<ExtendedException>> m_ExtendedException;
		std::shared_ptr<blockT<DWORD>> m_ReservedBlock2Size = emptyT<DWORD>();
		std::shared_ptr<blockBytes> m_ReservedBlock2 = emptyBB();
	};
} // namespace smartview