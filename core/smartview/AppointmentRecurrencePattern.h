#pragma once
#include <core/smartview/SmartViewParser.h>
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
	struct ExceptionInfo
	{
		blockT<DWORD> StartDateTime;
		blockT<DWORD> EndDateTime;
		blockT<DWORD> OriginalStartDate;
		blockT<WORD> OverrideFlags;
		blockT<WORD> SubjectLength;
		blockT<WORD> SubjectLength2;
		std::shared_ptr<blockStringA> Subject = emptySA();
		blockT<DWORD> MeetingType;
		blockT<DWORD> ReminderDelta;
		blockT<DWORD> ReminderSet;
		blockT<WORD> LocationLength;
		blockT<WORD> LocationLength2;
		std::shared_ptr<blockStringA> Location = emptySA();
		blockT<DWORD> BusyStatus;
		blockT<DWORD> Attachment;
		blockT<DWORD> SubType;
		blockT<DWORD> AppointmentColor;

		ExceptionInfo(std::shared_ptr<binaryParser>& parser);
	};

	struct ChangeHighlight
	{
		blockT<DWORD> ChangeHighlightSize;
		blockT<DWORD> ChangeHighlightValue;
		blockBytes Reserved;
	};

	// ExtendedException
	// =====================
	//   This structure specifies additional information about an exception
	struct ExtendedException
	{
		ChangeHighlight ChangeHighlight;
		blockT<DWORD> ReservedBlockEE1Size;
		blockBytes ReservedBlockEE1;
		blockT<DWORD> StartDateTime;
		blockT<DWORD> EndDateTime;
		blockT<DWORD> OriginalStartDate;
		blockT<WORD> WideCharSubjectLength;
		std::shared_ptr<blockStringW> WideCharSubject = emptySW();
		blockT<WORD> WideCharLocationLength;
		std::shared_ptr<blockStringW> WideCharLocation = emptySW();
		blockT<DWORD> ReservedBlockEE2Size;
		blockBytes ReservedBlockEE2;

		ExtendedException(std::shared_ptr<binaryParser>& parser, DWORD writerVersion2, WORD flags);
	};

	// AppointmentRecurrencePattern
	// =====================
	//   This structure specifies a recurrence pattern for a calendar object
	//   including information about exception property values.
	class AppointmentRecurrencePattern : public SmartViewParser
	{
	private:
		void Parse() override;
		void ParseBlocks() override;

		RecurrencePattern m_RecurrencePattern;
		blockT<DWORD> m_ReaderVersion2;
		blockT<DWORD> m_WriterVersion2;
		blockT<DWORD> m_StartTimeOffset;
		blockT<DWORD> m_EndTimeOffset;
		blockT<WORD> m_ExceptionCount;
		std::vector<std::shared_ptr<ExceptionInfo>> m_ExceptionInfo;
		blockT<DWORD> m_ReservedBlock1Size;
		blockBytes m_ReservedBlock1;
		std::vector<std::shared_ptr<ExtendedException>> m_ExtendedException;
		blockT<DWORD> m_ReservedBlock2Size;
		blockBytes m_ReservedBlock2;
	};
} // namespace smartview