#pragma once
#include "SmartViewParser.h"
#include "RecurrencePattern.h"

// ExceptionInfo
// =====================
//   This structure specifies an exception
struct ExceptionInfo
{
	DWORD StartDateTime;
	DWORD EndDateTime;
	DWORD OriginalStartDate;
	WORD OverrideFlags;
	WORD SubjectLength;
	WORD SubjectLength2;
	std::string Subject;
	DWORD MeetingType;
	DWORD ReminderDelta;
	DWORD ReminderSet;
	WORD LocationLength;
	WORD LocationLength2;
	std::string Location;
	DWORD BusyStatus;
	DWORD Attachment;
	DWORD SubType;
	DWORD AppointmentColor;
};

struct ChangeHighlight
{
	DWORD ChangeHighlightSize;
	DWORD ChangeHighlightValue;
	std::vector<BYTE> Reserved;
};

// ExtendedException
// =====================
//   This structure specifies additional information about an exception
struct ExtendedException
{
	ChangeHighlight ChangeHighlight;
	DWORD ReservedBlockEE1Size;
	std::vector<BYTE> ReservedBlockEE1;
	DWORD StartDateTime;
	DWORD EndDateTime;
	DWORD OriginalStartDate;
	WORD WideCharSubjectLength;
	std::wstring WideCharSubject;
	WORD WideCharLocationLength;
	std::wstring WideCharLocation;
	DWORD ReservedBlockEE2Size;
	std::vector<BYTE> ReservedBlockEE2;
};

// AppointmentRecurrencePattern
// =====================
//   This structure specifies a recurrence pattern for a calendar object
//   including information about exception property values.
class AppointmentRecurrencePattern : public SmartViewParser
{
public:
	AppointmentRecurrencePattern();

private:
	void Parse() override;
	_Check_return_ std::wstring ToStringInternal() override;

	RecurrencePattern m_RecurrencePattern;
	DWORD m_ReaderVersion2;
	DWORD m_WriterVersion2;
	DWORD m_StartTimeOffset;
	DWORD m_EndTimeOffset;
	WORD m_ExceptionCount;
	std::vector<ExceptionInfo> m_ExceptionInfo;
	DWORD m_ReservedBlock1Size;
	std::vector<BYTE> m_ReservedBlock1;
	std::vector<ExtendedException> m_ExtendedException;
	DWORD m_ReservedBlock2Size;
	std::vector<BYTE> m_ReservedBlock2;
};