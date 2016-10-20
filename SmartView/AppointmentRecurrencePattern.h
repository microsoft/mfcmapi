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
	LPSTR Subject;
	DWORD MeetingType;
	DWORD ReminderDelta;
	DWORD ReminderSet;
	WORD LocationLength;
	WORD LocationLength2;
	LPSTR Location;
	DWORD BusyStatus;
	DWORD Attachment;
	DWORD SubType;
	DWORD AppointmentColor;
};

struct ChangeHighlight
{
	DWORD ChangeHighlightSize;
	DWORD ChangeHighlightValue;
	LPBYTE Reserved;
};

// ExtendedException
// =====================
//   This structure specifies additional information about an exception
struct ExtendedException
{
	ChangeHighlight ChangeHighlight;
	DWORD ReservedBlockEE1Size;
	LPBYTE ReservedBlockEE1;
	DWORD StartDateTime;
	DWORD EndDateTime;
	DWORD OriginalStartDate;
	WORD WideCharSubjectLength;
	LPWSTR WideCharSubject;
	WORD WideCharLocationLength;
	LPWSTR WideCharLocation;
	DWORD ReservedBlockEE2Size;
	LPBYTE ReservedBlockEE2;
};

// AppointmentRecurrencePattern
// =====================
//   This structure specifies a recurrence pattern for a calendar object
//   including information about exception property values.
class AppointmentRecurrencePattern : public SmartViewParser
{
public:
	AppointmentRecurrencePattern();
	~AppointmentRecurrencePattern();

private:
	void Parse();
	_Check_return_ wstring ToStringInternal();

	RecurrencePattern* m_RecurrencePattern;
	DWORD m_ReaderVersion2;
	DWORD m_WriterVersion2;
	DWORD m_StartTimeOffset;
	DWORD m_EndTimeOffset;
	WORD m_ExceptionCount;
	ExceptionInfo* m_ExceptionInfo;
	DWORD m_ReservedBlock1Size;
	LPBYTE m_ReservedBlock1;
	ExtendedException* m_ExtendedException;
	DWORD m_ReservedBlock2Size;
	LPBYTE m_ReservedBlock2;
};