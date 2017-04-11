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
	string Subject;
	DWORD MeetingType;
	DWORD ReminderDelta;
	DWORD ReminderSet;
	WORD LocationLength;
	WORD LocationLength2;
	string Location;
	DWORD BusyStatus;
	DWORD Attachment;
	DWORD SubType;
	DWORD AppointmentColor;
};

struct ChangeHighlight
{
	DWORD ChangeHighlightSize;
	DWORD ChangeHighlightValue;
	vector<BYTE> Reserved;
};

// ExtendedException
// =====================
//   This structure specifies additional information about an exception
struct ExtendedException
{
	ChangeHighlight ChangeHighlight;
	DWORD ReservedBlockEE1Size;
	vector<BYTE> ReservedBlockEE1;
	DWORD StartDateTime;
	DWORD EndDateTime;
	DWORD OriginalStartDate;
	WORD WideCharSubjectLength;
	wstring WideCharSubject;
	WORD WideCharLocationLength;
	wstring WideCharLocation;
	DWORD ReservedBlockEE2Size;
	vector<BYTE> ReservedBlockEE2;
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
	_Check_return_ wstring ToStringInternal() override;

	RecurrencePattern m_RecurrencePattern;
	DWORD m_ReaderVersion2;
	DWORD m_WriterVersion2;
	DWORD m_StartTimeOffset;
	DWORD m_EndTimeOffset;
	WORD m_ExceptionCount;
	vector<ExceptionInfo> m_ExceptionInfo;
	DWORD m_ReservedBlock1Size;
	vector<BYTE> m_ReservedBlock1;
	vector<ExtendedException> m_ExtendedException;
	DWORD m_ReservedBlock2Size;
	vector<BYTE> m_ReservedBlock2;
};