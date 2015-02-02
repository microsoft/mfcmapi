#pragma once
#include <string>
using namespace std;

wstring InterpretPropSmartView(_In_ LPSPropValue lpProp, // required property value
	_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
	_In_opt_ LPMAPINAMEID lpNameID, // optional named property information to avoid GetNamesFromIDs call
	_In_opt_ LPSBinary lpMappingSignature, // optional mapping signature for object to speed named prop lookups
	bool bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
	bool bMVRow); // did the row come from a MV prop?

wstring InterpretBinaryAsString(SBinary myBin, __ParsingTypeEnum iStructType, _In_opt_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag);
wstring InterpretMVBinaryAsString(SBinaryArray myBinArray, __ParsingTypeEnum iStructType, _In_opt_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag);
wstring InterpretNumberAsString(_PV pV, ULONG ulPropTag, ULONG ulPropNameID, _In_opt_z_ LPWSTR lpszPropNameString, _In_opt_ LPCGUID lpguidNamedProp, bool bLabel);
wstring InterpretNumberAsStringProp(ULONG ulVal, ULONG ulPropTag);
wstring InterpretNumberAsStringNamedProp(ULONG ulVal, ULONG ulPropNameID, _In_opt_ LPCGUID lpguidNamedProp);
_Check_return_ wstring FidMidToSzString(LONGLONG llID, bool bLabel);

_Check_return_ __ParsingTypeEnum FindSmartViewParserForProp(const ULONG ulPropTag, const ULONG ulPropNameID, _In_opt_ const LPCGUID lpguidNamedProp, bool bMVRow);

// Nothing below this point actually needs to be public. It's only used internally by InterpretPropSmartView

// Functions to parse PT_BINARY properties

// [MS-OXOCAL].pdf
// PatternTypeSpecificStruct
// =====================
//   This structure specifies the details of the recurrence type
//
union PatternTypeSpecificStruct
{
	DWORD WeekRecurrencePattern;
	DWORD MonthRecurrencePattern;
	struct
	{
		DWORD DayOfWeek;
		DWORD N;
	} MonthNthRecurrencePattern;
};

// RecurrencePatternStruct
// =====================
//   This structure specifies a recurrence pattern.
//
struct RecurrencePatternStruct
{
	WORD ReaderVersion;
	WORD WriterVersion;
	WORD RecurFrequency;
	WORD PatternType;
	WORD CalendarType;
	DWORD FirstDateTime;
	DWORD Period;
	DWORD SlidingFlag;
	PatternTypeSpecificStruct PatternTypeSpecific;
	DWORD EndType;
	DWORD OccurrenceCount;
	DWORD FirstDOW;
	DWORD DeletedInstanceCount;
	DWORD* DeletedInstanceDates;
	DWORD ModifiedInstanceCount;
	DWORD* ModifiedInstanceDates;
	DWORD StartDate;
	DWORD EndDate;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeleteRecurrencePatternStruct.
_Check_return_ RecurrencePatternStruct* BinToRecurrencePatternStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
_Check_return_ RecurrencePatternStruct* BinToRecurrencePatternStructWithSize(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _Out_opt_ size_t* lpcbBytesRead);
void DeleteRecurrencePatternStruct(_In_ RecurrencePatternStruct* prpPattern);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR RecurrencePatternStructToString(_In_ RecurrencePatternStruct* prpPattern);

// ExceptionInfoStruct
// =====================
//   This structure specifies an exception
//
struct ExceptionInfoStruct
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

struct ChangeHighlightStruct
{
	DWORD ChangeHighlightSize;
	DWORD ChangeHighlightValue;
	LPBYTE Reserved;
};

// ExtendedExceptionStruct
// =====================
//   This structure specifies additional information about an exception
//
struct ExtendedExceptionStruct
{
	ChangeHighlightStruct ChangeHighlight;
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

// AppointmentRecurrencePatternStruct
// =====================
//   This structure specifies a recurrence pattern for a calendar object
//   including information about exception property values.
//
struct AppointmentRecurrencePatternStruct
{
	RecurrencePatternStruct* RecurrencePattern;
	DWORD ReaderVersion2;
	DWORD WriterVersion2;
	DWORD StartTimeOffset;
	DWORD EndTimeOffset;
	WORD ExceptionCount;
	ExceptionInfoStruct* ExceptionInfo;
	DWORD ReservedBlock1Size;
	LPBYTE ReservedBlock1;
	ExtendedExceptionStruct* ExtendedException;
	DWORD ReservedBlock2Size;
	LPBYTE ReservedBlock2;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeleteAppointmentRecurrencePatternStruct.
_Check_return_ AppointmentRecurrencePatternStruct* BinToAppointmentRecurrencePatternStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteAppointmentRecurrencePatternStruct(_In_ AppointmentRecurrencePatternStruct* parpPattern);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR AppointmentRecurrencePatternStructToString(_In_ AppointmentRecurrencePatternStruct* parpPattern);

void SDBinToString(SBinary myBin, _In_opt_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_z_ LPWSTR* lpszResultString);
void SIDBinToString(SBinary myBin, _Deref_out_z_ LPWSTR* lpszResultString);

// ExtendedFlagStruct
// =====================
//   This structure specifies an Extended Flag
//
struct ExtendedFlagStruct
{
	BYTE Id;
	BYTE Cb;
	union
	{
		DWORD ExtendedFlags;
		GUID SearchFolderID;
		DWORD SearchFolderTag;
		DWORD ToDoFolderVersion;
	} Data;
	BYTE* lpUnknownData;
};

// ExtendedFlagsStruct
// =====================
//   This structure specifies an array of Extended Flags
//
struct ExtendedFlagsStruct
{
	ULONG ulNumFlags;
	ExtendedFlagStruct* pefExtendedFlags;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeleteExtendedFlagsStruct.
_Check_return_ ExtendedFlagsStruct* BinToExtendedFlagsStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteExtendedFlagsStruct(_In_ ExtendedFlagsStruct* pefExtendedFlags);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR ExtendedFlagsStructToString(_In_ ExtendedFlagsStruct* pefExtendedFlags);