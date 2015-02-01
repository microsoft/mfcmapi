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

// TimeZoneStruct
// =====================
//   This is an individual description that defines when a daylight
//   savings shift, and the return to standard time occurs, and how
//   far the shift is.  This is basically the same as
//   TIME_ZONE_INFORMATION documented in MSDN, except that the strings
//   describing the names 'daylight' and 'standard' time are omitted.
//
struct TimeZoneStruct
{
	DWORD lBias; // offset from GMT
	DWORD lStandardBias; // offset from bias during standard time
	DWORD lDaylightBias; // offset from bias during daylight time
	WORD wStandardYear;
	SYSTEMTIME stStandardDate; // time to switch to standard time
	WORD wDaylightDate;
	SYSTEMTIME stDaylightDate; // time to switch to daylight time
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeleteTimeZoneStruct.
_Check_return_ TimeZoneStruct* BinToTimeZoneStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteTimeZoneStruct(_In_ TimeZoneStruct* ptzTimeZone);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR TimeZoneStructToString(_In_ TimeZoneStruct* ptzTimeZone);

// TZRule
// =====================
//   This structure represents both a description when a daylight.
//   savings shift occurs, and in addition, the year in which that
//   timezone rule came into effect.
//
struct TZRule
{
	BYTE bMajorVersion;
	BYTE bMinorVersion;
	WORD wReserved;
	WORD wTZRuleFlags;
	WORD wYear;
	BYTE X[14];
	DWORD lBias; // offset from GMT
	DWORD lStandardBias; // offset from bias during standard time
	DWORD lDaylightBias; // offset from bias during daylight time
	SYSTEMTIME stStandardDate; // time to switch to standard time
	SYSTEMTIME stDaylightDate; // time to switch to daylight time
};

// TimeZoneDefinitionStruct
// =====================
//   This represents an entire timezone including all historical, current
//   and future timezone shift rules for daylight savings time, etc.
//
struct TimeZoneDefinitionStruct
{
	BYTE bMajorVersion;
	BYTE bMinorVersion;
	WORD cbHeader;
	WORD wReserved;
	WORD cchKeyName;
	LPWSTR szKeyName;
	WORD cRules;
	TZRule* lpTZRule;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeleteTimeZoneDefinitionStruct.
_Check_return_ TimeZoneDefinitionStruct* BinToTimeZoneDefinitionStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteTimeZoneDefinitionStruct(_In_ TimeZoneDefinitionStruct* ptzdTimeZoneDefinition);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR TimeZoneDefinitionStructToString(_In_ TimeZoneDefinitionStruct* ptzdTimeZoneDefinition);

// [MS-OXOMSG].pdf
// ReportTagStruct
// =====================
//   This structure specifies a report tag for a mail object
//
struct ReportTagStruct
{
	CHAR Cookie[9]; // 8 characters + NULL terminator
	DWORD Version;
	ULONG cbStoreEntryID;
	LPBYTE lpStoreEntryID;
	ULONG cbFolderEntryID;
	LPBYTE lpFolderEntryID;
	ULONG cbMessageEntryID;
	LPBYTE lpMessageEntryID;
	ULONG cbSearchFolderEntryID;
	LPBYTE lpSearchFolderEntryID;
	ULONG cbMessageSearchKey;
	LPBYTE lpMessageSearchKey;
	ULONG cchAnsiText;
	LPSTR lpszAnsiText;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeleteReportTagStruct.
_Check_return_ ReportTagStruct* BinToReportTagStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteReportTagStruct(_In_ ReportTagStruct* prtReportTag);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR ReportTagStructToString(_In_ ReportTagStruct* prtReportTag);

// [MS-OXOMSG].pdf
// ResponseLevelStruct
// =====================
//   This structure specifies the response levels for a conversation index
//
struct ResponseLevelStruct
{
	bool DeltaCode;
	DWORD TimeDelta;
	BYTE Random;
	BYTE ResponseLevel;
};

// ConversationIndexStruct
// =====================
//   This structure specifies a report tag for a mail object
//
struct ConversationIndexStruct
{
	BYTE UnnamedByte;
	FILETIME ftCurrent;
	GUID guid;
	ULONG ulResponseLevels;
	ResponseLevelStruct* lpResponseLevels;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeleteConversationIndexStruct.
_Check_return_ ConversationIndexStruct* BinToConversationIndexStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteConversationIndexStruct(_In_ ConversationIndexStruct* pciConversationIndex);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR ConversationIndexStructToString(_In_ ConversationIndexStruct* pciConversationIndex);

// [MS-OXOTASK].pdf
// TaskAssignerStruct
// =====================
//   This structure specifies single task assigner
//
struct TaskAssignerStruct
{
	DWORD cbAssigner;
	ULONG cbEntryID;
	LPBYTE lpEntryID;
	LPSTR szDisplayName;
	LPWSTR wzDisplayName;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// TaskAssignersStruct
// =====================
//   This structure specifies an array of task assigners
//
struct TaskAssignersStruct
{
	DWORD cAssigners;
	TaskAssignerStruct* lpTaskAssigners;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeleteTaskAssignersStruct.
_Check_return_ TaskAssignersStruct* BinToTaskAssignersStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteTaskAssignersStruct(_In_ TaskAssignersStruct* ptaTaskAssigners);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR TaskAssignersStructToString(_In_ TaskAssignersStruct* ptaTaskAssigners);

// GlobalObjectIdStruct
// =====================
//   This structure specifies a Global Object Id
//
struct GlobalObjectIdStruct
{
	BYTE Id[16];
	WORD Year;
	BYTE Month;
	BYTE Day;
	FILETIME CreationTime;
	LARGE_INTEGER X;
	DWORD dwSize;
	LPBYTE lpData;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeleteGlobalObjectIdStruct.
_Check_return_ GlobalObjectIdStruct* BinToGlobalObjectIdStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteGlobalObjectIdStruct(_In_opt_ GlobalObjectIdStruct* pgoidGlobalObjectId);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR GlobalObjectIdStructToString(_In_ GlobalObjectIdStruct* pgoidGlobalObjectId);