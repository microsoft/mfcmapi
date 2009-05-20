#pragma once

// Function to convert property tags to their names
// Free lpszExactMatch and lpszPartialMatches with MAPIFreeBuffer
HRESULT PropTagToPropName(ULONG ulPropTag, BOOL bIsAB, LPTSTR* lpszExactMatch, LPTSTR* lpszPartialMatches);

HRESULT PropNameToPropTag(LPCTSTR lpszPropName, ULONG* ulPropTag);
HRESULT PropTypeNameToPropType(LPCTSTR lpszPropType, ULONG* ulPropType);
LPTSTR GUIDToStringAndName(LPCGUID lpGUID);
void GUIDNameToGUID(LPCTSTR szGUID, LPCGUID* lpGUID);

LPCWSTR NameIDToPropName(LPMAPINAMEID lpNameID);

HRESULT InterpretFlags(const LPSPropValue lpProp, LPTSTR* szFlagString);

HRESULT InterpretFlags(const ULONG ulFlagName, const LONG lFlagValue, LPTSTR* szFlagString);
HRESULT InterpretFlags(const ULONG ulFlagName, const LONG lFlagValue, LPCTSTR szPrefix, LPTSTR* szFlagString);
CString AllFlagsToString(const ULONG ulFlagName,BOOL bHex);

// Uber property interpreter - given an LPSPropValue, produces all manner of strings
// All LPTSTR strings allocated with new, delete with delete[]
void InterpretProp(LPSPropValue lpProp, // optional property value
				   ULONG ulPropTag, // optional 'original' prop tag
				   LPMAPIPROP lpMAPIProp, // optional source object
				   LPMAPINAMEID lpNameID, // optional named property information to avoid GetNamesFromIDs call
				   BOOL bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
				   LPTSTR* lpszNameExactMatches, // Built from ulPropTag & bIsAB
				   LPTSTR* lpszNamePartialMatches, // Built from ulPropTag & bIsAB
				   CString* PropType, // Built from ulPropTag
				   CString* PropTag, // Built from ulPropTag
				   CString* PropString, // Built from lpProp
				   CString* AltPropString, // Built from lpProp
				   LPTSTR* lpszSmartView, // Built from lpProp & lpMAPIProp
				   LPTSTR* lpszNamedPropName, // Built from ulPropTag & lpMAPIProp
				   LPTSTR* lpszNamedPropGUID, // Built from ulPropTag & lpMAPIProp
				   LPTSTR* lpszNamedPropDASL); // Built from ulPropTag & lpMAPIProp

extern UINT g_uidParsingTypesDropDown[];
extern ULONG g_cbuidParsingTypesDropDown;

void InterpretBinaryAsString(SBinary myBin, DWORD_PTR iStructType, LPMAPIPROP lpMAPIProp, ULONG ulPropTag, LPTSTR* lpszResultString);

HRESULT GetLargeBinaryProp(LPMAPIPROP lpMAPIProp, ULONG ulPropTag, LPSPropValue* lppProp);

// ExtendedFlagStruct
// =====================
//   This structure specifies an Extended Flag
//
typedef struct
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
} ExtendedFlagStruct;

// ExtendedFlagsStruct
// =====================
//   This structure specifies an array of Extended Flags
//
typedef struct
{
	ULONG ulNumFlags;
	ExtendedFlagStruct* pefExtendedFlags;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
} ExtendedFlagsStruct;

void ExtendedFlagsBinToString(SBinary myBin, LPTSTR* lpszResultString);
// Allocates return value with new. Clean up with DeleteExtendedFlagsStruct.
ExtendedFlagsStruct* BinToExtendedFlagsStruct(ULONG cbBin, LPBYTE lpBin);
void DeleteExtendedFlagsStruct(ExtendedFlagsStruct* pefExtendedFlags);
// result allocated with new, clean up with delete[]
LPTSTR ExtendedFlagsStructToString(ExtendedFlagsStruct* pefExtendedFlags);

void SDBinToString(SBinary myBin, LPMAPIPROP lpMAPIProp, ULONG ulPropTag, LPTSTR* lpszResultString);

// TimeZoneStruct
// =====================
//   This is an individual description that defines when a daylight
//   savings shift, and the return to standard time occurs, and how
//   far the shift is.  This is basically the same as
//   TIME_ZONE_INFORMATION documented in MSDN, except that the strings
//   describing the names 'daylight' and 'standard' time are omitted.
//
typedef struct
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
} TimeZoneStruct;

void TimeZoneToString(SBinary myBin, LPTSTR* lpszResultString);
// Allocates return value with new. Clean up with DeleteTimeZoneStruct.
TimeZoneStruct* BinToTimeZoneStruct(ULONG cbBin, LPBYTE lpBin);
void DeleteTimeZoneStruct(TimeZoneStruct* ptzTimeZone);
// result allocated with new, clean up with delete[]
LPTSTR TimeZoneStructToString(TimeZoneStruct* ptzTimeZone);

// TZRule
// =====================
//   This structure represents both a description when a daylight.
//   savings shift occurs, and in addition, the year in which that
//   timezone rule came into effect.
//
typedef struct
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
} TZRule;

// TimeZoneDefinitionStruct
// =====================
//   This represents an entire timezone including all historical, current
//   and future timezone shift rules for daylight savings time, etc.
//
typedef struct
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
} TimeZoneDefinitionStruct;

void TimeZoneDefinitionToString(SBinary myBin, LPTSTR* lpszResultString);
// Allocates return value with new. Clean up with DeleteTimeZoneDefinitionStruct.
TimeZoneDefinitionStruct* BinToTimeZoneDefinitionStruct(ULONG cbBin, LPBYTE lpBin);
void DeleteTimeZoneDefinitionStruct(TimeZoneDefinitionStruct* ptzdTimeZoneDefinition);
// result allocated with new, clean up with delete[]
LPTSTR TimeZoneDefinitionStructToString(TimeZoneDefinitionStruct* ptzdTimeZoneDefinition);

// [MS-OXOCAL].pdf
// PatternTypeSpecificStruct
// =====================
//   This structure specifies the details of the recurrence type
//
typedef union
{
	DWORD WeekRecurrencePattern;
	DWORD MonthRecurrencePattern;
	struct
	{
		DWORD DayOfWeek;
		DWORD N;
	} MonthNthRecurrencePattern;
} PatternTypeSpecificStruct;

typedef struct
{
	DWORD ChangeHighlightSize;
	DWORD ChangeHighlightValue;
	LPBYTE Reserved;
} ChangeHighlightStruct;

// RecurrencePatternStruct
// =====================
//   This structure specifies a recurrence pattern.
//
typedef struct
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
} RecurrencePatternStruct;

// ExceptionInfoStruct
// =====================
//   This structure specifies an exception
//
typedef struct
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
} ExceptionInfoStruct;

// ExtendedExceptionStruct
// =====================
//   This structure specifies additional information about an exception
//
typedef struct
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
} ExtendedExceptionStruct;

// AppointmentRecurrencePatternStruct
// =====================
//   This structure specifies a recurrence pattern for a calendar object
//   including information about exception property values.
//
typedef struct
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
} AppointmentRecurrencePatternStruct;

void AppointmentRecurrencePatternToString(SBinary myBin, LPTSTR* lpszResultString);
// Allocates return value with new. Clean up with DeleteAppointmentRecurrencePatternStruct.
AppointmentRecurrencePatternStruct* BinToAppointmentRecurrencePatternStruct(ULONG cbBin, LPBYTE lpBin);
void DeleteAppointmentRecurrencePatternStruct(AppointmentRecurrencePatternStruct* parpPattern);
// result allocated with new, clean up with delete[]
LPTSTR AppointmentRecurrencePatternStructToString(AppointmentRecurrencePatternStruct* parpPattern);

void RecurrencePatternToString(SBinary myBin, LPTSTR* lpszResultString);
// Allocates return value with new. Clean up with DeleteRecurrencePatternStruct.
RecurrencePatternStruct* BinToRecurrencePatternStruct(ULONG cbBin, LPBYTE lpBin, size_t* lpcbBytesRead);
void DeleteRecurrencePatternStruct(RecurrencePatternStruct* prpPattern);
// result allocated with new, clean up with delete[]
LPTSTR RecurrencePatternStructToString(RecurrencePatternStruct* prpPattern);

// [MS-OXOMSG].pdf
// ReportTagStruct
// =====================
//   This structure specifies a report tag for a mail object
//
typedef struct
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
} ReportTagStruct;

void ReportTagToString(SBinary myBin, LPTSTR* lpszResultString);
// Allocates return value with new. Clean up with DeleteReportTagStruct.
ReportTagStruct* BinToReportTagStruct(ULONG cbBin, LPBYTE lpBin);
void DeleteReportTagStruct(ReportTagStruct* prtReportTag);
// result allocated with new, clean up with delete[]
LPTSTR ReportTagStructToString(ReportTagStruct* prtReportTag);

// [MS-OXOMSG].pdf
// ResponseLevelStruct
// =====================
//   This structure specifies the response levels for a conversation index
//
typedef struct
{
	BOOL DeltaCode;
	DWORD TimeDelta;
	BYTE Random;
	BYTE ResponseLevel;
} ResponseLevelStruct;

// ConversationIndexStruct
// =====================
//   This structure specifies a report tag for a mail object
//
typedef struct
{
	BYTE UnnamedByte;
	FILETIME ftCurrent;
	GUID guid;
	ULONG ulResponseLevels;
	ResponseLevelStruct* lpResponseLevels;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
} ConversationIndexStruct;

void ConversationIndexToString(SBinary myBin, LPTSTR* lpszResultString);
// Allocates return value with new. Clean up with DeleteConversationIndexStruct.
ConversationIndexStruct* BinToConversationIndexStruct(ULONG cbBin, LPBYTE lpBin);
void DeleteConversationIndexStruct(ConversationIndexStruct* pciConversationIndex);
// result allocated with new, clean up with delete[]
LPTSTR ConversationIndexStructToString(ConversationIndexStruct* pciConversationIndex);


// [MS-OXOTASK].pdf

// TaskAssignerStruct
// =====================
//   This structure specifies single task assigner
//
typedef struct
{
	DWORD cbAssigner;
	ULONG cbEntryID;
	LPBYTE lpEntryID;
	LPSTR szDisplayName;
	LPWSTR wzDisplayName;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
} TaskAssignerStruct;

// TaskAssignersStruct
// =====================
//   This structure specifies an array of task assigners
//
typedef struct
{
	DWORD cAssigners;
	TaskAssignerStruct* lpTaskAssigners;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
} TaskAssignersStruct;

void TaskAssignersToString(SBinary myBin, LPTSTR* lpszResultString);
// Allocates return value with new. Clean up with DeleteTaskAssignersStruct.
TaskAssignersStruct* BinToTaskAssignersStruct(ULONG cbBin, LPBYTE lpBin);
void DeleteTaskAssignersStruct(TaskAssignersStruct* ptaTaskAssigners);
// result allocated with new, clean up with delete[]
LPTSTR TaskAssignersStructToString(TaskAssignersStruct* ptaTaskAssigners);

// GlobalObjectIdStruct
// =====================
//   This structure specifies a Global Object Id
//
typedef struct
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
} GlobalObjectIdStruct;

void GlobalObjectIdToString(SBinary myBin, LPTSTR* lpszResultString);
// Allocates return value with new. Clean up with DeleteGlobalObjectIdStruct.
GlobalObjectIdStruct* BinToGlobalObjectIdStruct(ULONG cbBin, LPBYTE lpBin);
void DeleteGlobalObjectIdStruct(GlobalObjectIdStruct* pgoidGlobalObjectId);
// result allocated with new, clean up with delete[]
LPTSTR GlobalObjectIdStructToString(GlobalObjectIdStruct* pgoidGlobalObjectId);

enum EIDStructType {
	eidtUnknown = 0,
	eidtShortTerm,
	eidtFolder,
	eidtMessage,
	eidtMessageDatabase,
	eidtOneOff,
	eidtAddressBook,
	eidtContact,
	eidtNewsGroupFolder
};
struct EntryIdStruct;
// EntryIdStruct
// =====================
//   This structure specifies an Entry Id
//
typedef struct EntryIdStruct
{
	BYTE abFlags[4];
	BYTE ProviderUID[16];
	ULONG ObjectType; // My own addition to simplify union parsing
	union
	{
		struct
		{
			WORD Type;
			union
			{
				struct
				{
					BYTE DatabaseGUID[16];
					BYTE GlobalCounter[6];
					BYTE Pad[2];
				} FolderObject;
				struct
				{
					BYTE FolderDatabaseGUID[16];
					BYTE FolderGlobalCounter[6];
					BYTE Pad1[2];
					BYTE MessageDatabaseGUID[16];
					BYTE MessageGlobalCounter[6];
					BYTE Pad2[2];
				} MessageObject;
			} Data;
		} FolderOrMessage;
		struct
		{
			BYTE Version;
			BYTE Flag;
			LPSTR DLLFileName;
			BOOL bIsExchange;
			ULONG WrappedFlags;
			BYTE WrappedProviderUID[16];
			ULONG WrappedType;
			LPSTR ServerShortname;
			LPSTR MailboxDN;
		} MessageDatabaseObject;
		struct
		{
			DWORD Bitmask;
			union
			{
				struct
				{
					LPWSTR DisplayName;
					LPWSTR AddressType;
					LPWSTR EmailAddress;
				} Unicode;
				struct
				{
					LPSTR DisplayName;
					LPSTR AddressType;
					LPSTR EmailAddress;
				} ANSI;
			} Strings;
		} OneOffRecipientObject;
		struct
		{
			DWORD Version;
			DWORD Type;
			LPSTR X500DN;
		} AddressBookObject;
		struct
		{
			DWORD Version;
			DWORD Type;
			DWORD Index;
			DWORD EntryIDCount;
			EntryIdStruct* lpEntryID;
		} ContactAddressBookObject;
	} ProviderData;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
} EntryIdStruct;

void EntryIdToString(SBinary myBin, LPTSTR* lpszResultString);
// Allocates return value with new. Clean up with DeleteEntryIdStruct.
EntryIdStruct* BinToEntryIdStruct(ULONG cbBin, LPBYTE lpBin);
void DeleteEntryIdStruct(EntryIdStruct* peidEntryId);
// result allocated with new, clean up with delete[]
LPTSTR EntryIdStructToString(EntryIdStruct* peidEntryId);

// EntryListEntryStruct
// =====================
//   This structure specifies an Entry in an Entry List
//
typedef struct EntryListEntryStruct
{
	DWORD EntryLength;
	DWORD EntryLengthPad;
	EntryIdStruct* EntryId;
} EntryListEntryStruct;


// EntryListStruct
// =====================
//   This structure specifies an Entry List
//
typedef struct EntryListStruct
{
	DWORD EntryCount;
	DWORD Pad;

	EntryListEntryStruct* Entry;

	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
} EntryListStruct;

void EntryListToString(SBinary myBin, LPTSTR* lpszResultString);
// Allocates return value with new. Clean up with DeleteEntryListStruct.
EntryListStruct* BinToEntryListStruct(ULONG cbBin, LPBYTE lpBin);
void DeleteEntryListStruct(EntryListStruct* pelEntryList);
// result allocated with new, clean up with delete[]
LPTSTR EntryListStructToString(EntryListStruct* pelEntryList);

// PropertyStruct
// =====================
//   This structure specifies an array of Properties
//
typedef struct
{
	DWORD PropCount;
	LPSPropValue Prop;

	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
} PropertyStruct;

void PropertyToString(SBinary myBin, LPTSTR* lpszResultString);
// Allocates return value with new. Clean up with DeletePropertyStruct.
PropertyStruct* BinToPropertyStruct(ULONG cbBin, LPBYTE lpBin, DWORD dwPropCount, size_t* lpcbBytesRead);
// Allocates return value with new. Clean up with DeletePropertyStruct.
LPSPropValue BinToSPropValue(ULONG cbBin, LPBYTE lpBin, DWORD dwPropCount, size_t* lpcbBytesRead);
// Neuters an array of SPropValues - caller must use delete to delete the SPropValue
void DeleteSPropVal(ULONG cVal, LPSPropValue lpsPropVal);
void DeletePropertyStruct(PropertyStruct* ppProperty);
// result allocated with new, clean up with delete[]
LPTSTR PropertyStructToString(PropertyStruct* ppProperty);

// RestrictionStruct
// =====================
//   This structure specifies a Restriction
//
typedef struct
{
	LPSRestriction lpRes;

	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
} RestrictionStruct;

void RestrictionToString(SBinary myBin, LPTSTR* lpszResultString);
// Allocates return value with new. Clean up with DeletePropertyStruct.
RestrictionStruct* BinToRestrictionStruct(ULONG cbBin, LPBYTE lpBin, size_t* lpcbBytesRead);
// Caller allocates with new. Clean up with DeleteRestrictionStruct.
void BinToRestrictionStruct(ULONG cbBin, LPBYTE lpBin, size_t* lpcbBytesRead, LPSRestriction psrRestriction);
// Neuters an SRestriction - caller must use delete to delete the SRestriction
void DeleteRestriction(LPSRestriction lpRes);
void DeleteRestrictionStruct(RestrictionStruct* prRestriction);
// result allocated with new, clean up with delete[]
LPTSTR RestrictionStructToString(RestrictionStruct* prRestriction);

// AddressListEntryStruct
// =====================
//   This structure specifies an entry in an Address List
//
typedef struct
{
	DWORD PropertyCount;
	DWORD Pad;
	PropertyStruct Properties;
} AddressListEntryStruct;

// SearchFolderDefinitionStruct
// =====================
//   This structure specifies a Search Folder Definition
//
typedef struct
{
	DWORD Version;
	DWORD Flags;
	DWORD NumericSearch;
	BYTE TextSearchLength;
	WORD TextSearchLengthExtended;
	LPWSTR TextSearch;
	DWORD SkipLen1;
	LPBYTE SkipBytes1;
	DWORD DeepSearch;
	BYTE FolderList1Length;
	WORD FolderList1LengthExtended;
	LPWSTR FolderList1;
	DWORD FolderList2Length;
	EntryListStruct* FolderList2;
	DWORD AddressCount; // SFST_BINARY
	AddressListEntryStruct* Addresses; // SFST_BINARY
	DWORD SkipLen2;
	LPBYTE SkipBytes2;
	RestrictionStruct* Restriction; // SFST_MRES
	DWORD AdvancedSearchLen; // SFST_FILTERSTREAM
	LPBYTE AdvancedSearchBytes; // SFST_FILTERSTREAM
	DWORD SkipLen3;
	LPBYTE SkipBytes3;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
} SearchFolderDefinitionStruct;

void SearchFolderDefinitionToString(SBinary myBin, LPTSTR* lpszResultString);
// Allocates return value with new. Clean up with DeleteSearchFolderDefinitionStruct.
SearchFolderDefinitionStruct* BinToSearchFolderDefinitionStruct(ULONG cbBin, LPBYTE lpBin);
void DeleteSearchFolderDefinitionStruct(SearchFolderDefinitionStruct* psfdSearchFolderDefinition);
// result allocated with new, clean up with delete[]
LPTSTR SearchFolderDefinitionStructToString(SearchFolderDefinitionStruct* psfdSearchFolderDefinition);