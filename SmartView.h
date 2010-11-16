#pragma once

extern UINT g_uidParsingTypes[];
extern ULONG g_cuidParsingTypes;

// lpszSmartView allocated with new, delete with delete[]
void InterpretPropSmartView(_In_ LPSPropValue lpProp, // required property value
							_In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
							_In_opt_ LPMAPINAMEID lpNameID, // optional named property information to avoid GetNamesFromIDs call
							_In_opt_ LPSBinary lpMappingSignature, // optional mapping signature for object to speed named prop lookups
							BOOL bMVRow, // did the row come from a MV prop?
							_Deref_out_opt_z_ LPTSTR* lpszSmartView); // Built from lpProp & lpMAPIProp

void InterpretBinaryAsString(SBinary myBin, DWORD_PTR iStructType, _In_opt_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_z_ LPTSTR* lpszResultString);

// Nothing below this point actually needs to be public. It's only used internally by InterpretPropSmartView

// Functions to parse PT_LONG/PT-I2 properties

_Check_return_ CString RTimeToString(DWORD rTime);
_Check_return_ LPTSTR RTimeToSzString(DWORD rTime);

// End: Functions to parse PT_LONG/PT-I2 properties

// Functions to parse PT_BINARY properties

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

// Allocates return value with new. Clean up with DeleteRecurrencePatternStruct.
_Check_return_ RecurrencePatternStruct* BinToRecurrencePatternStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
_Check_return_ RecurrencePatternStruct* BinToRecurrencePatternStructWithSize(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _Out_opt_ size_t* lpcbBytesRead);
void DeleteRecurrencePatternStruct(_In_ RecurrencePatternStruct* prpPattern);
// result allocated with new, clean up with delete[]
_Check_return_ LPTSTR RecurrencePatternStructToString(_In_ RecurrencePatternStruct* prpPattern);

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

typedef struct
{
	DWORD ChangeHighlightSize;
	DWORD ChangeHighlightValue;
	LPBYTE Reserved;
} ChangeHighlightStruct;

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

// Allocates return value with new. Clean up with DeleteAppointmentRecurrencePatternStruct.
_Check_return_ AppointmentRecurrencePatternStruct* BinToAppointmentRecurrencePatternStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteAppointmentRecurrencePatternStruct(_In_ AppointmentRecurrencePatternStruct* parpPattern);
// result allocated with new, clean up with delete[]
_Check_return_ LPTSTR AppointmentRecurrencePatternStructToString(_In_ AppointmentRecurrencePatternStruct* parpPattern);

void SDBinToString(SBinary myBin, _In_opt_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_z_ LPTSTR* lpszResultString);
void SIDBinToString(SBinary myBin, _Deref_out_z_ LPTSTR* lpszResultString);

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

// Allocates return value with new. Clean up with DeleteExtendedFlagsStruct.
_Check_return_ ExtendedFlagsStruct* BinToExtendedFlagsStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteExtendedFlagsStruct(_In_ ExtendedFlagsStruct* pefExtendedFlags);
// result allocated with new, clean up with delete[]
_Check_return_ LPTSTR ExtendedFlagsStructToString(_In_ ExtendedFlagsStruct* pefExtendedFlags);

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

// Allocates return value with new. Clean up with DeleteTimeZoneStruct.
_Check_return_ TimeZoneStruct* BinToTimeZoneStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteTimeZoneStruct(_In_ TimeZoneStruct* ptzTimeZone);
// result allocated with new, clean up with delete[]
_Check_return_ LPTSTR TimeZoneStructToString(_In_ TimeZoneStruct* ptzTimeZone);

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

// Allocates return value with new. Clean up with DeleteTimeZoneDefinitionStruct.
_Check_return_ TimeZoneDefinitionStruct* BinToTimeZoneDefinitionStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteTimeZoneDefinitionStruct(_In_ TimeZoneDefinitionStruct* ptzdTimeZoneDefinition);
// result allocated with new, clean up with delete[]
_Check_return_ LPTSTR TimeZoneDefinitionStructToString(_In_ TimeZoneDefinitionStruct* ptzdTimeZoneDefinition);

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

// Allocates return value with new. Clean up with DeleteReportTagStruct.
_Check_return_ ReportTagStruct* BinToReportTagStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteReportTagStruct(_In_ ReportTagStruct* prtReportTag);
// result allocated with new, clean up with delete[]
_Check_return_ LPTSTR ReportTagStructToString(_In_ ReportTagStruct* prtReportTag);

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

// Allocates return value with new. Clean up with DeleteConversationIndexStruct.
_Check_return_ ConversationIndexStruct* BinToConversationIndexStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteConversationIndexStruct(_In_ ConversationIndexStruct* pciConversationIndex);
// result allocated with new, clean up with delete[]
_Check_return_ LPTSTR ConversationIndexStructToString(_In_ ConversationIndexStruct* pciConversationIndex);

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

// Allocates return value with new. Clean up with DeleteTaskAssignersStruct.
_Check_return_ TaskAssignersStruct* BinToTaskAssignersStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteTaskAssignersStruct(_In_ TaskAssignersStruct* ptaTaskAssigners);
// result allocated with new, clean up with delete[]
_Check_return_ LPTSTR TaskAssignersStructToString(_In_ TaskAssignersStruct* ptaTaskAssigners);

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

// Allocates return value with new. Clean up with DeleteGlobalObjectIdStruct.
_Check_return_ GlobalObjectIdStruct* BinToGlobalObjectIdStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteGlobalObjectIdStruct(_In_ GlobalObjectIdStruct* pgoidGlobalObjectId);
// result allocated with new, clean up with delete[]
_Check_return_ LPTSTR GlobalObjectIdStructToString(_In_ GlobalObjectIdStruct* pgoidGlobalObjectId);

enum EIDStructType {
	eidtUnknown = 0,
	eidtShortTerm,
	eidtFolder,
	eidtMessage,
	eidtMessageDatabase,
	eidtOneOff,
	eidtAddressBook,
	eidtContact,
	eidtNewsGroupFolder,
	eidtWAB,
};

// EntryIdStruct
// =====================
//   This structure specifies an Entry Id
//
struct EntryIdStruct;
typedef struct EntryIdStruct
{
	BYTE abFlags[4];
	BYTE ProviderUID[16];
	EIDStructType ObjectType; // My own addition to simplify union parsing
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
		struct
		{
			BYTE Type;
			EntryIdStruct* lpEntryID;
		} WAB;
	} ProviderData;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
} EntryIdStruct;

// Allocates return value with new. Clean up with DeleteEntryIdStruct.
_Check_return_ EntryIdStruct* BinToEntryIdStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
_Check_return_ EntryIdStruct* BinToEntryIdStructWithSize(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _Out_opt_ size_t* lpcbBytesRead);
void DeleteEntryIdStruct(_In_ EntryIdStruct* peidEntryId);
// result allocated with new, clean up with delete[]
_Check_return_ LPTSTR EntryIdStructToString(_In_ EntryIdStruct* peidEntryId);

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

// Allocates return value with new. Clean up with DeletePropertyStruct.
_Check_return_ PropertyStruct* BinToPropertyStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, DWORD dwPropCount);
// Allocates return value with new. Clean up with DeletePropertyStruct.
_Check_return_ LPSPropValue BinToSPropValue(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, DWORD dwPropCount, _Out_ size_t* lpcbBytesRead, BOOL bStringPropsExcludeLength);
// Neuters an array of SPropValues - caller must use delete to delete the SPropValue
void DeleteSPropVal(ULONG cVal, _In_count_(cVal) LPSPropValue lpsPropVal);
void DeletePropertyStruct(_In_ PropertyStruct* ppProperty);
// result allocated with new, clean up with delete[]
_Check_return_ LPTSTR PropertyStructToString(_In_ PropertyStruct* ppProperty);

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

// Allocates return value with new. Clean up with DeleteRestrictionStruct.
_Check_return_ RestrictionStruct* BinToRestrictionStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
_Check_return_ RestrictionStruct* BinToRestrictionStructWithSize(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _Out_opt_ size_t* lpcbBytesRead);
// Caller allocates with new. Clean up with DeleteRestriction and delete[].
void BinToRestriction(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _Out_ size_t* lpcbBytesRead, _In_ LPSRestriction psrRestriction, BOOL bRuleCondition, BOOL bExtendedCount);
// Neuters an SRestriction - caller must use delete to delete the SRestriction
void DeleteRestriction(_In_ LPSRestriction lpRes);
void DeleteRestrictionStruct(_In_ RestrictionStruct* prRestriction);
// result allocated with new, clean up with delete[]
_Check_return_ LPTSTR RestrictionStructToString(_In_ RestrictionStruct* prRestriction);

// http://msdn.microsoft.com/en-us/library/ee158295.aspx
// http://msdn.microsoft.com/en-us/library/ee179073.aspx

// [MS-OXCDATA]
// PropertyNameStruct
// =====================
//   This structure specifies a Property Name
//
typedef struct
{
	BYTE Kind;
	GUID Guid;
	DWORD LID;
	BYTE NameSize;
	LPWSTR Name;
} PropertyNameStruct;

// [MS-OXORULE]
// NamedPropertyInformationStruct
// =====================
//   This structure specifies named property information for a rule condition
//
typedef struct
{
	WORD NoOfNamedProps;
	WORD* PropId;
	DWORD NamedPropertiesSize;
	PropertyNameStruct* PropertyName;
} NamedPropertyInformationStruct;

// [MS-OXORULE]
// RuleConditionStruct
// =====================
//   This structure specifies a Rule Condition
//
typedef struct
{
	NamedPropertyInformationStruct NamedPropertyInformation;
	LPSRestriction lpRes;

	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
} RuleConditionStruct;

// Rule Condition - these are used in rules messages
void RuleConditionToString(SBinary myBin, _Deref_out_opt_z_ LPTSTR* lpszResultString, BOOL bExtended);
// Allocates return value with new. Clean up with DeleteRuleConditionStruct.
_Check_return_ RuleConditionStruct* BinToRuleConditionStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _Out_opt_ size_t* lpcbBytesRead, BOOL bExtended);
void DeleteRuleConditionStruct(_In_ RuleConditionStruct* prcRuleCondition);
// result allocated with new, clean up with delete[]
_Check_return_ LPTSTR RuleConditionStructToString(_In_ RuleConditionStruct* prcRuleCondition, BOOL bExtended);

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

// Allocates return value with new. Clean up with DeleteEntryListStruct.
_Check_return_ EntryListStruct* BinToEntryListStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteEntryListStruct(_In_ EntryListStruct* pelEntryList);
// result allocated with new, clean up with delete[]
_Check_return_ LPTSTR EntryListStructToString(_In_ EntryListStruct* pelEntryList);

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

// Allocates return value with new. Clean up with DeleteSearchFolderDefinitionStruct.
_Check_return_ SearchFolderDefinitionStruct* BinToSearchFolderDefinitionStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteSearchFolderDefinitionStruct(_In_ SearchFolderDefinitionStruct* psfdSearchFolderDefinition);
// result allocated with new, clean up with delete[]
_Check_return_ LPTSTR SearchFolderDefinitionStructToString(_In_ SearchFolderDefinitionStruct* psfdSearchFolderDefinition);

// PackedUnicodeString
// =====================
//   This structure specifies a Packed Unicode String
//
typedef struct
{
	BYTE cchLength;
	WORD cchExtendedLength;
	LPWSTR szCharacters;
} PackedUnicodeString;

// PackedAnsiString
// =====================
//   This structure specifies a Packed Ansi String
//
typedef struct
{
	BYTE cchLength;
	WORD cchExtendedLength;
	LPSTR szCharacters;
} PackedAnsiString;

// SkipBlock
// =====================
//   This structure specifies a Skip Block
//
typedef struct
{
	DWORD dwSize;
	BYTE* lpbContent;
} SkipBlock;

// FieldDefinition
// =====================
//   This structure specifies a Field Definition
//
typedef struct
{
	DWORD dwFlags;
	WORD wVT;
	DWORD dwDispid;
	WORD wNmidNameLength;
	LPWSTR szNmidName;
	PackedAnsiString pasNameANSI;
	PackedAnsiString pasFormulaANSI;
	PackedAnsiString pasValidationRuleANSI;
	PackedAnsiString pasValidationTextANSI;
	PackedAnsiString pasErrorANSI;
	DWORD dwInternalType;
	DWORD dwSkipBlockCount;
	SkipBlock* psbSkipBlocks;
} FieldDefinition;

// PropertyDefinitionStreamStruct
// =====================
//   This structure specifies a Property Definition Stream
//
typedef struct
{
	WORD wVersion;
	DWORD dwFieldDefinitionCount;
	FieldDefinition* pfdFieldDefinitions;

	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
} PropertyDefinitionStreamStruct;

// Allocates return value with new. Clean up with DeletePropertyDefinitionStreamStruct.
_Check_return_ PropertyDefinitionStreamStruct* BinToPropertyDefinitionStreamStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeletePropertyDefinitionStreamStruct(_In_ PropertyDefinitionStreamStruct* ppdsPropertyDefinitionStream);
// result allocated with new, clean up with delete[]
_Check_return_ LPTSTR PropertyDefinitionStreamStructToString(_In_ PropertyDefinitionStreamStruct* ppdsPropertyDefinitionStream);

// PersistElement
// =====================
//   This structure specifies a Persist Element block
//
typedef struct
{
	WORD wElementID;
	WORD wElementDataSize;
	LPBYTE lpbElementData;
} PersistElement;

// PersistData
// =====================
//   This structure specifies a Persist Data block
//
typedef struct
{
	WORD wPersistID;
	WORD wDataElementsSize;
	WORD wDataElementCount;
	PersistElement* ppeDataElement;

	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
} PersistData;

// AdditionalRenEntryIDsStruct
// =====================
//   This structure specifies a Additional Ren Entry ID blob
//
typedef struct
{
	WORD wPersistDataCount;
	PersistData* ppdPersistData;

	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
} AdditionalRenEntryIDsStruct;

// Allocates return value with new. Clean up with DeleteAdditionalRenEntryIDsStruct.
void BinToPersistData(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _Out_ size_t* lpcbBytesRead, _Out_ PersistData* ppdPersistData);
// Allocates return value with new. Clean up with DeleteAdditionalRenEntryIDsStruct.
_Check_return_ AdditionalRenEntryIDsStruct* BinToAdditionalRenEntryIDsStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteAdditionalRenEntryIDsStruct(_In_ AdditionalRenEntryIDsStruct* pareiAdditionalRenEntryIDs);
// result allocated with new, clean up with delete[]
_Check_return_ LPTSTR AdditionalRenEntryIDsStructToString(_In_ AdditionalRenEntryIDsStruct* pareiAdditionalRenEntryIDs);

// FlatEntryIDStruct
// =====================
//   This structure specifies a Flat Entry ID in a Flat Entry List blob
//
typedef struct
{
	DWORD dwSize;
	EntryIdStruct* lpEntryID;

	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
} FlatEntryIDStruct;

// FlatEntryListStruct
// =====================
//   This structure specifies a Flat Entry List blob
//
typedef struct
{
	DWORD cEntries;
	DWORD cbEntries;
	FlatEntryIDStruct* pEntryIDs;

	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
} FlatEntryListStruct;

// Allocates return value with new. Clean up with DeleteFlatEntryListStruct.
_Check_return_ FlatEntryListStruct* BinToFlatEntryListStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteFlatEntryListStruct(_In_ FlatEntryListStruct* pfelFlatEntryList);
// result allocated with new, clean up with delete[]
_Check_return_ LPTSTR FlatEntryListStructToString(_In_ FlatEntryListStruct* pfelFlatEntryList);

// WebViewPersistStruct
// =====================
//   This structure specifies a single Web View Persistance Object
//
typedef struct
{
	DWORD dwVersion;
	DWORD dwType;
	DWORD dwFlags;
	DWORD dwUnused[7];
	DWORD cbData;
	LPBYTE lpData;
} WebViewPersistStruct;

// WebViewPersistStreamStruct
// =====================
//   This structure specifies a Web View Persistance Object stream struct
//
typedef struct
{
	DWORD cWebViews;
	WebViewPersistStruct* lpWebViews;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
} WebViewPersistStreamStruct;

// Allocates return value with new. Clean up with DeleteWebViewPersistStreamStruct.
_Check_return_ WebViewPersistStreamStruct* BinToWebViewPersistStreamStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteWebViewPersistStreamStruct(_In_ WebViewPersistStreamStruct* pwvpsWebViewPersistStream);
// result allocated with new, clean up with delete[]
_Check_return_ LPTSTR WebViewPersistStreamStructToString(_In_ WebViewPersistStreamStruct* pwvpsWebViewPersistStream);

// RecipientRowStreamStruct
// =====================
//   This structure specifies an recipient row stream struct
//
typedef struct
{
	DWORD cVersion;
	DWORD cRowCount;
	LPADRENTRY lpAdrEntry;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
} RecipientRowStreamStruct;

// Allocates return value with new. Clean up with DeleteRecipientRowStreamStruct.
_Check_return_ RecipientRowStreamStruct* BinToRecipientRowStreamStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteRecipientRowStreamStruct(_In_ RecipientRowStreamStruct* prrsRecipientRowStream);
// result allocated with new, clean up with delete[]
_Check_return_ LPTSTR RecipientRowStreamStructToString(_In_ RecipientRowStreamStruct* prrsRecipientRowStream);

// FolderFieldDefinitionCommon
// =====================
//   This structure specifies a folder field definition common struct
//
typedef struct
{
	GUID PropSetGuid;
	DWORD fcapm;
	DWORD dwString;
	DWORD dwBitmap;
	DWORD dwDisplay;
	DWORD iFmt;
	WORD wszFormulaLength;
	LPWSTR wszFormula;
} FolderFieldDefinitionCommon;

// FolderFieldDefinitionA
// =====================
//   This structure specifies a folder field definition ANSI struct
//
typedef struct
{
	DWORD FieldType;
	WORD FieldNameLength;
	LPSTR FieldName;
	FolderFieldDefinitionCommon Common;
} FolderFieldDefinitionA;

// FolderFieldDefinitionW
// =====================
//   This structure specifies a folder field definition Unicode struct
//
typedef struct
{
	DWORD FieldType;
	WORD FieldNameLength;
	LPWSTR FieldName;
	FolderFieldDefinitionCommon Common;
} FolderFieldDefinitionW;

// FolderUserFieldA
// =====================
//   This structure specifies a folder user field ANSI stream struct
//
typedef struct
{
	DWORD FieldDefinitionCount;
	FolderFieldDefinitionA* FieldDefinitions;
} FolderUserFieldA;

// FolderUserFieldW
// =====================
//   This structure specifies a folder user field Unicode stream struct
//
typedef struct
{
	DWORD FieldDefinitionCount;
	FolderFieldDefinitionW* FieldDefinitions;
} FolderUserFieldW;

// FolderUserFieldStreamStruct
// =====================
//   This structure specifies a folder user field stream struct
//
typedef struct
{
	FolderUserFieldA FolderUserFieldsAnsi;
	FolderUserFieldW FolderUserFieldsUnicode;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
} FolderUserFieldStreamStruct;

// Allocates return value with new. Clean up with DeleteFolderUserFieldStreamStruct.
_Check_return_ FolderUserFieldStreamStruct* BinToFolderUserFieldStreamStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteFolderUserFieldStreamStruct(_In_ FolderUserFieldStreamStruct* pfufsFolderUserFieldStream);
// result allocated with new, clean up with delete[]
_Check_return_ LPTSTR FolderUserFieldStreamStructToString(_In_ FolderUserFieldStreamStruct* pfufsFolderUserFieldStream);
// End Functions to parse PT_BINARY properties