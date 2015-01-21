#pragma once

extern NAME_ARRAY_ENTRY g_uidParsingTypes[];
extern ULONG g_cuidParsingTypes;

// lpszSmartView allocated with new, delete with delete[]
ULONG InterpretPropSmartView(_In_ LPSPropValue lpProp, // required property value
							 _In_opt_ LPMAPIPROP lpMAPIProp, // optional source object
							 _In_opt_ LPMAPINAMEID lpNameID, // optional named property information to avoid GetNamesFromIDs call
							 _In_opt_ LPSBinary lpMappingSignature, // optional mapping signature for object to speed named prop lookups
							 bool bIsAB, // true if we know we're dealing with an address book property (they can be > 8000 and not named props)
							 bool bMVRow, // did the row come from a MV prop?
							 _Deref_out_opt_z_ LPWSTR* lpszSmartView); // Built from lpProp & lpMAPIProp

void InterpretBinaryAsString(SBinary myBin, DWORD_PTR iStructType, _In_opt_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_z_ LPWSTR* lpszResultString);
void InterpretMVBinaryAsString(SBinaryArray myBinArray, DWORD_PTR iStructType, _In_opt_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag, _Deref_out_opt_z_ LPWSTR* lpszResultString);
ULONG InterpretNumberAsString(_PV pV, ULONG ulPropTag, ULONG ulPropNameID, _In_opt_z_ LPWSTR lpszPropNameString, _In_opt_ LPCGUID lpguidNamedProp, bool bLabel, _Deref_out_opt_z_ LPWSTR* lpszResultString);
ULONG InterpretNumberAsStringProp(ULONG ulVal, ULONG ulPropTag, _Deref_out_opt_z_ LPWSTR* lpszResultString);
ULONG InterpretNumberAsStringNamedProp(ULONG ulVal, ULONG ulPropNameID, _In_opt_ LPCGUID lpguidNamedProp, _Deref_out_opt_z_ LPWSTR* lpszResultString);
_Check_return_ LPWSTR FidMidToSzString(LONGLONG llID, bool bLabel);

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

enum EIDStructType
{
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

struct MDB_STORE_EID_V2
{
	ULONG ulMagic; // MDB_STORE_EID_V2_MAGIC
	ULONG ulSize; // size of this struct plus the size of szServerDN and wszServerFQDN
	ULONG ulVersion; // MDB_STORE_EID_V2_VERSION
	ULONG ulOffsetDN; // offset past the beginning of the MDB_STORE_EID_V2 struct where szServerDN starts
	ULONG ulOffsetFQDN; // offset past the beginning of the MDB_STORE_EID_V2 struct where wszServerFQDN starts
};

// EntryIdStruct
// =====================
//   This structure specifies an Entry Id
//
struct EntryIdStruct
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
			bool bIsExchange;
			ULONG WrappedFlags;
			BYTE WrappedProviderUID[16];
			ULONG WrappedType;
			LPSTR ServerShortname;
			LPSTR MailboxDN;
			BOOL bV2;
			MDB_STORE_EID_V2 v2;
			LPSTR v2DN;
			LPWSTR v2FQDN;
			BYTE v2Reserved[2];
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
			DWORD Index; // CONTAB_USER, CONTAB_DISTLIST only
			DWORD EntryIDCount; // CONTAB_USER, CONTAB_DISTLIST only
			BYTE muidID[16]; // CONTAB_CONTAINER only
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
};

// Allocates return value with new. Clean up with DeleteEntryIdStruct.
_Check_return_ EntryIdStruct* BinToEntryIdStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
_Check_return_ EntryIdStruct* BinToEntryIdStructWithSize(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _Out_opt_ size_t* lpcbBytesRead);
void DeleteEntryIdStruct(_In_ EntryIdStruct* peidEntryId);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR EntryIdStructToString(_In_ EntryIdStruct* peidEntryId);

// PropertyStruct
// =====================
//   This structure specifies an array of Properties
//
struct PropertyStruct
{
	DWORD PropCount;
	LPSPropValue Prop;

	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeletePropertyStruct.
_Check_return_ PropertyStruct* BinToPropertyStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
// Allocates return value with new. Clean up with DeletePropertyStruct.
_Check_return_ LPSPropValue BinToSPropValue(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, DWORD dwPropCount, _Out_ size_t* lpcbBytesRead, bool bStringPropsExcludeLength);
// Neuters an array of SPropValues - caller must use delete to delete the SPropValue
void DeleteSPropVal(ULONG cVal, _In_count_(cVal) LPSPropValue lpsPropVal);
void DeletePropertyStruct(_In_ PropertyStruct* ppProperty);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR PropertyStructToString(_In_ PropertyStruct* ppProperty);

// RestrictionStruct
// =====================
//   This structure specifies a Restriction
//
struct RestrictionStruct
{
	LPSRestriction lpRes;

	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeleteRestrictionStruct.
_Check_return_ RestrictionStruct* BinToRestrictionStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
_Check_return_ RestrictionStruct* BinToRestrictionStructWithSize(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _Out_opt_ size_t* lpcbBytesRead);
// Caller allocates with new. Clean up with DeleteRestriction and delete[].
bool BinToRestriction(ULONG ulDepth, ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _Out_ size_t* lpcbBytesRead, _In_ LPSRestriction psrRestriction, bool bRuleCondition, bool bExtendedCount);
// Neuters an SRestriction - caller must use delete to delete the SRestriction
void DeleteRestriction(_In_ LPSRestriction lpRes);
void DeleteRestrictionStruct(_In_ RestrictionStruct* prRestriction);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR RestrictionStructToString(_In_ RestrictionStruct* prRestriction);

// http://msdn.microsoft.com/en-us/library/ee158295.aspx
// http://msdn.microsoft.com/en-us/library/ee179073.aspx

// [MS-OXCDATA]
// PropertyNameStruct
// =====================
//   This structure specifies a Property Name
//
struct PropertyNameStruct
{
	BYTE Kind;
	GUID Guid;
	DWORD LID;
	BYTE NameSize;
	LPWSTR Name;
};

// [MS-OXORULE]
// NamedPropertyInformationStruct
// =====================
//   This structure specifies named property information for a rule condition
//
struct NamedPropertyInformationStruct
{
	WORD NoOfNamedProps;
	WORD* PropId;
	DWORD NamedPropertiesSize;
	PropertyNameStruct* PropertyName;
};

// [MS-OXORULE]
// RuleConditionStruct
// =====================
//   This structure specifies a Rule Condition
//
struct RuleConditionStruct
{
	NamedPropertyInformationStruct NamedPropertyInformation;
	LPSRestriction lpRes;

	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Rule Condition - these are used in rules messages
void RuleConditionToString(SBinary myBin, _Deref_out_opt_z_ LPWSTR* lpszResultString, bool bExtended);
// Allocates return value with new. Clean up with DeleteRuleConditionStruct.
_Check_return_ RuleConditionStruct* BinToRuleConditionStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _Out_opt_ size_t* lpcbBytesRead, bool bExtended);
void DeleteRuleConditionStruct(_In_ RuleConditionStruct* prcRuleCondition);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR RuleConditionStructToString(_In_ RuleConditionStruct* prcRuleCondition, bool bExtended);

// EntryListEntryStruct
// =====================
//   This structure specifies an Entry in an Entry List
//
struct EntryListEntryStruct
{
	DWORD EntryLength;
	DWORD EntryLengthPad;
	EntryIdStruct* EntryId;
};

// EntryListStruct
// =====================
//   This structure specifies an Entry List
//
struct EntryListStruct
{
	DWORD EntryCount;
	DWORD Pad;

	EntryListEntryStruct* Entry;

	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeleteEntryListStruct.
_Check_return_ EntryListStruct* BinToEntryListStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteEntryListStruct(_In_ EntryListStruct* pelEntryList);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR EntryListStructToString(_In_ EntryListStruct* pelEntryList);

// AddressListEntryStruct
// =====================
//   This structure specifies an entry in an Address List
//
struct AddressListEntryStruct
{
	DWORD PropertyCount;
	DWORD Pad;
	PropertyStruct Properties;
};

// SearchFolderDefinitionStruct
// =====================
//   This structure specifies a Search Folder Definition
//
struct SearchFolderDefinitionStruct
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
};

// Allocates return value with new. Clean up with DeleteSearchFolderDefinitionStruct.
_Check_return_ SearchFolderDefinitionStruct* BinToSearchFolderDefinitionStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteSearchFolderDefinitionStruct(_In_ SearchFolderDefinitionStruct* psfdSearchFolderDefinition);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR SearchFolderDefinitionStructToString(_In_ SearchFolderDefinitionStruct* psfdSearchFolderDefinition);

// PackedUnicodeString
// =====================
//   This structure specifies a Packed Unicode String
//
struct PackedUnicodeString
{
	BYTE cchLength;
	WORD cchExtendedLength;
	LPWSTR szCharacters;
};

// PackedAnsiString
// =====================
//   This structure specifies a Packed Ansi String
//
struct PackedAnsiString
{
	BYTE cchLength;
	WORD cchExtendedLength;
	LPSTR szCharacters;
};

// SkipBlock
// =====================
//   This structure specifies a Skip Block
//
struct SkipBlock
{
	DWORD dwSize;
	BYTE* lpbContent;
};

// FieldDefinition
// =====================
//   This structure specifies a Field Definition
//
struct FieldDefinition
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
};

// PropertyDefinitionStreamStruct
// =====================
//   This structure specifies a Property Definition Stream
//
struct PropertyDefinitionStreamStruct
{
	WORD wVersion;
	DWORD dwFieldDefinitionCount;
	FieldDefinition* pfdFieldDefinitions;

	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeletePropertyDefinitionStreamStruct.
_Check_return_ PropertyDefinitionStreamStruct* BinToPropertyDefinitionStreamStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeletePropertyDefinitionStreamStruct(_In_ PropertyDefinitionStreamStruct* ppdsPropertyDefinitionStream);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR PropertyDefinitionStreamStructToString(_In_ PropertyDefinitionStreamStruct* ppdsPropertyDefinitionStream);

// PersistElement
// =====================
//   This structure specifies a Persist Element block
//
struct PersistElement
{
	WORD wElementID;
	WORD wElementDataSize;
	LPBYTE lpbElementData;
};

// PersistData
// =====================
//   This structure specifies a Persist Data block
//
struct PersistData
{
	WORD wPersistID;
	WORD wDataElementsSize;
	WORD wDataElementCount;
	PersistElement* ppeDataElement;

	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// AdditionalRenEntryIDsStruct
// =====================
//   This structure specifies a Additional Ren Entry ID blob
//
struct AdditionalRenEntryIDsStruct
{
	WORD wPersistDataCount;
	PersistData* ppdPersistData;

	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeleteAdditionalRenEntryIDsStruct.
void BinToPersistData(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin, _Out_ size_t* lpcbBytesRead, _Out_ PersistData* ppdPersistData);
// Allocates return value with new. Clean up with DeleteAdditionalRenEntryIDsStruct.
_Check_return_ AdditionalRenEntryIDsStruct* BinToAdditionalRenEntryIDsStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteAdditionalRenEntryIDsStruct(_In_ AdditionalRenEntryIDsStruct* pareiAdditionalRenEntryIDs);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR AdditionalRenEntryIDsStructToString(_In_ AdditionalRenEntryIDsStruct* pareiAdditionalRenEntryIDs);

// FlatEntryIDStruct
// =====================
//   This structure specifies a Flat Entry ID in a Flat Entry List blob
//
struct FlatEntryIDStruct
{
	DWORD dwSize;
	EntryIdStruct* lpEntryID;

	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// FlatEntryListStruct
// =====================
//   This structure specifies a Flat Entry List blob
//
struct FlatEntryListStruct
{
	DWORD cEntries;
	DWORD cbEntries;
	FlatEntryIDStruct* pEntryIDs;

	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeleteFlatEntryListStruct.
_Check_return_ FlatEntryListStruct* BinToFlatEntryListStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteFlatEntryListStruct(_In_ FlatEntryListStruct* pfelFlatEntryList);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR FlatEntryListStructToString(_In_ FlatEntryListStruct* pfelFlatEntryList);

// WebViewPersistStruct
// =====================
//   This structure specifies a single Web View Persistance Object
//
struct WebViewPersistStruct
{
	DWORD dwVersion;
	DWORD dwType;
	DWORD dwFlags;
	DWORD dwUnused[7];
	DWORD cbData;
	LPBYTE lpData;
};

// WebViewPersistStreamStruct
// =====================
//   This structure specifies a Web View Persistance Object stream struct
//
struct WebViewPersistStreamStruct
{
	DWORD cWebViews;
	WebViewPersistStruct* lpWebViews;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeleteWebViewPersistStreamStruct.
_Check_return_ WebViewPersistStreamStruct* BinToWebViewPersistStreamStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteWebViewPersistStreamStruct(_In_ WebViewPersistStreamStruct* pwvpsWebViewPersistStream);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR WebViewPersistStreamStructToString(_In_ WebViewPersistStreamStruct* pwvpsWebViewPersistStream);

// RecipientRowStreamStruct
// =====================
//   This structure specifies an recipient row stream struct
//
struct RecipientRowStreamStruct
{
	DWORD cVersion;
	DWORD cRowCount;
	LPADRENTRY lpAdrEntry;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeleteRecipientRowStreamStruct.
_Check_return_ RecipientRowStreamStruct* BinToRecipientRowStreamStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteRecipientRowStreamStruct(_In_ RecipientRowStreamStruct* prrsRecipientRowStream);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR RecipientRowStreamStructToString(_In_ RecipientRowStreamStruct* prrsRecipientRowStream);

// FolderFieldDefinitionCommon
// =====================
//   This structure specifies a folder field definition common struct
//
struct FolderFieldDefinitionCommon
{
	GUID PropSetGuid;
	DWORD fcapm;
	DWORD dwString;
	DWORD dwBitmap;
	DWORD dwDisplay;
	DWORD iFmt;
	WORD wszFormulaLength;
	LPWSTR wszFormula;
};

// FolderFieldDefinitionA
// =====================
//   This structure specifies a folder field definition ANSI struct
//
struct FolderFieldDefinitionA
{
	DWORD FieldType;
	WORD FieldNameLength;
	LPSTR FieldName;
	FolderFieldDefinitionCommon Common;
};

// FolderFieldDefinitionW
// =====================
//   This structure specifies a folder field definition Unicode struct
//
struct FolderFieldDefinitionW
{
	DWORD FieldType;
	WORD FieldNameLength;
	LPWSTR FieldName;
	FolderFieldDefinitionCommon Common;
};

// FolderUserFieldA
// =====================
//   This structure specifies a folder user field ANSI stream struct
//
struct FolderUserFieldA
{
	DWORD FieldDefinitionCount;
	FolderFieldDefinitionA* FieldDefinitions;
};

// FolderUserFieldW
// =====================
//   This structure specifies a folder user field Unicode stream struct
//
struct FolderUserFieldW
{
	DWORD FieldDefinitionCount;
	FolderFieldDefinitionW* FieldDefinitions;
};

// FolderUserFieldStreamStruct
// =====================
//   This structure specifies a folder user field stream struct
//
struct FolderUserFieldStreamStruct
{
	FolderUserFieldA FolderUserFieldsAnsi;
	FolderUserFieldW FolderUserFieldsUnicode;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeleteFolderUserFieldStreamStruct.
_Check_return_ FolderUserFieldStreamStruct* BinToFolderUserFieldStreamStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteFolderUserFieldStreamStruct(_In_ FolderUserFieldStreamStruct* pfufsFolderUserFieldStream);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR FolderUserFieldStreamStructToString(_In_ FolderUserFieldStreamStruct* pfufsFolderUserFieldStream);

// NickNameCacheStruct
// =====================
//   This structure specifies a nickname cache struct
//
struct NickNameCacheStruct
{
	BYTE Metadata1[4];
	ULONG ulMajorVersion;
	ULONG ulMinorVersion;
	DWORD cRowCount;
	LPSRow lpRows;
	ULONG cbEI;
	LPBYTE lpbEI;
	BYTE Metadata2[8];
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeleteNickNameCacheStruct.
_Check_return_ NickNameCacheStruct* BinToNickNameCacheStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteNickNameCacheStruct(_In_ NickNameCacheStruct* pnncNickNameCache);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR NickNameCacheStructToString(_In_ NickNameCacheStruct* pnncNickNameCache);

// VerbDataStruct
// =====================
//   This structure specifies verb data
//
struct VerbDataStruct
{
	DWORD VerbType;
	BYTE DisplayNameCount;
	LPSTR DisplayName;
	BYTE MsgClsNameCount;
	LPSTR MsgClsName;
	BYTE Internal1StringCount;
	LPSTR Internal1String;
	BYTE DisplayNameCountRepeat;
	LPSTR DisplayNameRepeat;
	DWORD Internal2;
	BYTE Internal3;
	DWORD fUseUSHeaders;
	DWORD Internal4;
	DWORD SendBehavior;
	DWORD Internal5;
	DWORD ID;
	DWORD Internal6;
};

// VerbExtraDataStruct
// =====================
//   This structure specifies extra verb data
//
struct VerbExtraDataStruct
{
	BYTE DisplayNameCount;
	LPWSTR DisplayName;
	BYTE DisplayNameCountRepeat;
	LPWSTR DisplayNameRepeat;
};

// VerbStreamStruct
// =====================
//   This structure specifies a verb stream
//
struct VerbStreamStruct
{
	WORD Version;
	DWORD Count;
	VerbDataStruct* lpVerbData;
	WORD Version2;
	VerbExtraDataStruct* lpVerbExtraData;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeleteVerbStreamStruct.
_Check_return_ VerbStreamStruct* BinToVerbStreamStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteVerbStreamStruct(_In_ VerbStreamStruct* pvsVerbStream);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR VerbStreamStructToString(_In_ VerbStreamStruct* pvsVerbStream);

// TombstoneRecord
// =====================
//   This structure specifies a tombstone record
//
struct TombstoneRecord
{
	DWORD StartTime;
	DWORD EndTime;
	DWORD GlobalObjectIdSize;
	LPBYTE lpGlobalObjectId;
	WORD UsernameSize;
	LPSTR szUsername;
};


// TombstoneStruct
// =====================
//   This structure specifies a tombstone stream
//
struct TombstoneStruct
{
	DWORD Identifier;
	DWORD HeaderSize;
	DWORD Version;
	DWORD RecordsCount;
	DWORD ActualRecordsCount; // computed based on state, not read value
	DWORD RecordsSize;
	TombstoneRecord* lpRecords;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeleteTombstoneStruct.
_Check_return_ TombstoneStruct* BinToTombstoneStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeleteTombstoneStruct(_In_ TombstoneStruct* ptsTombstone);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR TombstoneStructToString(_In_ TombstoneStruct* ptsTombstone);

// SizedXID
// =====================
//   This structure specifies a sized XID record
//
struct SizedXID
{
	BYTE XidSize;
	GUID NamespaceGuid;
	DWORD cbLocalId;
	LPBYTE LocalID;
};

// PCLStruct
// =====================
//   This structure specifies a Predecessor Change List stream
//
struct PCLStruct
{
	DWORD cXID;
	SizedXID* lpXID;
	size_t JunkDataSize;
	LPBYTE JunkData; // My own addition to account for unparsed data in persisted property
};

// Allocates return value with new. Clean up with DeletePCLStruct.
_Check_return_ PCLStruct* BinToPCLStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin);
void DeletePCLStruct(_In_ PCLStruct* pcl);
// result allocated with new, clean up with delete[]
_Check_return_ LPWSTR PCLStructToString(_In_ PCLStruct* pcl);
// End Functions to parse PT_BINARY properties