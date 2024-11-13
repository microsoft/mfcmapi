#pragma once

enum __NonPropFlag
{
	flagSearchFlag = 0x10000, // ensure that all flags in the enum are > 0xffff
	flagSearchState,
	flagTableStatus,
	flagTableType,
	flagObjectType,
	flagSecurityVersion,
	flagSecurityInfo,
	flagACEFlag,
	flagACEType,
	flagACEMaskContainer,
	flagACEMaskNonContainer,
	flagACEMaskFreeBusy,
	flagStreamFlag,
	flagRestrictionType,
	flagBitmask,
	flagRelop,
	flagActionType,
	flagBounceCode,
	flagOPReply,
	flagOpForward,
	flagFuzzyLevel,
	flagRulesVersion,
	flagNotifEventType,
	flagTableEventType,
	flagTZRule,
	flagRuleFlag,
	flagExtendedFolderFlagType,
	flagExtendedFolderFlag,
	flagRecurFrequency,
	flagPatternType,
	flagCalendarType,
	flagDOW,
	flagN,
	flagEndType,
	flagFirstDOW,
	flagOverrideFlags,
	flagReportTagVersion,
	flagGlobalObjectIdMonth,
	flagOneOffEntryId,
	flagEntryId0,
	flagEntryId1,
	flagMessageDatabaseObjectType,
	flagContabVersion,
	flagContabType,
	flagContabIndex,
	flagExchangeABVersion,
	flagMDBVersion,
	flagMDBFlag,
	flagPropDefVersion,
	flagPDOFlag,
	flagVarEnum,
	flagInternalType,
	flagPersistID,
	flagElementID,
	flagWABEntryIDType,
	flagWebViewVersion,
	flagWebViewType,
	flagWebViewFlags,
	flagFolderType,
	flagFieldCap,
	flagCcsf,
	flagIet,
	flagEidMagic,
	flagEidVersion,
	flagToDoSwapFlag,
	flagCapabilitiesFolder,
	flagCapabilitiesRestriction,
};

#define PR_FREEBUSY_NT_SECURITY_DESCRIPTOR (PROP_TAG(PT_BINARY, 0x0F00))

// http://support.microsoft.com/kb/171670
// Entry ID for the Journal Folder
#define PR_IPM_JOURNAL_ENTRYID (PROP_TAG(PT_BINARY, 0x36D2))
// Entry ID for the Notes Folder
#define PR_IPM_NOTE_ENTRYID (PROP_TAG(PT_BINARY, 0x36D3))
// Entry IDs for the Reminders Folder
#define PR_REM_ONLINE_ENTRYID (PROP_TAG(PT_BINARY, 0x36D5))
#define PR_REM_OFFLINE_ENTRYID PROP_TAG(PT_BINARY, 0x36D6)

#define PR_FREEBUSY_ENTRYIDS PROP_TAG(PT_MV_BINARY, 0x36E4)
#ifndef PR_RECIPIENT_TRACKSTATUS
#define PR_RECIPIENT_TRACKSTATUS PROP_TAG(PT_LONG, 0x5FFF)
#endif
#define PR_RECIPIENT_FLAGS PROP_TAG(PT_LONG, 0x5FFD)
#define PR_RECIPIENT_ENTRYID PROP_TAG(PT_BINARY, 0x5FF7)

#ifndef PR_NT_SECURITY_DESCRIPTOR
#define PR_NT_SECURITY_DESCRIPTOR (PROP_TAG(PT_BINARY, 0x0E27))
#endif
#ifndef PR_BODY_HTML
#define PR_BODY_HTML (PROP_TAG(PT_TSTRING, 0x1013))
#endif

#ifndef PR_SEND_INTERNET_ENCODING
#define PR_SEND_INTERNET_ENCODING PROP_TAG(PT_LONG, 0x3A71)
#endif
#ifndef PR_ATTACH_FLAGS
#define PR_ATTACH_FLAGS PROP_TAG(PT_LONG, 0x3714)
#endif

// http://support.microsoft.com/kb/837364
#ifndef PR_CONFLICT_ITEMS
#define PR_CONFLICT_ITEMS PROP_TAG(PT_MV_BINARY, 0x1098)
#endif

// http://support.microsoft.com/kb/278321
#ifndef PR_INETMAIL_OVERRIDE_FORMAT
#define PR_INETMAIL_OVERRIDE_FORMAT PROP_TAG(PT_LONG, 0x5902)
#endif

// http://support.microsoft.com/kb/312900
#define PR_CERT_DEFAULTS PROP_TAG(PT_LONG, 0x0020)
// Values for PR_CERT_DEFAULTS
#define MSG_DEFAULTS_NONE 0
#define MSG_DEFAULTS_FOR_FORMAT 1 // Default certificate for S/MIME.
#define MSG_DEFAULTS_GLOBAL 2 // Default certificate for all formats.
#define MSG_DEFAULTS_SEND_CERT 4 // Send certificate with message.

// http://support.microsoft.com/kb/194955
#define PR_AGING_GRANULARITY PROP_TAG(PT_LONG, 0x36EE)

// http://msdn2.microsoft.com/en-us/library/bb176434.aspx
#define PR_AGING_DEFAULT PROP_TAG(PT_LONG, 0x685E)

#define AG_DEFAULT_FILE 0x01
#define AG_DEFAULT_ALL 0x02

// IMPORTANT NOTE: This property holds additional Ren special folder EntryIDs.
// The EntryID is for the special folder is located at sf* - sfRenMVEntryIDs
// This is the only place you should add new special folder ids, and the order
// of these entry ids must be preserved for legacy clients.
// Also, all new (as of Office.NET) special folders should have the extended
// folder flag XEFF_SPECIAL_FOLDER to tell the folder tree data to check if
// this folder is actually special or not.
// See comment above for places in Outlook that will need modification if you
// add a new sf* index.
// It currently contains:
// sfConflicts 0
// sfSyncFailures 1
// sfLocalFailures 2
// sfServerFailures 3
// sfJunkEmail 4
// sfSpamTagDontUse 5
//
// NOTE: sfSpamTagDontUse is not the real special folder but used #5 slot
// Therefore, we need to skip it when enum through sf* special folders.
#define PR_ADDITIONAL_REN_ENTRYIDS PROP_TAG(PT_MV_BINARY, 0x36D8)

// [MS-NSPI].pdf
#define DT_CONTAINER ((ULONG) 0x00000100)
#define DT_TEMPLATE ((ULONG) 0x00000101)
#define DT_ADDRESS_TEMPLATE ((ULONG) 0x00000102)
#define DT_SEARCH ((ULONG) 0x00000200)

// http://msdn2.microsoft.com/en-us/library/bb821036.aspx
// #define PR_FLAG_STATUS PROP_TAG(PT_LONG, 0x1090)
enum FollowUpStatus
{
	flwupNone = 0,
	flwupComplete,
	flwupMarked
};

// http://msdn2.microsoft.com/en-us/library/bb821062.aspx
#define PR_FOLLOWUP_ICON PROP_TAG(PT_LONG, 0x1095)
enum OlFlagIcon
{
	olNoFlagIcon = 0,
	olPurpleFlagIcon = 1,
	olOrangeFlagIcon = 2,
	olGreenFlagIcon = 3,
	olYellowFlagIcon = 4,
	olBlueFlagIcon = 5,
	olRedFlagIcon = 6,
};

// http://msdn2.microsoft.com/en-us/library/bb821130.aspx
enum Gender
{
	genderMin = 0,
	genderUnspecified = genderMin,
	genderFemale,
	genderMale,
	genderCount,
	genderMax = genderCount - 1
};

// [MS-OXCFOLD]
// Use CI searches exclusively (never use old-style search)
#define CONTENT_INDEXED_SEARCH ((ULONG) 0x00010000)

// Never use CI search (old-style search only)
#define NON_CONTENT_INDEXED_SEARCH ((ULONG) 0x00020000)

// Make the search static (no backlinks/dynamic updates). This is independent
// of whether or not the search uses CI.
#define STATIC_SEARCH ((ULONG) 0x00040000)

// http://msdn.microsoft.com/en-us/library/ee219969(EXCHG.80).aspx
// The search used the content index (CI) in some fashion, and is
// non-dynamic
// NOTE: If SEARCH_REBUILD is set, the query is still being processed, and
// the static-ness may not yet have been determined (see SEARCH_MAYBE_STATIC).
#define SEARCH_STATIC ((ULONG) 0x00010000)

// The search is still being evaluated (SEARCH_REBUILD should always
// be set when this bit is set), and could become either static or dynamic.
// This bit is needed to distinguish this in-progress state separately from
// static-only (SEARCH_STATIC) or dynamic-only (default).
#define SEARCH_MAYBE_STATIC ((ULONG) 0x00020000)

// CI TWIR Search State for a query request.
// Currently, we have 4 different search states for CI/TWIR, they are:
// CI Totally, CI with TWIR residual, TWIR Mostly, and TWIR Totally
#define SEARCH_COMPLETE ((ULONG) 0x00001000)
#define SEARCH_PARTIAL ((ULONG) 0x00002000)
#define CI_TOTALLY ((ULONG) 0x01000000)
#define CI_WITH_TWIR_RESIDUAL ((ULONG) 0x02000000)
#define TWIR_MOSTLY ((ULONG) 0x04000000)
#define TWIR_TOTALLY ((ULONG) 0x08000000)

// [MS-OXCSPAM].pdf
#define PR_SENDER_ID_STATUS PROP_TAG(PT_LONG, 0x4079)
// Values that PR_SENDER_ID_STATUS can take
#define SENDER_ID_NEUTRAL 0x1
#define SENDER_ID_PASS 0x2
#define SENDER_ID_FAIL 0x3
#define SENDER_ID_SOFT_FAIL 0x4
#define SENDER_ID_NONE 0x5
#define SENDER_ID_TEMP_ERROR 0x80000006
#define SENDER_ID_PERM_ERROR 0x80000007

#define PR_JUNK_THRESHOLD PROP_TAG(PT_LONG, 0x6101)
#define SPAM_FILTERING_NONE 0xFFFFFFFF
#define SPAM_FILTERING_LOW 0x00000006
#define SPAM_FILTERING_MEDIUM 0x00000005
#define SPAM_FILTERING_HIGH 0x00000003
#define SPAM_FILTERING_TRUSTED_ONLY 0x80000000

// [MS-OXOCAL].pdf
#define RECIP_UNSENDABLE (int) 0x0000
#define RECIP_SENDABLE (int) 0x0001
#define RECIP_ORGANIZER (int) 0x0002 // send bit plus this one
#define RECIP_EXCEPTIONAL_RESPONSE (int) 0x0010 // recipient has exceptional response
#define RECIP_EXCEPTIONAL_DELETED (int) 0x0020 // recipient is NOT in this exception
#define RECIP_ORIGINAL (int) 0x0100 // this was an original recipient on the meeting request

#define respNone 0
#define respOrganized 1
#define respTentative 2
#define respAccepted 3
#define respDeclined 4
#define respNotResponded 5

// [MS-OXOFLAG].pdf
#define PR_TODO_ITEM_FLAGS PROP_TAG(PT_LONG, 0x0e2b)
#define TDIP_None 0x00000000
#define TDIP_Active 0x00000001 // Object is time flagged
#define TDIP_ActiveRecip \
	0x00000008 // SHOULD only be set on a draft message object, and means that the object is flagged for recipients.

// [MS-OXOSRCH].pdf
#define PR_WB_SF_STORAGE_TYPE PROP_TAG(PT_LONG, 0x6846)
enum t_SmartFolderStorageType
{
	SFST_NUMBER = 0x01, // for template's data (numbers)
	SFST_TEXT = 0x02, // for template's data (strings)
	SFST_BINARY = 0x04, // for template's data (binary form, such as entry id, etc.)
	SFST_MRES = 0x08, // for condition in MRES format
	SFST_FILTERSTREAM = 0x10, // for condition in IStream format
	SFST_FOLDERNAME = 0x20, // for folder list's names
	SFST_FOLDERLIST = 0x40, // for the binary folder entrylist

	SFST_TIMERES_MONTHLY = 0x1000, // monthly update
	SFST_TIMERES_WEEKLY = 0x2000, // weekly update
	SFST_TIMERES_DAILY = 0x4000, // the restriction(or filter) has a time condition in it
	SFST_DEAD = 0x8000, // used to indicate there is not a valid SPXBIN
};

// [MS-OXOCAL].pdf
#define dispidBusyStatus 0x8205
enum OlBusyStatus
{
	olFree = 0,
	olTentative = 1,
	olBusy = 2,
	olOutOfOffice = 3,
};
#define dispidApptAuxFlags 0x8207
#define auxApptFlagCopied 0x0001
#define auxApptFlagForceMtgResponse 0x0002
#define auxApptFlagForwarded 0x0004

#define dispidApptColor 0x8214
#define dispidApptStateFlags 0x8217
#define asfNone 0x0000
#define asfMeeting 0x0001
#define asfReceived 0x0002
#define asfCancelled 0x0004
#define asfForward 0x0008

#define dispidResponseStatus 0x8218
#define dispidRecurType 0x8231
#define rectypeNone (int) 0
#define rectypeDaily (int) 1
#define rectypeWeekly (int) 2
#define rectypeMonthly (int) 3
#define rectypeYearly (int) 4

#define dispidConfType 0x8241
enum confType
{
	confNetMeeting = 0,
	confNetShow,
	confExchange
};

#define dispidChangeHighlight 0x8204
#define BIT_CH_START 0x00000001
#define BIT_CH_END 0x00000002
#define BIT_CH_RECUR 0x00000004
#define BIT_CH_LOCATION 0x00000008
#define BIT_CH_SUBJECT 0x00000010
#define BIT_CH_REQATT 0x00000020
#define BIT_CH_OPTATT 0x00000040
#define BIT_CH_BODY 0x00000080
#define BIT_CH_CUSTOM 0x00000100
#define BIT_CH_SILENT 0x00000200
#define BIT_CH_ALLOWPROPOSE 0x00000400
#define BIT_CH_CONF 0x00000800
#define BIT_CH_ATT_REM 0x00001000
#define BIT_CH_NOTUSED 0x80000000

#define dispidMeetingType 0x0026
#define mtgEmpty 0x00000000
#define mtgRequest 0x00000001
#define mtgFull 0x00010000
#define mtgInfo 0x00020000
#define mtgOutofDate 0x00080000
#define mtgDelegatorCopy 0x00100000

#define dispidIntendedBusyStatus 0x8224

#define PR_ATTACHMENT_FLAGS PROP_TAG(PT_LONG, 0x7FFD)
#define afException 0x02 // This is an exception to a recurrence

#define dispidNonSendToTrackStatus 0x8543
#define dispidNonSendCcTrackStatus 0x8544
#define dispidNonSendBccTrackStatus 0x8545

#define LID_CALENDAR_TYPE 0x001C
#define CAL_DEFAULT 0
#define CAL_JAPAN_LUNAR 14
#define CAL_CHINESE_LUNAR 15
#define CAL_SAKA 16
#define CAL_LUNAR_KOREAN 20

// [MS-OXCMSG].pdf
#define dispidSideEffects 0x8510
#define seOpenToDelete 0x0001 // Additional processing is required on the Message object when deleting.
#define seNoFrame 0x0008 // No UI is associated with the Message object.
#define seCoerceToInbox 0x0010 // Additional processing is required on the Message object when moving or
// copying to a Folder object with a PR_CONTAINER_CLASS of 'IPF.Note'.
#define seOpenToCopy 0x0020 // Additional processing is required on the Message object when copying to
// another folder.
#define seOpenToMove 0x0040 // Additional processing is required on the Message object when moving to
// another folder.
#define seOpenForCtxMenu 0x0100 // Additional processing is required on the Message object when displaying verbs
// to the end-user.
#define seCannotUndoDelete 0x0400 // Cannot undo delete operation, MUST NOT be set unless seOpenToDelete is set.
#define seCannotUndoCopy 0x0800 // Cannot undo copy operation, MUST NOT be set unless seOpenToCopy is set.
#define seCannotUndoMove 0x1000 // Cannot undo move operation, MUST NOT be set unless seOpenToMove is set.
#define seHasScript 0x2000 // The Message object contains end-user script.
#define seOpenToPermDelete 0x4000 // Additional processing is required to permanently delete the Message object.

// [MS-OXOCNTC].pdf
#define dispidFileUnderId 0x8006
#define dispidFileUnderList 0x8026
enum
{
	FILEUNDERID_NONE = 0,
	FILEUNDERID_CUSTOM = 0xffffffff,
	FILEUNDERID_CALLINIT = 0xfffffffe,
	FILEUNDERID_CALCULATE = 0xfffffffd
};
#define dispidLastNameAndFirstName 0x8017
#define dispidCompanyAndFullName 0x8018
#define dispidFullNameAndCompany 0x8019
#define dispidLastFirstNoSpace 0x8030
#define dispidLastFirstSpaceOnly 0x8031
#define dispidCompanyLastFirstNoSpace 0x8032
#define dispidCompanyLastFirstSpaceOnly 0x8033
#define dispidLastFirstNoSpaceCompany 0x8034
#define dispidLastFirstSpaceOnlyCompany 0x8035
#define dispidLastFirstAndSuffix 0x8036
#define dispidFirstMiddleLastSuffix 0x8037
#define dispidLastFirstNoSpaceAndSuffix 0x8038

#define dispidPostalAddressId 0x8022
enum PostalAddressIndex
{
	ADDRESS_NONE = 0,
	ADDRESS_HOME,
	ADDRESS_WORK,
	ADDRESS_OTHER
};

// [MS_OXOTASK].pdf
#define dispidTaskMode 0x8518
enum TaskDelegMsgType
{
	tdmtNothing = 0, // The task object is not assigned.
	tdmtTaskReq, // The task object is embedded in a task request.
	tdmtTaskAcc, // The task object has been accepted by the task assignee.
	tdmtTaskDec, // The task object was rejected by the task assignee.
	tdmtTaskUpd, // The task object is embedded in a task update.
	tdmtTaskSELF // The task object was assigned to the task assigner (self-delegation).
};

#define dispidTaskStatus 0x8101
enum TaskStatusValue
{
	tsvNotStarted = 0, // The user has not started work on the task object. If this value is set,
	// dispidPercentComplete MUST be 0.0.
	tsvInProgress, // The user's work on this task object is in progress. If this value is set,
	// dispidPercentComplete MUST be greater than 0.0 and less than 1.0
	tsvComplete, // The user's work on this task object is complete. If this value is set,
	// dispidPercentComplete MUST be 1.0, dispidTaskDateCompleted
	// MUST be the current date, and dispidTaskComplete MUST be true.
	tsvWaiting, // The user is waiting on somebody else.
	tsvDeferred // The user has deferred work on the task object.
};

#define dispidTaskState 0x8113
enum TaskDelegState
{
	tdsNOM = 0, // This task object was created to correspond to a task object that was
	// embedded in a task rejection but could not be found locally.
	tdsOWNNEW, // The task object is not assigned.
	tdsOWN, // The task object is the task assignee's copy of an assigned task object.
	tdsACC, // The task object is the task assigner's copy of an assigned task object.
	tdsDEC // The task object is the task assigner's copy of a rejected task object.
};

#define dispidTaskHistory 0x811A
enum TaskHistory
{
	thNone = 0, // No changes were made.
	thAccepted, // The task assignee accepted this task object.
	thDeclined, // The task assignee rejected this task object.
	thUpdated, // Another property was changed.
	thDueDateChanged, // The dispidTaskDueDate property changed.
	thAssigned // The task object has been assigned to a task assignee.
};

#define dispidTaskMultRecips 0x8120
enum TaskMultRecips
{
	tmrNone = 0x0000, // none
	tmrSent = 0x0001, // The task object has multiple primary recipients.
	tmrRcvd = 0x0002, // Although the 'Sent' hint was not present, the client detected
	// that the task object has multiple primary recipients.
	tmrTeamTask = 0x0004, // This value is reserved.
};

#define dispidTaskOwnership 0x8129
enum TaskOwnershipValue
{
	tovNew, // The task object is not assigned.
	tovDeleg, // The task object is the task assigner's copy of the task object.
	tovMe // The task object is the task assignee's copy of the task object.
};

#define dispidTaskDelegValue 0x812A
enum TaskDelegValue
{
	tdvNone, // The task object is not assigned.
	tdvUnknown, // The task object's acceptance status is unknown.
	tdvAccepted, // The task assignee has accepted the task object. This value is set when
	// the client processes a task acceptance.
	tdvDeclined // The task assignee has rejected the task object. This value is set when the
	// client processes a task rejection.
};

#define dispidLogFlags 0x870C
#define lfContactLog ((ULONG) 0x40000000L) // This journal object has a journal associated attachment

// [MS-OXOMSG].pdf
#define dispidSniffState 0x851A
enum SniffState
{
	ssNone, // Don't auto-process the message.
	ssOnSniff, // Process the message automatically or when the message is opened.
	ssOnOpen // Process when the message is opened only.
};

#define dispidNoteColor 0x8B00
#define ncBlue 0
#define ncGreen 1
#define ncPink 2
#define ncYellow 3
#define ncWhite 4

// http://msdn.microsoft.com/en-us/library/bb821181.aspx
#define dispidTimeZoneStruct 0x8233
#define dispidApptTZDefStartDisplay 0x825E
#define dispidApptTZDefEndDisplay 0x825F
#define dispidApptTZDefRecur 0x8260

#define PR_EXTENDED_FOLDER_FLAGS PROP_TAG(PT_BINARY, 0x36DA)
enum ExtendedFolderPropByte
{
	EFPB_FLAGS = 1,
	EFPB_CLSIDID = 2,
	EFPB_SFTAG = 3,
	EFPB_TODO_VERSION = 5,
};
// possible values for PR_EXTENDED_FOLDER_FLAGS
enum
{
	XEFF_NORMAL = 0x00000000,
	XEFF_SHOW_UNREAD_COUNT = 0x00000001,
	XEFF_SHOW_CONTENT_COUNT = 0x00000002,
	XEFF_SHOW_NO_POLICY = 0x00000020,
};

#define dispidFlagStringEnum 0x85C0

// [MS-OXOCAL].pdf
#define dispidApptRecur 0x8216
#define ARO_SUBJECT 0x0001
#define ARO_MEETINGTYPE 0x0002
#define ARO_REMINDERDELTA 0x0004
#define ARO_REMINDER 0x0008
#define ARO_LOCATION 0x0010
#define ARO_BUSYSTATUS 0x0020
#define ARO_ATTACHMENT 0x0040
#define ARO_SUBTYPE 0x0080
#define ARO_APPTCOLOR 0x0100
#define ARO_EXCEPTIONAL_BODY 0x0200

enum IdGroup
{
	IDC_RCEV_PAT_ORB_DAILY = 0x200A,
	IDC_RCEV_PAT_ORB_WEEKLY,
	IDC_RCEV_PAT_ORB_MONTHLY,
	IDC_RCEV_PAT_ORB_YEARLY,
	IDC_RCEV_PAT_ERB_END = 0x2021,
	IDC_RCEV_PAT_ERB_AFTERNOCCUR,
	IDC_RCEV_PAT_ERB_NOEND,
};

enum
{
	rptMinute = 0,
	rptWeek,
	rptMonth,
	rptMonthNth,
	rptMonthEnd,
	rptHjMonth = 10,
	rptHjMonthNth,
	rptHjMonthEnd
};

constexpr ULONG rpn1st = 1;
constexpr ULONG rpn2nd = 2;
constexpr ULONG rpn3rd = 3;
constexpr ULONG rpn4th = 4;
constexpr ULONG rpnLast = 5;

constexpr ULONG rdfSun = 0x01;
constexpr ULONG rdfMon = 0x02;
constexpr ULONG rdfTue = 0x04;
constexpr ULONG rdfWed = 0x08;
constexpr ULONG rdfThu = 0x10;
constexpr ULONG rdfFri = 0x20;
constexpr ULONG rdfSat = 0x40;

// [MS-OXOTASK].pdf
#define dispidTaskRecur 0x8116
#define dispidTaskMyDelegators 0x8117

// [MS-OXOCAL].pdf
#define LID_GLOBAL_OBJID 0x0003
#define LID_CLEAN_GLOBAL_OBJID 0x0023

// [MS-OXOSRCH].pdf
#define PR_WB_SF_DEFINITION PROP_TAG(PT_BINARY, 0x6845)
#define PR_WB_SF_LAST_USED PROP_TAG(PT_LONG, 0x6834)
#define PR_WB_SF_EXPIRATION PROP_TAG(PT_LONG, 0x683A)
#define PT_SVREID ((ULONG) 0x00FB)

// [MS-OXOPFFB].pdf
#define PR_FREEBUSY_PUBLISH_START PROP_TAG(PT_LONG, 0x6847)
#define PR_FREEBUSY_PUBLISH_END PROP_TAG(PT_LONG, 0x6848)

// [MS-OSCDATA].pdf
#define OOP_DONT_LOOKUP ((ULONG) 0x10000000)

// [MS-OXCDATA].pdf
#define eitLTPrivateFolder (0x1) // PrivateFolder
#define eitLTPublicFolder (0x3) // PublicFolder
#define eitLTWackyFolder (0x5) // MappedPublicFolder
#define eitLTPrivateMessage (0x7) // PrivateMessage
#define eitLTPublicMessage (0x9) // PublicMessage
#define eitLTWackyMessage (0xB) // MappedPublicMessage
#define eitLTPublicFolderByName (0xC) // PublicNewsgroupFolder

// Exchange Address Book Version
#define EMS_VERSION 0x000000001

// Wrapped Message Store Version
#define MAPIMDB_VERSION ((BYTE) 0x00)

// Wrapped Message Store Flag
#define MAPIMDB_NORMAL ((BYTE) 0x00)

// Contact Address Book Version
#define CONTAB_VERSION 0x000000003

// Contact Addess Book Types
#define CONTAB_ROOT 0x00000001
#define CONTAB_SUBROOT 0x00000002
#define CONTAB_CONTAINER 0x00000003
#define CONTAB_USER 0x00000004
#define CONTAB_DISTLIST 0x00000005

// Contact Address Book Index
enum EEmailIndex
{
	EEI_EMAIL_1 = 0,
	EEI_EMAIL_2,
	EEI_EMAIL_3,
	EEI_FAX_1,
	EEI_FAX_2,
	EEI_FAX_3,
};
#define EMAIL_TYPE_UNDEFINED 0xFF

enum
{
	BFLAGS_INTERNAL_MAILUSER = 0x03, // Outlook Contact
	BFLAGS_INTERNAL_DISTLIST, // Outlook Distribution List
	BFLAGS_EXTERNAL_MAILUSER, // external (MAPI) Contact
	BFLAGS_EXTERNAL_DISTLIST, // external (MAPI) Distribution List
	BFLAGS_MASK_OUTLOOK = 0x80, // bit pattern 1000 0000
};

#define dispidEmail1OriginalEntryID 0x8085
#define dispidEmail2OriginalEntryID 0x8095
#define dispidEmail3OriginalEntryID 0x80A5
#define dispidFax1EntryID 0x80B5
#define dispidFax2EntryID 0x80C5
#define dispidFax3EntryID 0x80D5
#define dispidSelectedOriginalEntryID 0x800A
#define dispidAnniversaryEventEID 0x804E
#define dispidOrigStoreEid 0x8237
#define dispidReferenceEID 0x85BD
#define dispidSharingInitiatorEid 0x8A09
#define dispidSharingFolderEid 0x8A15
#define dispidSharingOriginalMessageEid 0x8A29
#define dispidSharingBindingEid 0x8A2D
#define dispidSharingIndexEid 0x8A2E
#define dispidSharingParentBindingEid 0x8A5C
#define dispidVerbStream 0x8520

// Property Definition Stream
#define PropDefV1 0x102
#define PropDefV2 0x103

#define PDO_IS_CUSTOM 0x00000001 // This FieldDefinition structure contains a definition of a user-defined field.
#define PDO_REQUIRED \
	0x00000002 // For a form control bound to this field, the checkbox for A value is required for this field is selected in the Validation tab of the Properties dialog box.
#define PDO_PRINT_SAVEAS \
	0x00000004 // For a form control bound to this field, the checkbox for Include this field for printing and Save As is selected in the Validation tab of the Properties dialog box.
#define PDO_CALC_AUTO \
	0x00000008 // For a form control bound to this field, the checkbox for Calculate this formula automatically option is selected in the Value tab of the Properties dialog box.
#define PDO_FT_CONCAT \
	0x00000010 // This is a field of type Combination and it has the 'Joining fields and any text fragments with each other' option selected in its Combination Formula Field dialog.
#define PDO_FT_SWITCH \
	0x00000020 // This field is of type Combination and has the Showing only the first non-empty field, ignoring subsequent ones option selected in the Combination Formula Field dialog box.
#define PDO_PRINT_SAVEAS_DEF 0x000000040 // This flag is not used

enum iTypeEnum
{
	iTypeUnknown = -1,
	iTypeString, // 0
	iTypeNumber, // 1
	iTypePercent, // 2
	iTypeCurrency, // 3
	iTypeBool, // 4
	iTypeDateTime, // 5
	iTypeDuration, // 6
	iTypeCombination, // 7
	iTypeFormula, // 8
	iTypeIntResult, // 9
	iTypeVariant, // 10
	iTypeFloatResult, // 11
	iTypeConcat, // 12
	iTypeKeywords, // 13
	iTypeInteger, // 14
};

// [MS-OXOSFLD].pdf
#define PR_ADDITIONAL_REN_ENTRYIDS_EX PROP_TAG(PT_BINARY, 0x36d9)

#define RSF_PID_TREAT_AS_SF 0x8000
#define RSF_PID_RSS_SUBSCRIPTION (RSF_PID_TREAT_AS_SF | 1)
#define RSF_PID_SEND_AND_TRACK (RSF_PID_TREAT_AS_SF | 2)
#define RSF_PID_TODO_SEARCH (RSF_PID_TREAT_AS_SF | 4)
#define RSF_PID_CONV_ACTIONS (RSF_PID_TREAT_AS_SF | 6)
#define RSF_PID_COMBINED_ACTIONS (RSF_PID_TREAT_AS_SF | 7)
#define RSF_PID_SUGGESTED_CONTACTS (RSF_PID_TREAT_AS_SF | 8)

#define RSF_ELID_ENTRYID 1
#define RSF_ELID_HEADER 2

#define PR_TARGET_ENTRYID PROP_TAG(PT_BINARY, 0x3010)

#define PR_SCHDINFO_DELEGATE_ENTRYIDS PROP_TAG(PT_MV_BINARY, 0x6845)

#define PR_EMSMDB_SECTION_UID PROP_TAG(PT_BINARY, 0x3d15)

#define EDK_PROFILEUISTATE_ENCRYPTNETWORK 0x4000

#define PR_CONTAB_FOLDER_ENTRYIDS PROP_TAG(PT_MV_BINARY, 0x6620)
#define PR_CONTAB_STORE_ENTRYIDS PROP_TAG(PT_MV_BINARY, 0x6626)
#define PR_CONTAB_STORE_SUPPORT_MASKS PROP_TAG(PT_MV_LONG, 0x6621)
#define PR_DELEGATE_FLAGS PROP_TAG(PT_MV_LONG, 0x686b)

// [MS-OXOCNTC].pdf
#define dispidDLOneOffMembers 0x8054
#define dispidDLMembers 0x8055
#define dispidABPEmailList 0x8028
#define dispidABPArrayType 0x8029

#define PR_FOLDER_WEBVIEWINFO PROP_TAG(PT_BINARY, 0x36DF)

#define WEBVIEW_PERSISTENCE_VERSION 0x000000002
#define WEBVIEWURL 0x00000001
#define WEBVIEW_FLAGS_SHOWBYDEFAULT 0x00000001

#define dispidContactLinkEntry 0x8585

#define dispidApptUnsendableRecips 0x825D
#define dispidForwardNotificationRecipients 0x8261

// http://msdn.microsoft.com/en-us/library/ee218029(EXCHG.80).aspx
// #define PR_NATIVE_BODY_INFO PROP_TAG(PT_LONG, 0x1016)
enum NBI
{
	nbiUndefined = 0,
	nbiPlainText,
	nbiRtfCompressed,
	nbiHtml,
	nbiClearSigned,
};

#define ptagSenderFlags PROP_TAG(PT_LONG, 0x4019)
#define ptagSentRepresentingFlags PROP_TAG(PT_LONG, 0x401A)

#define PR_LAST_VERB_EXECUTED PROP_TAG(PT_LONG, 0x1081)
#define NOTEIVERB_OPEN 0
#define NOTEIVERB_REPLYTOSENDER 102
#define NOTEIVERB_REPLYTOALL 103
#define NOTEIVERB_FORWARD 104
#define NOTEIVERB_PRINT 105
#define NOTEIVERB_SAVEAS 106
#define NOTEIVERB_REPLYTOFOLDER 108
#define NOTEIVERB_SAVE 500
#define NOTEIVERB_PROPERTIES 510
#define NOTEIVERB_FOLLOWUP 511
#define NOTEIVERB_ACCEPT 512
#define NOTEIVERB_TENTATIVE 513
#define NOTEIVERB_REJECT 514
#define NOTEIVERB_DECLINE 515
#define NOTEIVERB_INVITE 516
#define NOTEIVERB_UPDATE 517
#define NOTEIVERB_CANCEL 518
#define NOTEIVERB_SILENTINVITE 519
#define NOTEIVERB_SILENTCANCEL 520
#define NOTEIVERB_RECALL_MESSAGE 521
#define NOTEIVERB_FORWARD_RESPONSE 522
#define NOTEIVERB_FORWARD_CANCEL 523
#define NOTEIVERB_FOLLOWUPCLEAR 524
#define NOTEIVERB_FORWARD_APPT 525
#define NOTEIVERB_OPENRESEND 526
#define NOTEIVERB_STATUSREPORT 527
#define NOTEIVERB_JOURNALOPEN 528
#define NOTEIVERB_JOURNALOPENLINK 529
#define NOTEIVERB_COMPOSEREPLACE 530
#define NOTEIVERB_EDIT 531
#define NOTEIVERB_DELETEPROCESS 532
#define NOTEIVERB_TENTPNTIME 533
#define NOTEIVERB_EDITTEMPLATE 534
#define NOTEIVERB_FINDINCALENDAR 535
#define NOTEIVERB_FORWARDASFILE 536
#define NOTEIVERB_CHANGE_ATTENDEES 537
#define NOTEIVERB_RECALC_TITLE 538
#define NOTEIVERB_PROP_CHANGE 539
#define NOTEIVERB_FORWARD_AS_VCAL 540
#define NOTEIVERB_FORWARD_AS_ICAL 541
#define NOTEIVERB_FORWARD_AS_BCARD 542
#define NOTEIVERB_DECLPNTIME 543
#define NOTEIVERB_PROCESS 544
#define NOTEIVERB_OPENWITHWORD 545
#define NOTEIVERB_OPEN_INSTANCE_OF_SERIES 546
#define NOTEIVERB_FILLOUT_THIS_FORM 547
#define NOTEIVERB_FOLLOWUP_DEFAULT 548
#define NOTEIVERB_REPLY_WITH_MAIL 549
#define NOTEIVERB_TODO_TODAY 566
#define NOTEIVERB_TODO_TOMORROW 567
#define NOTEIVERB_TODO_THISWEEK 568
#define NOTEIVERB_TODO_NEXTWEEK 569
#define NOTEIVERB_TODO_THISMONTH 570
#define NOTEIVERB_TODO_NEXTMONTH 571
#define NOTEIVERB_TODO_NODATE 572
#define NOTEIVERB_FOLLOWUPCOMPLETE 573
#define NOTEIVERB_COPYTOPOSTFOLDER 574
#define NOTEIVERB_PARTIALRECIP_SILENTINVITE 575
#define NOTEIVERB_PARTIALRECIP_SILENTCANCEL 576

#define PR_NT_USER_NAME_W CHANGE_PROP_TYPE(PR_NT_USER_NAME, PT_UNICODE)

#define PR_PROFILE_OFFLINE_STORE_PATH_A PROP_TAG(PT_STRING8, 0x6610)

// Documented via widely shipped mapisvc.inf files.
#define CONFIG_USE_SMTP_ADDRESSES ((ULONG) 0x00000040)

#define CONFIG_OST_CACHE_PRIVATE ((ULONG) 0x00000180)
#define CONFIG_OST_CACHE_PUBLIC ((ULONG) 0x00000400)

#define PR_STORE_UNICODE_MASK PROP_TAG(PT_LONG, 0x340F)

#define PR_PST_CONFIG_FLAGS PROP_TAG(PT_LONG, 0x6770)
#define PST_CONFIG_CREATE_NOWARN 0x00000001
#define PST_CONFIG_PRESERVE_DISPLAY_NAME 0x00000002
#define OST_CONFIG_POLICY_DELAY_IGNORE_OST 0x00000008
#define OST_CONFIG_CREATE_NEW_DEFAULT_OST 0x00000010
#define PST_CONFIG_UNICODE 0x80000000

// http://msdn.microsoft.com/en-us/library/dd941354.aspx
#define dispidImgAttchmtsCompressLevel 0x8593
enum PictureCompressLevel
{
	pclOriginal = 0,
	pclEmail = 1,
	pclWeb = 2,
	pclDocuments = 3,
};

// http://msdn.microsoft.com/en-us/library/dd941362.aspx
#define dispidOfflineStatus 0x85B9
enum MSOCOST
{
	costNil,
	costCheckedOut,
	costSimpleOffline
};

// http://msdn.microsoft.com/en-us/library/ff368035(EXCHG.80).aspx
#define dispidClientIntent 0x0015
enum ClientIntent
{
	ciNone = 0x0000,
	ciManager = 0x0001,
	ciDelegate = 0x0002,
	ciDeletedWithNoResponse = 0x0004,
	ciDeletedExceptionWithNoResponse = 0x0008,
	ciRespondedTentative = 0x0010,
	ciRespondedAccept = 0x0020,
	ciRespondedDecline = 0x0040,
	ciModifiedStartTime = 0x0080,
	ciModifiedEndTime = 0x0100,
	ciModifiedLocation = 0x0200,
	ciRespondedExceptionDecline = 0x0400,
	ciCanceled = 0x0800,
	ciExceptionCanceled = 0x1000,
};

// http://support.microsoft.com/kb/278321
#define INTERNET_FORMAT_DEFAULT 0
#define INTERNET_FORMAT_MIME 1
#define INTERNET_FORMAT_UUENCODE 2
#define INTERNET_FORMAT_BINHEX 3

// http://msdn.microsoft.com/en-us/library/cc765809.aspx
#define dispidContactItemData 0x8007
#define dispidEmail1DisplayName 0x8080
#define dispidEmail2DisplayName 0x8090
#define dispidEmail3DisplayName 0x80A0

// http://blogs.msdn.com/b/stephen_griffin/archive/2010/09/13/you-chose-wisely.aspx
#define PR_AB_CHOOSE_DIRECTORY_AUTOMATICALLY PROP_TAG(PT_BOOLEAN, 0x3D1C)

// http://blogs.msdn.com/b/stephen_griffin/archive/2011/04/13/setsearchpath-not-really.aspx
enum SearchPathReorderType
{
	SEARCHPATHREORDERTYPE_RAW = 0,
	SEARCHPATHREORDERTYPE_ACCT_PREFERGAL,
	SEARCHPATHREORDERTYPE_ACCT_PREFERCONTACTS,
};

#define PR_USERFIELDS PROP_TAG(PT_BINARY, 0x36E3)

#define ftNull 0x00
#define ftString 0x01
#define ftInteger 0x03
// This definition conflicts with Windows headers.
// Was only use for Flags.h and we can handle that differently.
// Definition left commened for reference
//#define ftTime 0x05
#define ftBoolean 0x06
#define ftDuration 0x07
#define ftMultiString 0x0B
#define ftFloat 0x0C
#define ftCurrency 0x0E
#define ftCalc 0x12
#define ftSwitch 0x13
#define ftConcat 0x17

#define FCAPM_CAN_EDIT 0x00000001 // field is editable
#define FCAPM_CAN_SORT 0x00000002 // field is sortable
#define FCAPM_CAN_GROUP 0x00000004 // Field is groupable
#define FCAPM_MULTILINE_TEXT 0x00000100 // can hold multiple lines of text
#define FCAPM_PERCENT 0x01000000 // FLOAT this is a percentage field
#define FCAPM_DATEONLY 0x01000000 // TIME a date-only 'time' field
#define FCAPM_UNITLESS 0x01000000 // INT no unit allowed in display
#define FCAPM_CAN_EDIT_IN_ITEM 0x80000000 // Field can be edited in the item (specifically for custom forms)

#define PR_SPAM_THRESHOLD PROP_TAG(PT_LONG, 0x041B)
#define SPAM_FILTERING_NONE 0xFFFFFFFF
#define SPAM_FILTERING_LOW 0x00000006
#define SPAM_FILTERING_MEDIUM 0x00000005
#define SPAM_FILTERING_HIGH 0x00000003
#define SPAM_FILTERING_TRUSTED_ONLY 0x80000000

// Constants - http://msdn2.microsoft.com/en-us/library/bb905201.aspx
enum CCSFLAGS
{
	CCSF_SMTP = 0x00000002, // the converter is being passed an SMTP message
	CCSF_NOHEADERS = 0x00000004, // the converter should ignore the headers on the outside message
	CCSF_USE_TNEF = 0x00000010, // the converter should embed TNEF in the MIME message
	CCSF_INCLUDE_BCC = 0x00000020, // the converter should include Bcc recipients
	CCSF_8BITHEADERS = 0x00000040, // the converter should allow 8 bit headers
	CCSF_USE_RTF = 0x00000080, // the converter should do HTML->RTF conversion
	CCSF_PLAIN_TEXT_ONLY = 0x00001000, // the converter should just send plain text
	CCSF_NO_MSGID = 0x00004000, // don't include Message-Id from MAPI message in outgoing messages create a new one
	CCSF_EMBEDDED_MESSAGE = 0x00008000, // We're translating an embedded message
	CCSF_PRESERVE_SOURCE =
		0x00040000, // The convertor should not modify the source message so no conversation index update, no message id, and no header dump.
	CCSF_GLOBAL_MESSAGE = 0x00200000, // The converter should build an international message (EAI/RFC6530)
};

#define PidTagFolderId PROP_TAG(PT_I8, 0x6748)
#define PidTagMid PROP_TAG(PT_I8, 0x674A)

#define MDB_STORE_EID_V1_VERSION (0x0)
#define MDB_STORE_EID_V2_VERSION (0x1)
#define MDB_STORE_EID_V3_VERSION (0x2)
#define MDB_STORE_EID_V2_MAGIC (0xf32135d8)
#define MDB_STORE_EID_V3_MAGIC (0xf43246e9)

// http://blogs.msdn.com/b/dvespa/archive/2009/03/16/how-to-sign-or-encrypt-a-message-programmatically-from-oom.aspx
#define PR_SECURITY_FLAGS PROP_TAG(PT_LONG, 0x6E01)
#define SECFLAG_NONE 0x0000 // Message has no security
#define SECFLAG_ENCRYPTED 0x0001 // Message is sealed
#define SECFLAG_SIGNED 0x0002 // Message is signed

#define PR_ROAMING_BINARYSTREAM PROP_TAG(PT_BINARY, 0x7C09)

#define PR_QUOTA_WARNING PROP_TAG(PT_LONG, 0x341A)
#define PR_QUOTA_SEND PROP_TAG(PT_LONG, 0x341B)
#define PR_QUOTA_RECEIVE PROP_TAG(PT_LONG, 0x341C)

#define PR_SCHDINFO_APPT_TOMBSTONE PROP_TAG(PT_BINARY, 0x686A)

#define PRXF_IGNORE_SEC_WARNING 0x10

#define PR_WLINK_AB_EXSTOREEID PROP_TAG(PT_BINARY, 0x6891)
#define PR_WLINK_ABEID PROP_TAG(PT_BINARY, 0x6854)
#define PR_WLINK_ENTRYID PROP_TAG(PT_BINARY, 0x684C)
#define PR_WLINK_STORE_ENTRYID PROP_TAG(PT_BINARY, 0x684E)
#define PR_WLINK_RO_GROUP_TYPE PROP_TAG(PT_LONG, 0x6892)
#define PR_WLINK_SECTION PROP_TAG(PT_LONG, 0x6852)
#define PR_WLINK_TYPE PROP_TAG(PT_LONG, 0x6849)
#define PR_WLINK_FLAGS PROP_TAG(PT_LONG, 0x684A)

enum WBROGroupType
{
	wbrogUndefined = -1,
	wbrogMyDepartment = 0,
	wbrogOtherDepartment,
	wbrogDirectReportGroup,
	wbrogCoworkerGroup,
	wbrogDL,
};

enum WBSID
{
	wbsidMailFavorites = 1,
	wbsidCalendar = 3,
	wbsidContacts,
	wbsidTasks,
	wbsidNotes,
	wbsidJournal,
};

enum WLinkType
{
	wblNormalFolder = 0,
	wblSearchFolder = 1,
	wblSharedFolder = 2,
	wblHeader = 4,
};

#define sipPublicFolder 0x00000001
#define sipImapFolder 0x00000004
#define sipWebDavFolder 0x00000008
#define sipSharePointFolder 0x00000010
#define sipRootFolder 0x00000020
#define sipSharedOut 0x00000100
#define sipSharedIn 0x00000200
#define sipPersonFolder 0x00000400
#define sipiCal 0x00000800
#define sipOverlay 0x00001000
#define sipOneOffName 0x00002000

#define dispidConvActionMoveFolderEid 0x85C6
#define dispidConvActionMoveStoreEid 0x85C7

#define RETENTION_FLAGS_KEEP_IN_PLACE ((ULONG) 0x00000020)
#define RETENTION_FLAGS_SYSTEM_DATA ((ULONG) 0X00000040)
#define RETENTION_FLAGS_NEEDS_RESCAN ((ULONG) 0X00000080)
#define RETENTION_FLAGS_PENDING_RESCAN ((ULONG) 0X00000100)

#define PR_PROFILE_MDB_DN PROP_TAG(PT_STRING8, 0x7CFF)
#define PR_FORCE_USE_ENTRYID_SERVER PROP_TAG(PT_BOOLEAN, 0x7CFE)

#define PR_EMS_AB_THUMBNAIL_PHOTO PROP_TAG(PT_BINARY, 0x8C9E)

#define OUTL_RPC_AUTHN_ANONYMOUS 0x8000F001

#define PR_CONTENT_FILTER_SCL PROP_TAG(PT_LONG, 0x4076)

#define EPHEMERAL (UCHAR)(~(MAPI_NOTRECIP | MAPI_THISSESSION | MAPI_NOW | MAPI_NOTRESERVED))

#define ATTACH_BY_WEB_REF ((ULONG) 0x00000007)

// http://msdn2.microsoft.com/en-us/library/ms527629.aspx
#define MSGFLAG_ORIGIN_X400 ((ULONG) 0x00001000)
#define MSGFLAG_ORIGIN_INTERNET ((ULONG) 0x00002000)
#define MSGFLAG_ORIGIN_MISC_EXT ((ULONG) 0x00008000)
#define MSGFLAG_OUTLOOK_NON_EMS_XP ((ULONG) 0x00010000)

// http://msdn2.microsoft.com/en-us/library/ms528848.aspx
#define MSGSTATUS_DRAFT ((ULONG) 0x00000100)
#define MSGSTATUS_ANSWERED ((ULONG) 0x00000200)

// Various flags gleaned from product documentation and KB articles
// http://msdn2.microsoft.com/en-us/library/ms526744.aspx
#define STORE_HTML_OK ((ULONG) 0x00010000)
#define STORE_ANSI_OK ((ULONG) 0x00020000)
#define STORE_LOCALSTORE ((ULONG) 0x00080000)

// http://msdn2.microsoft.com/en-us/library/bb820937.aspx
#define STORE_PUSHER_OK ((ULONG) 0x00800000)

// http://blogs.msdn.com/stephen_griffin/archive/2006/05/11/595338.aspx
#define STORE_FULLTEXT_QUERY_OK ((ULONG) 0x02000000)
#define STORE_FILTER_SEARCH_OK ((ULONG) 0x04000000)

// http://msdn2.microsoft.com/en-us/library/ms531462.aspx
#define ATT_INVISIBLE_IN_HTML ((ULONG) 0x00000001)
#define ATT_INVISIBLE_IN_RTF ((ULONG) 0x00000002)
#define ATT_MHTML_REF ((ULONG) 0x00000004)

#define ENCODING_PREFERENCE ((ULONG) 0x00020000)
#define ENCODING_MIME ((ULONG) 0x00040000)
#define BODY_ENCODING_HTML ((ULONG) 0x00080000)
#define BODY_ENCODING_TEXT_AND_HTML ((ULONG) 0x00100000)
#define MAC_ATTACH_ENCODING_UUENCODE ((ULONG) 0x00200000)
#define MAC_ATTACH_ENCODING_APPLESINGLE ((ULONG) 0x00400000)
#define MAC_ATTACH_ENCODING_APPLEDOUBLE ((ULONG) 0x00600000)

// http://blogs.msdn.com/stephen_griffin/archive/2007/03/19/mapi-and-exchange-2007.aspx
#define CONNECT_IGNORE_NO_PF ((ULONG) 0x8000)

#define MAPI_NATIVE_BODY 0x00010000

/* out param type infomation for WrapCompressedRTFStreamEx */
#define MAPI_NATIVE_BODY_TYPE_RTF 0x00000001
#define MAPI_NATIVE_BODY_TYPE_HTML 0x00000002
#define MAPI_NATIVE_BODY_TYPE_PLAINTEXT 0x00000004

// Will match prefix on words instead of the whole prop value
#define FL_PREFIX_ON_ANY_WORD 0x00000010

// Phrase match means the words have to be exactly matched and the
// sequence matters. This is different than FL_FULLSTRING because
// it doesn't require the whole property value to be the same. One
// term exactly matching a term in the property value is enough for
// a match even if there are more terms in the property.
#define FL_PHRASE_MATCH 0x00000020

#define fnevIndexing ((ULONG) 0x00010000)

constexpr WORD TZRULE_FLAG_RECUR_CURRENT_TZREG = 0x0001; // see dispidApptTZDefRecur
constexpr WORD TZRULE_FLAG_EFFECTIVE_TZREG = 0x0002;

// http://msdn2.microsoft.com/en-us/library/bb905283.aspx
#define dispidFormStorage 0x850F
#define dispidPageDirStream 0x8513
#define dispidFormPropStream 0x851B
#define dispidPropDefStream 0x8540
#define dispidScriptStream 0x8541
#define dispidCustomFlag 0x8542

#define INSP_ONEOFFFLAGS 0xD000000
#define INSP_PROPDEFINITION 0x2000000

#define TABLE_SORT_CATEG_MAX ((ULONG) 0x00000004)
#define TABLE_SORT_CATEG_MIN ((ULONG) 0x00000008)

// Stores that are pusher enabled (PR_STORE_SUPPORT_MASK contains STORE_PUSHER_OK)
// are required to send notifications regarding the process that is pushing.
#define INDEXING_SEARCH_OWNER ((ULONG) 0x00000001)

struct INDEX_SEARCH_PUSHER_PROCESS
{
	DWORD dwPID; // PID for process pushing
};

#define MAPI_FORCE_ACCESS 0x00080000

// This declaration is missing from the MAPI headers
STDAPI STDAPICALLTYPE LaunchWizard(
	HWND hParentWnd,
	ULONG ulFlags,
	LPCSTR* lppszServiceNameToAdd,
	ULONG cchBufferMax,
	_Out_cap_(cchBufferMax) LPSTR lpszNewProfileName);

#define FWD_AS_SMS_ALERT 8

#define PR_BODY_HTML_W CHANGE_PROP_TYPE(PR_BODY_HTML, PT_UNICODE)

// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxprops/951b479a-dcdb-46fb-a293-9177baa0a4b9
#define PR_SWAPPED_TODO_DATA PROP_TAG(PT_BINARY, 0x0E2D)

#define TDP_NONE 0x00000000
#define TDP_ITEM_FLAGS 0x00000001
#define TDP_START_DATE 0x00000008
#define TDP_DUE_DATE 0x00000010
#define TDP_TITLE 0x00000020
#define TDP_REMINDER_FLAG 0x00000040
#define TDP_REMINDER_TIME 0x00000080

#define RULE_ERR_FOLDER_OVER_QUOTA 14 // the folder quota has been exceeded

#define STORE_RULES_OK ((ULONG) 0x10000000)
