#pragma once
#include <mapidefs.h>

#undef PR_CONVERSION_STATE
#undef PR_DOTSTUFF_STATE
#undef PR_USER_SID

#define PR_CONVERSION_STATE			PROP_TAG(PT_LONG, 0x6778)
#define PR_DOTSTUFF_STATE			PROP_TAG(PT_LONG, 0x6001)
#define PR_USER_SID					PROP_TAG(PT_BINARY, 0x6783)

// http://support.microsoft.com/kb/898835
#define PR_ROH_FLAGS                PROP_TAG(PT_LONG,0x6623)
#define PR_ROH_PROXY_SERVER         PROP_TAG(PT_UNICODE,0x6622)
#define PR_ROH_PROXY_PRINCIPAL_NAME PROP_TAG(PT_UNICODE,0x6625)
#define PR_ROH_PROXY_AUTH_SCHEME    PROP_TAG(PT_LONG,0x6627)
#define	PR_RULE_VERSION				PROP_TAG(PT_I2, pidSpecialMin+0x1D)

#define PR_FREEBUSY_NT_SECURITY_DESCRIPTOR (PROP_TAG(PT_BINARY,0x0F00))

// http://support.microsoft.com/kb/171670
//Entry ID for the Calendar
#define PR_IPM_APPOINTMENT_ENTRYID (PROP_TAG(PT_BINARY,0x36D0))
//Entry ID for the Contact Folder
#define PR_IPM_CONTACT_ENTRYID (PROP_TAG(PT_BINARY,0x36D1))
//Entry ID for the Journal Folder
#define PR_IPM_JOURNAL_ENTRYID (PROP_TAG(PT_BINARY,0x36D2))
//Entry ID for the Notes Folder
#define PR_IPM_NOTE_ENTRYID (PROP_TAG(PT_BINARY,0x36D3))
//Entry ID for the Task Folder
#define PR_IPM_TASK_ENTRYID (PROP_TAG(PT_BINARY,0x36D4))
//Entry IDs for the Reminders Folder
#define PR_REM_ONLINE_ENTRYID (PROP_TAG(PT_BINARY,0x36D5))
#define PR_REM_OFFLINE_ENTRYID PROP_TAG(PT_BINARY, 0x36D6)
//Entry ID for the Drafts Folder
#define PR_IPM_DRAFTS_ENTRYID (PROP_TAG(PT_BINARY,0x36D7))

#define PR_DEF_POST_MSGCLASS PROP_TAG(PT_STRING8, 0x36E5)
#define PR_DEF_POST_DISPLAYNAME	PROP_TAG(PT_TSTRING, 0x36E6)
#define PR_FREEBUSY_ENTRYIDS PROP_TAG(PT_MV_BINARY, 0x36E4)
#ifndef PR_RECIPIENT_TRACKSTATUS
#define PR_RECIPIENT_TRACKSTATUS PROP_TAG(PT_LONG, 0x5FFF)
#endif
#define PR_RECIPIENT_FLAGS PROP_TAG(PT_LONG, 0x5FFD)
#define PR_RECIPIENT_ENTRYID PROP_TAG(PT_BINARY, 0x5FF7)
#define PR_RECIPIENT_DISPLAY_NAME PROP_TAG(PT_TSTRING, 0x5FF6)
#define PR_ICON_INDEX PROP_TAG(PT_LONG, 0x1080)
#define PR_OST_OSTID PROP_TAG(PT_BINARY, 0x7c04)
#define PR_OFFLINE_FOLDER PROP_TAG(PT_LONG, 0x7c05)
#define PR_FAV_PARENT_SOURCE_KEY PROP_TAG(PT_BINARY, 0x7d02)

#ifndef PR_USER_X509_CERTIFICATE
#define PR_USER_X509_CERTIFICATE (PROP_TAG(PT_MV_BINARY,0x3a70))
#endif
#ifndef PR_NT_SECURITY_DESCRIPTOR
#define PR_NT_SECURITY_DESCRIPTOR (PROP_TAG(PT_BINARY,0x0E27))
#endif
#ifndef PR_BODY_HTML
#define PR_BODY_HTML (PROP_TAG(PT_TSTRING,0x1013))
#endif
#ifndef PR_BODY_HTML_A
#define PR_BODY_HTML_A (PROP_TAG(PT_STRING8,0x1013))
#endif
#ifndef PR_BODY_HTML_W
#define PR_BODY_HTML_W (PROP_TAG(PT_UNICODE,0x1013))
#endif

// http://support.microsoft.com/kb/816477
#ifndef PR_MSG_EDITOR_FORMAT
#define PR_MSG_EDITOR_FORMAT PROP_TAG( PT_LONG, 0x5909 )
#endif
#ifndef PR_INTERNET_MESSAGE_ID
#define PR_INTERNET_MESSAGE_ID PROP_TAG(PT_TSTRING, 0x1035)
#endif
#ifndef PR_INTERNET_MESSAGE_ID_A
#define PR_INTERNET_MESSAGE_ID_A PROP_TAG(PT_STRING8, 0x1035)
#endif
#ifndef PR_INTERNET_MESSAGE_ID_W
#define PR_INTERNET_MESSAGE_ID_W PROP_TAG(PT_UNICODE, 0x1035)
#endif
#ifndef PR_SMTP_ADDRESS
#define PR_SMTP_ADDRESS PROP_TAG(PT_TSTRING,0x39FE)
#endif
#ifndef PR_SMTP_ADDRESS_A
#define PR_SMTP_ADDRESS_A PROP_TAG(PT_STRING8,0x39FE)
#endif
#ifndef PR_SMTP_ADDRESS_W
#define PR_SMTP_ADDRESS_W PROP_TAG(PT_UNICODE,0x39FE)
#endif
#ifndef PR_INTERNET_ARTICLE_NUMBER
#define PR_INTERNET_ARTICLE_NUMBER PROP_TAG(PT_LONG, 0x0E23)
#endif
#ifndef PR_SEND_INTERNET_ENCODING
#define PR_SEND_INTERNET_ENCODING PROP_TAG(PT_LONG, 0x3A71)
#endif
#ifndef PR_IN_REPLY_TO_ID
#define PR_IN_REPLY_TO_ID PROP_TAG(PT_TSTRING, 0x1042)
#endif
#ifndef PR_IN_REPLY_TO_ID_A
#define PR_IN_REPLY_TO_ID_A PROP_TAG(PT_STRING8, 0x1042)
#endif
#ifndef PR_IN_REPLY_TO_ID_W
#define PR_IN_REPLY_TO_ID_W PROP_TAG(PT_UNICODE, 0x1042)
#endif
#ifndef PR_ATTACH_MIME_SEQUENCE
#define PR_ATTACH_MIME_SEQUENCE PROP_TAG(PT_LONG, 0x3710)
#endif
#ifndef PR_ATTACH_CONTENT_BASE
#define PR_ATTACH_CONTENT_BASE PROP_TAG(PT_TSTRING, 0x3711)
#endif
#ifndef PR_ATTACH_CONTENT_BASE_A
#define PR_ATTACH_CONTENT_BASE_A PROP_TAG(PT_STRING8, 0x3711)
#endif
#ifndef PR_ATTACH_CONTENT_BASE_W
#define PR_ATTACH_CONTENT_BASE_W PROP_TAG(PT_UNICODE, 0x3711)
#endif
#ifndef PR_ATTACH_CONTENT_ID
#define PR_ATTACH_CONTENT_ID PROP_TAG(PT_TSTRING, 0x3712)
#endif
#ifndef PR_ATTACH_CONTENT_ID_A
#define PR_ATTACH_CONTENT_ID_A PROP_TAG(PT_STRING8, 0x3712)
#endif
#ifndef PR_ATTACH_CONTENT_ID_W
#define PR_ATTACH_CONTENT_ID_W PROP_TAG(PT_UNICODE, 0x3712)
#endif
#ifndef PR_ATTACH_CONTENT_LOCATION
#define PR_ATTACH_CONTENT_LOCATION PROP_TAG(PT_TSTRING, 0x3713)
#endif
#ifndef PR_ATTACH_CONTENT_LOCATION_A
#define PR_ATTACH_CONTENT_LOCATION_A PROP_TAG(PT_STRING8, 0x3713)
#endif
#ifndef PR_ATTACH_CONTENT_LOCATION_W
#define PR_ATTACH_CONTENT_LOCATION_W PROP_TAG(PT_UNICODE, 0x3713)
#endif
#ifndef PR_ATTACH_FLAGS
#define PR_ATTACH_FLAGS PROP_TAG(PT_LONG, 0x3714)
#endif

// http://support.microsoft.com/kb/837364
#ifndef PR_CONFLICT_ITEMS
#define PR_CONFLICT_ITEMS PROP_TAG(PT_MV_BINARY,0x1098)
#endif
#ifndef PR_INTERNET_APPROVED
#define PR_INTERNET_APPROVED PROP_TAG(PT_TSTRING,0x1030)
#endif
#ifndef PR_INTERNET_APPROVED_A
#define PR_INTERNET_APPROVED_A PROP_TAG(PT_STRING8,0x1030)
#endif
#ifndef PR_INTERNET_APPROVED_W
#define PR_INTERNET_APPROVED_W PROP_TAG(PT_UNICODE,0x1030)
#endif
#ifndef PR_INTERNET_CONTROL
#define PR_INTERNET_CONTROL PROP_TAG(PT_TSTRING,0x1031)
#endif
#ifndef PR_INTERNET_CONTROL_A
#define PR_INTERNET_CONTROL_A PROP_TAG(PT_STRING8,0x1031)
#endif
#ifndef PR_INTERNET_CONTROL_W
#define PR_INTERNET_CONTROL_W PROP_TAG(PT_UNICODE,0x1031)
#endif
#ifndef PR_INTERNET_DISTRIBUTION
#define PR_INTERNET_DISTRIBUTION PROP_TAG(PT_TSTRING,0x1032)
#endif
#ifndef PR_INTERNET_DISTRIBUTION_A
#define PR_INTERNET_DISTRIBUTION_A PROP_TAG(PT_STRING8,0x1032)
#endif
#ifndef PR_INTERNET_DISTRIBUTION_W
#define PR_INTERNET_DISTRIBUTION_W PROP_TAG(PT_UNICODE,0x1032)
#endif
#ifndef PR_INTERNET_FOLLOWUP_TO
#define PR_INTERNET_FOLLOWUP_TO PROP_TAG(PT_TSTRING,0x1033)
#endif
#ifndef PR_INTERNET_FOLLOWUP_TO_A
#define PR_INTERNET_FOLLOWUP_TO_A PROP_TAG(PT_STRING8,0x1033)
#endif
#ifndef PR_INTERNET_FOLLOWUP_TO_W
#define PR_INTERNET_FOLLOWUP_TO_W PROP_TAG(PT_UNICODE,0x1033)
#endif
#ifndef PR_INTERNET_LINES
#define PR_INTERNET_LINES PROP_TAG(PT_LONG,0x1034)
#endif
#ifndef PR_INTERNET_NEWSGROUPS
#define PR_INTERNET_NEWSGROUPS PROP_TAG(PT_TSTRING,0x1036)
#endif
#ifndef PR_INTERNET_NEWSGROUPS_A
#define PR_INTERNET_NEWSGROUPS_A PROP_TAG(PT_STRING8,0x1036)
#endif
#ifndef PR_INTERNET_NEWSGROUPS_W
#define PR_INTERNET_NEWSGROUPS_W PROP_TAG(PT_UNICODE,0x1036)
#endif
#ifndef PR_INTERNET_NNTP_PATH
#define PR_INTERNET_NNTP_PATH PROP_TAG(PT_TSTRING,0x1038)
#endif
#ifndef PR_INTERNET_NNTP_PATH_A
#define PR_INTERNET_NNTP_PATH_A PROP_TAG(PT_STRING8,0x1038)
#endif
#ifndef PR_INTERNET_NNTP_PATH_W
#define PR_INTERNET_NNTP_PATH_W PROP_TAG(PT_UNICODE,0x1038)
#endif
#ifndef PR_INTERNET_ORGANIZATION
#define PR_INTERNET_ORGANIZATION PROP_TAG(PT_TSTRING,0x1037)
#endif
#ifndef PR_INTERNET_ORGANIZATION_A
#define PR_INTERNET_ORGANIZATION_A PROP_TAG(PT_STRING8,0x1037)
#endif
#ifndef PR_INTERNET_ORGANIZATION_W
#define PR_INTERNET_ORGANIZATION_W PROP_TAG(PT_UNICODE,0x1037)
#endif
#ifndef PR_INTERNET_PRECEDENCE
#define PR_INTERNET_PRECEDENCE PROP_TAG(PT_TSTRING,0x1041)
#endif
#ifndef PR_INTERNET_PRECEDENCE_A
#define PR_INTERNET_PRECEDENCE_A PROP_TAG(PT_STRING8,0x1041)
#endif
#ifndef PR_INTERNET_PRECEDENCE_W
#define PR_INTERNET_PRECEDENCE_W PROP_TAG(PT_UNICODE,0x1041)
#endif
#ifndef PR_INTERNET_REFERENCES
#define PR_INTERNET_REFERENCES PROP_TAG(PT_TSTRING,0x1039)
#endif
#ifndef PR_INTERNET_REFERENCES_A
#define PR_INTERNET_REFERENCES_A PROP_TAG(PT_STRING8,0x1039)
#endif
#ifndef PR_INTERNET_REFERENCES_W
#define PR_INTERNET_REFERENCES_W PROP_TAG(PT_UNICODE,0x1039)
#endif
#ifndef PR_NEWSGROUP_NAME
#define PR_NEWSGROUP_NAME PROP_TAG(PT_TSTRING,0x0E24)
#endif
#ifndef PR_NEWSGROUP_NAME_A
#define PR_NEWSGROUP_NAME_A PROP_TAG(PT_STRING8,0x0E24)
#endif
#ifndef PR_NEWSGROUP_NAME_W
#define PR_NEWSGROUP_NAME_W PROP_TAG(PT_UNICODE,0x0E24)
#endif
#ifndef PR_NNTP_XREF
#define PR_NNTP_XREF PROP_TAG(PT_TSTRING,0x1040)
#endif
#ifndef PR_NNTP_XREF_A
#define PR_NNTP_XREF_A PROP_TAG(PT_STRING8,0x1040)
#endif
#ifndef PR_NNTP_XREF_W
#define PR_NNTP_XREF_W PROP_TAG(PT_UNICODE,0x1040)
#endif
#ifndef PR_POST_FOLDER_ENTRIES
#define PR_POST_FOLDER_ENTRIES PROP_TAG(PT_BINARY,0x103B)
#endif
#ifndef PR_POST_FOLDER_NAMES
#define PR_POST_FOLDER_NAMES PROP_TAG(PT_TSTRING,0x103C)
#endif
#ifndef PR_POST_FOLDER_NAMES_A
#define PR_POST_FOLDER_NAMES_A PROP_TAG(PT_STRING8,0x103C)
#endif
#ifndef PR_POST_FOLDER_NAMES_W
#define PR_POST_FOLDER_NAMES_W PROP_TAG(PT_UNICODE,0x103C)
#endif
#ifndef PR_POST_REPLY_DENIED
#define PR_POST_REPLY_DENIED PROP_TAG(PT_BOOL,0x103F)
#endif
#ifndef PR_POST_REPLY_FOLDER_ENTRIES
#define PR_POST_REPLY_FOLDER_ENTRIES PROP_TAG(PT_BINARY,0x103D)
#endif
#ifndef PR_POST_REPLY_FOLDER_NAMES
#define PR_POST_REPLY_FOLDER_NAMES PROP_TAG(PT_TSTRING,0x103E)
#endif
#ifndef PR_POST_REPLY_FOLDER_NAMES_A
#define PR_POST_REPLY_FOLDER_NAMES_A PROP_TAG(PT_STRING8,0x103E)
#endif
#ifndef PR_POST_REPLY_FOLDER_NAMES_W
#define PR_POST_REPLY_FOLDER_NAMES_W PROP_TAG(PT_UNICODE,0x103E)
#endif
#ifndef PR_SUPERSEDES
#define PR_SUPERSEDES PROP_TAG(PT_TSTRING,0x103A)
#endif
#ifndef PR_SUPERSEDES_A
#define PR_SUPERSEDES_A PROP_TAG(PT_STRING8,0x103A)
#endif
#ifndef PR_SUPERSEDES_W
#define PR_SUPERSEDES_W PROP_TAG(PT_UNICODE,0x103A)
#endif

//Outlook 2003 Integration API - http://msdn2.microsoft.com/en-us/library/aa193120(office.11).aspx
#ifndef PR_PRIMARY_SEND_ACCT
#define PR_PRIMARY_SEND_ACCT PROP_TAG(PT_UNICODE,0x0E28)
#endif
#ifndef PR_NEXT_SEND_ACCT
#define PR_NEXT_SEND_ACCT PROP_TAG(PT_UNICODE,0x0E29)
#endif

// http://support.microsoft.com/kb/225009
#ifndef PR_PROCESSED
#define PR_PROCESSED PROP_TAG(PT_BOOLEAN, 0x7d01)
#endif

// http://support.microsoft.com/kb/278321
#ifndef PR_INETMAIL_OVERRIDE_FORMAT
#define PR_INETMAIL_OVERRIDE_FORMAT PROP_TAG(PT_LONG,0x5902)
#endif

// http://support.microsoft.com/kb/312900
#ifndef PR_SECURITY_PROFILES
#define PR_SECURITY_PROFILES PROP_TAG(PT_MV_BINARY, 0x355)
#endif
#define PR_CERT_PROP_VERSION            PROP_TAG(PT_LONG,       0x0001)
#define PR_CERT_MESSAGE_ENCODING        PROP_TAG(PT_LONG,       0x0006)
#define PR_CERT_DEFAULTS                PROP_TAG(PT_LONG,       0x0020)
#define PR_CERT_DISPLAY_NAME_A          PROP_TAG(PT_STRING8,    0x000B)
#define PR_CERT_KEYEX_SHA1_HASH         PROP_TAG(PT_BINARY,     0x0022)
#define PR_CERT_SIGN_SHA1_HASH          PROP_TAG(PT_BINARY,     0x0009)
#define PR_CERT_ASYMETRIC_CAPS          PROP_TAG(PT_BINARY,     0x0002)
// Values for PR_CERT_DEFAULTS
#define MSG_DEFAULTS_NONE               0
#define MSG_DEFAULTS_FOR_FORMAT         1 // Default certificate for S/MIME.
#define MSG_DEFAULTS_GLOBAL             2 // Default certificate for all formats.
#define MSG_DEFAULTS_SEND_CERT          4 // Send certificate with message.

// http://support.microsoft.com/kb/912237
#ifndef PR_ATTACHMENT_CONTACTPHOTO
#define PR_ATTACHMENT_CONTACTPHOTO PROP_TAG(PT_BOOLEAN, 0x7fff)
#endif

// http://support.microsoft.com/kb/194955
#define PR_AGING_PERIOD         PROP_TAG(PT_LONG,0x36EC)
#define PR_AGING_GRANULARITY    PROP_TAG(PT_LONG,0x36EE)

// http://msdn2.microsoft.com/en-us/library/bb820938.aspx
#define PR_PROVIDER_ITEMID          	PROP_TAG(PT_BINARY, 	0x0EA3)
// http://msdn2.microsoft.com/en-us/library/bb820939.aspx
#define PR_PROVIDER_PARENT_ITEMID   	PROP_TAG(PT_BINARY, 	0x0EA4)

// PH props
#define PR_SEARCH_ATTACHMENTS_W      PROP_TAG(PT_UNICODE, 0x0EA5)
#define PR_SEARCH_RECIP_EMAIL_TO_W   PROP_TAG(PT_UNICODE, 0x0EA6)
#define PR_SEARCH_RECIP_EMAIL_CC_W   PROP_TAG(PT_UNICODE, 0x0EA7)
#define PR_SEARCH_RECIP_EMAIL_BCC_W  PROP_TAG(PT_UNICODE, 0x0EA8)
#define PR_SEARCH_OWNER_ID           PROP_TAG(PT_LONG,    0x3419)

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
//       sfConflicts       0
//       sfSyncFailures    1
//       sfLocalFailures   2
//       sfServerFailures  3
//       sfJunkEmail       4
//       sfSpamTagDontUse  5
//
// NOTE: sfSpamTagDontUse is not the real special folder but used #5 slot
// Therefore, we need to skip it when enum through sf* special folders.
#define PR_ADDITIONAL_REN_ENTRYIDS    PROP_TAG(PT_MV_BINARY, 0x36D8)

// http://msdn2.microsoft.com/en-us/library/bb820966.aspx
#define	PR_PROFILE_SERVER_FULL_VERSION	PROP_TAG( PT_BINARY, pidProfileMin+0x3b)

// http://msdn2.microsoft.com/en-us/library/bb820973.aspx
// Additional display attributes, to supplement PR_DISPLAY_TYPE.
#define PR_DISPLAY_TYPE_EX PROP_TAG( PT_LONG, 0x3905)

// PR_DISPLAY_TYPE_EX has the following format
// 
// 33222222222211111111110000000000
// 10987654321098765432109876543210
//
// FAxxxxxxxxxxxxxxRRRRRRRRLLLLLLLL
//
// F = 1 if remote is valid, 0 if it is not
// A = 1 if the user is ACL-able, 0 if the user is not
// x - unused at this time, do not interpret as this may be used in the future
// R = display type from 

#define DTE_FLAG_REMOTE_VALID 0x80000000
#define DTE_FLAG_ACL_CAPABLE  0x40000000
#define DTE_MASK_REMOTE       0x0000ff00
#define DTE_MASK_LOCAL        0x000000ff

#define DTE_IS_REMOTE_VALID(v) (!!((v) & DTE_FLAG_REMOTE_VALID))
#define DTE_IS_ACL_CAPABLE(v)  (!!((v) & DTE_FLAG_ACL_CAPABLE))
#define DTE_REMOTE(v)          (((v) & DTE_MASK_REMOTE) >> 8)
#define DTE_LOCAL(v)           ((v) & DTE_MASK_LOCAL)
 
#define DT_ROOM	        ((ULONG) 0x00000007)
#define DT_EQUIPMENT    ((ULONG) 0x00000008)
#define DT_SEC_DISTLIST ((ULONG) 0x00000009)

// http://msdn2.microsoft.com/en-us/library/bb821036.aspx
#define PR_FLAG_STATUS PROP_TAG( PT_LONG, 0x1090 )
enum FollowUpStatus {
	flwupNone = 0,
	flwupComplete,
	flwupMarked,
	flwupMAX};

// http://msdn2.microsoft.com/en-us/library/bb821062.aspx
#define PR_FOLLOWUP_ICON PROP_TAG( PT_LONG, 0x1095 )
typedef enum OlFlagIcon {
	olNoFlagIcon=0,
	olPurpleFlagIcon=1,
	olOrangeFlagIcon=2,
	olGreenFlagIcon=3,
	olYellowFlagIcon=4,
	olBlueFlagIcon=5,
	olRedFlagIcon=6,
} OlFlagIcon;

// http://msdn2.microsoft.com/en-us/library/bb821130.aspx
enum Gender {
	genderMin = 0,
	genderUnspecified = genderMin,
	genderFemale,
	genderMale,
	genderCount,
	genderMax = genderCount - 1
};