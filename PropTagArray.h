#pragma once

extern NAME_ARRAY_ENTRY g_PropTagArray[];
extern ULONG g_ulPropTagArray;

extern NAME_ARRAY_ENTRY g_PropTypeArray[];
extern ULONG g_ulPropTypeArray;

extern GUID_ARRAY_ENTRY g_PropGuidArray[];
extern ULONG g_ulPropGuidArray;

extern NAMEID_ARRAY_ENTRY g_NameIDArray[];
extern ULONG g_ulNameIDArray;

extern FLAG_ARRAY_ENTRY g_FlagArray[];
extern ULONG g_ulFlagArray;

enum __NonPropFlag
{
	flagSearchFlag = 0x10000,//ensure that all flags in the enum are > 0xffff
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
	flagAccountType,
	flagBounceCode,
	flagOPReply,
	flagOpForward,
	flagFuzzyLevel,
	flagRulesVersion,
	flagNotifEventType,
	flagTableEventType,
};