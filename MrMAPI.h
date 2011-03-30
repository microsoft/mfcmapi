#pragma once
// MrMAPI.h : MrMAPI command line

#define ulNoMatch 0xffffffff

enum CmdMode {
	cmdmodeUnknown = 0,
	cmdmodePropTag,
	cmdmodeGuid,
	cmdmodeSmartView,
	cmdmodeAcls,
	cmdmodeRules,
	cmdmodeErr,
	cmdmodeContents,
	cmdmodeMAPIMIME,
	cmdmodeXML,
};

struct MYOPTIONS
{
	CmdMode Mode;
	BOOL  bDoPartialSearch;
	BOOL  bDoType;
	BOOL  bDoDispid;
	BOOL  bDoDecimal;
	BOOL  bBinaryFile;
	BOOL  bDoContents;
	BOOL  bDoAssociatedContents;
	BOOL  bRetryStreamProps;
	BOOL  bVerbose;
	BOOL  bNoAddins;
	BOOL  bOnline;
	LPWSTR lpszUnswitchedOption;
	LPWSTR lpszProfile;
	ULONG ulTypeNum;
	ULONG ulParser;
	LPWSTR lpszInput;
	LPWSTR lpszOutput;
	ULONG ulFolder;
	ULONG ulMAPIMIMEFlags;
	ULONG ulConvertFlags;
	ULONG ulWrapLines;
	ULONG ulEncodingType;
	ULONG ulCodePage;
	CHARSETTYPE cSetType;
	CSETAPPLYTYPE cSetApplyType;
};

void InitMFC();
HRESULT MrMAPILogonEx(LPCWSTR lpszProfile, LPMAPISESSION FAR* lppSession);

////////////////////////////////////////////////////////////////////////////////
// Command line interface functions
////////////////////////////////////////////////////////////////////////////////

void DisplayUsage();
