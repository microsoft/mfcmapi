#pragma once
// MrMAPI.h : MrMAPI command line

#define ulNoMatch 0xffffffff

enum CmdMode
{
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
	bool  bHelp;
	bool  bDoPartialSearch;
	bool  bDoType;
	bool  bDoDispid;
	bool  bDoDecimal;
	bool  bBinaryFile;
	bool  bDoContents;
	bool  bDoAssociatedContents;
	bool  bRetryStreamProps;
	bool  bVerbose;
	bool  bNoAddins;
	bool  bOnline;
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
HRESULT MrMAPILogonEx(_In_opt_z_ LPCWSTR lpszProfile, _Deref_out_opt_ LPMAPISESSION* lppSession);

////////////////////////////////////////////////////////////////////////////////
// Command line interface functions
////////////////////////////////////////////////////////////////////////////////

void DisplayUsage(BOOL bFull);
