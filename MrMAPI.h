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
	cmdmodeChildFolders,
	cmdmodeFidMid,
};

struct MYOPTIONS
{
	CmdMode Mode;
	bool bHelp;
	bool bDoPartialSearch;
	bool bDoType;
	bool bDoDispid;
	bool bDoDecimal;
	bool bBinaryFile;
	bool bDoContents;
	bool bDoAssociatedContents;
	bool bRetryStreamProps;
	bool bVerbose;
	bool bNoAddins;
	bool bOnline;
	bool bMSG;
	bool bList;
	LPWSTR lpszUnswitchedOption;
	LPWSTR lpszProfile;
	ULONG ulTypeNum;
	ULONG ulParser;
	LPWSTR lpszInput;
	LPWSTR lpszOutput;
	LPWSTR lpszSubject;
	LPWSTR lpszMessageClass;
	LPWSTR lpszFolderPath;
	LPWSTR lpszFid;
	LPWSTR lpszMid;
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
_Check_return_ HRESULT MrMAPILogonEx(_In_opt_z_ LPCWSTR lpszProfile, _Deref_out_opt_ LPMAPISESSION* lppSession);
_Check_return_ HRESULT OpenExchangeOrDefaultMessageStore(
							 _In_ LPMAPISESSION lpMAPISession,
							 _Deref_out_opt_ LPMDB* lppMDB);

////////////////////////////////////////////////////////////////////////////////
// Command line interface functions
////////////////////////////////////////////////////////////////////////////////

void DisplayUsage(BOOL bFull);
