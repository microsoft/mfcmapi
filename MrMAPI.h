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
};

struct MYOPTIONS
{
	CmdMode Mode;
	BOOL  bDoPartialSearch;
	BOOL  bDoType;
	BOOL  bDoDispid;
	BOOL  bDoDecimal;
	BOOL  bBinaryFile;
	LPWSTR lpszUnswitchedOption;
	ULONG ulTypeNum;
	ULONG ulParser;
	LPWSTR lpszInput;
	LPWSTR lpszOutput;
	LPWSTR lpszProfile;
	ULONG ulFolder;
};

void InitMFC();

////////////////////////////////////////////////////////////////////////////////
// Command line interface functions
////////////////////////////////////////////////////////////////////////////////

void DisplayUsage();
