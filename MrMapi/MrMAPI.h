#pragma once
// MrMAPI command line

#define ulNoMatch 0xffffffff
#include "Interpret/ExtraPropTags.h"

enum CmdMode
{
	cmdmodeUnknown = 0,
	cmdmodeHelp,
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
	cmdmodeStoreProperties,
	cmdmodeFlagSearch,
	cmdmodeFolderProps,
	cmdmodeFolderSize,
	cmdmodePST,
	cmdmodeProfile,
	cmdmodeReceiveFolder,
	cmdmodeSearchState,
};

#define OPT_DOPARTIALSEARCH 0x00001
#define OPT_DOTYPE 0x00002
#define OPT_DODISPID 0x00004
#define OPT_DODECIMAL 0x00008
#define OPT_DOFLAG 0x00010
#define OPT_BINARYFILE 0x00020
#define OPT_DOCONTENTS 0x00040
#define OPT_DOASSOCIATEDCONTENTS 0x00080
#define OPT_RETRYSTREAMPROPS 0x00100
#define OPT_VERBOSE 0x00200
#define OPT_NOADDINS 0x00400
#define OPT_ONLINE 0x00800
#define OPT_MSG 0x01000
#define OPT_LIST 0x02000
#define OPT_NEEDMAPIINIT 0x04000
#define OPT_NEEDMAPILOGON 0x08000
#define OPT_INITMFC 0x10000
#define OPT_NEEDFOLDER 0x20000
#define OPT_NEEDINPUTFILE 0x40000
#define OPT_NEEDOUTPUTFILE 0x80000
#define OPT_PROFILE 0x100000
#define OPT_NEEDSTORE 0x200000
#define OPT_SKIPATTACHMENTS 0x400000
#define OPT_MID 0x800000

struct MYOPTIONS
{
	CmdMode Mode;
	ULONG ulOptions;
	std::wstring lpszUnswitchedOption;
	std::wstring lpszProfile;
	ULONG ulTypeNum;
	ULONG ulSVParser;
	std::wstring lpszInput;
	std::wstring lpszOutput;
	std::wstring lpszSubject;
	std::wstring lpszMessageClass;
	std::wstring lpszFolderPath;
	std::wstring lpszFid;
	std::wstring lpszMid;
	std::wstring lpszFlagName;
	std::wstring lpszVersion;
	std::wstring lpszProfileSection;
	ULONG ulStore;
	ULONG ulFolder;
	ULONG ulMAPIMIMEFlags;
	CCSFLAGS convertFlags;
	ULONG ulWrapLines;
	ULONG ulEncodingType;
	ULONG ulCodePage;
	ULONG ulFlagValue;
	ULONG ulCount;
	bool bByteSwapped;
	CHARSETTYPE cSetType;
	CSETAPPLYTYPE cSetApplyType;
	LPMAPISESSION lpMAPISession;
	LPMDB lpMDB;
	LPMAPIFOLDER lpFolder;

	MYOPTIONS();
};

_Check_return_ LPMAPISESSION MrMAPILogonEx(const std::wstring& lpszProfile);
_Check_return_ LPMDB OpenExchangeOrDefaultMessageStore(_In_ LPMAPISESSION lpMAPISession);

////////////////////////////////////////////////////////////////////////////////
// Command line interface functions
////////////////////////////////////////////////////////////////////////////////

void DisplayUsage(BOOL bFull);
