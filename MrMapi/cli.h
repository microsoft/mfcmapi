#pragma once
// MrMAPI command line

#define ulNoMatch 0xffffffff
#include <Interpret/ExtraPropTags.h>
#include <MrMapi/MMMapiMime.h>

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

enum OPTIONFLAGS
{
	OPT_NOOPT = 0x00000,
	OPT_DOPARTIALSEARCH = 0x00001,
	OPT_DOTYPE = 0x00002,
	OPT_DODISPID = 0x00004,
	OPT_DODECIMAL = 0x00008,
	OPT_DOFLAG = 0x00010,
	OPT_BINARYFILE = 0x00020,
	OPT_DOCONTENTS = 0x00040,
	OPT_DOASSOCIATEDCONTENTS = 0x00080,
	OPT_RETRYSTREAMPROPS = 0x00100,
	OPT_VERBOSE = 0x00200,
	OPT_NOADDINS = 0x00400,
	OPT_ONLINE = 0x00800,
	OPT_MSG = 0x01000,
	OPT_LIST = 0x02000,
	OPT_NEEDMAPIINIT = 0x04000,
	OPT_NEEDMAPILOGON = 0x08000,
	OPT_INITMFC = 0x10000,
	OPT_NEEDFOLDER = 0x20000,
	OPT_NEEDINPUTFILE = 0x40000,
	OPT_NEEDOUTPUTFILE = 0x80000,
	OPT_PROFILE = 0x100000,
	OPT_NEEDSTORE = 0x200000,
	OPT_SKIPATTACHMENTS = 0x400000,
	OPT_MID = 0x800000,
};
inline OPTIONFLAGS& operator|=(OPTIONFLAGS& a, OPTIONFLAGS b)
{
	return reinterpret_cast<OPTIONFLAGS&>(reinterpret_cast<int&>(a) |= static_cast<int>(b));
}
inline OPTIONFLAGS operator|(OPTIONFLAGS a, OPTIONFLAGS b)
{
	return static_cast<OPTIONFLAGS>(static_cast<int>(a) | static_cast<int>(b));
}

struct MYOPTIONS
{
	CmdMode Mode;
	OPTIONFLAGS ulOptions;
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
	MAPIMIMEFLAGS MAPIMIMEFlags;
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

////////////////////////////////////////////////////////////////////////////////
// Command line interface functions
////////////////////////////////////////////////////////////////////////////////

void DisplayUsage(BOOL bFull);

// Parses command line arguments and fills out MYOPTIONS
bool ParseArgs(_In_ int argc, _In_count_(argc) char* argv[], _Out_ MYOPTIONS* pRunOpts);
void PrintArgs(_In_ MYOPTIONS ProgOpts);