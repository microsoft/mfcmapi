#pragma once
#include <core/utility/cli.h>

// MrMAPI command line
namespace cli
{
	extern option switchSearch;
	extern option switchDecimal;
	extern option switchFolder;
	extern option switchOutput;
	extern option switchDispid;
	extern option switchType;
	extern option switchGuid;
	extern option switchError;
	extern option switchParser;
	extern option switchInput;
	extern option switchBinary;
	extern option switchAcl;
	extern option switchRule;
	extern option switchContents;
	extern option switchAssociatedContents;
	extern option switchMoreProperties;
	extern option switchNoAddins;
	extern option switchOnline;
	extern option switchMAPI;
	extern option switchMIME;
	extern option switchCCSFFlags;
	extern option switchRFC822;
	extern option switchWrap;
	extern option switchEncoding;
	extern option switchCharset;
	extern option switchAddressBook;
	extern option switchUnicode;
	extern option switchProfile;
	extern option switchXML;
	extern option switchSubject;
	extern option switchMessageClass;
	extern option switchMSG;
	extern option switchList;
	extern option switchChildFolders;
	extern option switchFid;
	extern option switchMid;
	extern option switchFlag;
	extern option switchRecent;
	extern option switchStore;
	extern option switchVersion;
	extern option switchSize;
	extern option switchPST;
	extern option switchProfileSection;
	extern option switchByteSwapped;
	extern option switchReceiveFolder;
	extern option switchSkip;
	extern option switchSearchState;
	extern std::vector<option*> g_options;

#define ulNoMatch 0xffffffff

	enum CmdMode
	{
		cmdmodePropTag = cmdmodeFirstMode,
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
		// Declared in flagsEnum
		//OPT_NOOPT = 0x00000,
		//OPT_VERBOSE = 0x00001,
		//OPT_INITMFC = 0x00002,
		OPT_DOPARTIALSEARCH = 0x00004,
		OPT_DOTYPE = 0x00008,
		OPT_DODISPID = 0x00010,
		OPT_DODECIMAL = 0x00020,
		OPT_BINARYFILE = 0x00080,
		OPT_DOCONTENTS = 0x00100,
		OPT_DOASSOCIATEDCONTENTS = 0x00200,
		OPT_RETRYSTREAMPROPS = 0x00400,
		OPT_NOADDINS = 0x00800,
		OPT_ONLINE = 0x01000,
		OPT_MSG = 0x02000,
		OPT_LIST = 0x04000,
		OPT_NEEDMAPIINIT = 0x08000,
		OPT_NEEDMAPILOGON = 0x10000,
		OPT_NEEDFOLDER = 0x20000,
		OPT_NEEDINPUTFILE = 0x40000,
		OPT_NEEDOUTPUTFILE = 0x80000,
		OPT_PROFILE = 0x100000,
		OPT_NEEDSTORE = 0x200000,
		OPT_SKIPATTACHMENTS = 0x400000,
	};
	inline OPTIONFLAGS& operator|=(OPTIONFLAGS& a, OPTIONFLAGS b) noexcept
	{
		return reinterpret_cast<OPTIONFLAGS&>(reinterpret_cast<int&>(a) |= static_cast<int>(b));
	}
	inline OPTIONFLAGS operator|(OPTIONFLAGS a, OPTIONFLAGS b) noexcept
	{
		return static_cast<OPTIONFLAGS>(static_cast<int>(a) | static_cast<int>(b));
	}

	void DisplayUsage(BOOL bFull);
	void PostParseCheck(OPTIONS* _options);
} // namespace cli