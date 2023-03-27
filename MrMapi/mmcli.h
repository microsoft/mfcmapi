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
	extern option switchAccounts;
	extern option switchIterate;
	extern option switchWizard;
	extern option switchFindProperty;
	extern option switchFindNamedProperty;

	extern std::vector<option*> g_options;

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
		cmdmodeNamedProps,
		cmdmodeEnumAccounts,
	};

	enum OPTIONFLAGS
	{
		// Declared in flagsEnum
		//OPT_NOOPT = 0x0000,
		//OPT_INITMFC = 0x0001,
		//OPT_NEEDNUM = 0x0002, // Any arguments must be decimal numbers. No strings.
		OPT_NEEDMAPIINIT = 0x0010,
		OPT_NEEDMAPILOGON = 0x0020,
		OPT_NEEDFOLDER = 0x0040,
		OPT_NEEDINPUTFILE = 0x0080,
		OPT_NEEDOUTPUTFILE = 0x0100,
		OPT_NEEDSTORE = 0x0200,
		OPT_PROFILE = 0x0400,
	};

	void DisplayUsage(BOOL bFull) noexcept;
	void PostParseCheck(OPTIONS& _options);
} // namespace cli