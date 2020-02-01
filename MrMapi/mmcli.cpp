#include <StdAfx.h>
#include <MrMapi/mmcli.h>
#include <core/utility/cli.h>
#include <core/mapi/mapiFunctions.h>
#include <core/addin/addin.h>
#include <core/addin/mfcmapi.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/interpret/proptype.h>

namespace error
{
	extern ULONG g_ulErrorArray;
} // namespace error

#define OPT_INITALL (OPT_INITMFC | OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON)
#define OPT_INOUT (OPT_NEEDINPUTFILE | OPT_NEEDOUTPUTFILE)
namespace cli
{
#pragma warning(push)
#pragma warning(disable : 5054) // warning C5054: operator '|': deprecated between enumerations of different types
	option switchSearch{L"Search", cmdmodeUnknown, 0, 0, OPT_NOOPT};
	option switchDecimal{L"Number", cmdmodeUnknown, 0, 0, OPT_NOOPT};
	option switchFolder{L"Folder", cmdmodeUnknown, 1, 1, OPT_INITALL | OPT_NEEDFOLDER};
	option switchOutput{L"Output", cmdmodeUnknown, 1, 1, OPT_NOOPT};
	option switchDispid{L"Dispids", cmdmodePropTag, 0, 0, OPT_NOOPT};
	option switchType{L"Type", cmdmodePropTag, 0, 1, OPT_NOOPT};
	option switchGuid{L"Guids", cmdmodeGuid, 0, 0, OPT_NOOPT};
	option switchError{L"Error", cmdmodeErr, 0, 0, OPT_NOOPT};
	option switchParser{L"ParserType", cmdmodeSmartView, 1, 1, OPT_INITMFC | OPT_NEEDINPUTFILE};
	option switchInput{L"Input", cmdmodeUnknown, 1, 1, OPT_NOOPT};
	option switchBinary{L"Binary", cmdmodeSmartView, 0, 0, OPT_NOOPT};
	option switchAcl{L"Acl", cmdmodeAcls, 0, 0, OPT_INITALL | OPT_NEEDFOLDER};
	option switchRule{L"Rules", cmdmodeRules, 0, 0, OPT_INITALL | OPT_NEEDFOLDER};
	option switchContents{L"Contents", cmdmodeContents, 0, 0, OPT_INITALL};
	option switchAssociatedContents{L"HiddenContents", cmdmodeContents, 0, 0, OPT_INITALL};
	option switchMoreProperties{L"MoreProperties", cmdmodeUnknown, 0, 0, OPT_NOOPT};
	option switchNoAddins{L"NoAddins", cmdmodeUnknown, 0, 0, OPT_NOOPT};
	option switchOnline{L"Online", cmdmodeUnknown, 0, 0, OPT_NOOPT};
	option switchMAPI{L"MAPI", cmdmodeMAPIMIME, 0, 0, OPT_NEEDMAPIINIT | OPT_INITMFC | OPT_INOUT};
	option switchMIME{L"MIME", cmdmodeMAPIMIME, 0, 0, OPT_NEEDMAPIINIT | OPT_INITMFC | OPT_INOUT};
	option switchCCSFFlags{L"CCSFFlags", cmdmodeMAPIMIME, 1, 1, OPT_NOOPT};
	option switchRFC822{L"RFC822", cmdmodeMAPIMIME, 0, 0, OPT_NOOPT};
	option switchWrap{L"Wrap", cmdmodeMAPIMIME, 1, 1, OPT_NOOPT};
	option switchEncoding{L"Encoding", cmdmodeMAPIMIME, 1, 1, OPT_NOOPT};
	option switchCharset{L"Charset", cmdmodeMAPIMIME, 3, 3, OPT_NOOPT};
	option switchAddressBook{L"AddressBook", cmdmodeMAPIMIME, 0, 0, OPT_NEEDMAPILOGON};
	option switchUnicode{L"Unicode", cmdmodeMAPIMIME, 0, 0, OPT_NOOPT};
	option switchProfile{L"Profile", cmdmodeUnknown, 0, 1, OPT_PROFILE};
	option switchXML{L"XML", cmdmodeXML, 0, 0, OPT_NEEDMAPIINIT | OPT_INITMFC | OPT_NEEDINPUTFILE};
	option switchSubject{L"Subject", cmdmodeContents, 1, 1, OPT_NOOPT};
	option switchMessageClass{L"MessageClass", cmdmodeContents, 1, 1, OPT_NOOPT};
	option switchMSG{L"MSG", cmdmodeContents, 0, 0, OPT_NOOPT};
	option switchList{L"List", cmdmodeContents, 0, 0, OPT_NOOPT};
	option switchChildFolders{L"ChildFolders", cmdmodeChildFolders, 0, 0, OPT_INITALL | OPT_NEEDFOLDER};
	option switchFid{L"FID", cmdmodeFidMid, 0, 1, OPT_INITALL | OPT_NEEDSTORE};
	option switchMid{L"MID", cmdmodeFidMid, 0, 1, OPT_INITALL};
	option switchFlag{L"Flag", cmdmodeUnknown, 1, 1, OPT_NOOPT};
	option switchRecent{L"Recent", cmdmodeContents, 1, 1, OPT_NOOPT};
	option switchStore{L"Store", cmdmodeStoreProperties, 0, 1, OPT_INITALL | OPT_NEEDNUM};
	option switchVersion{L"Version", cmdmodeUnknown, 1, 1, OPT_NOOPT};
	option switchSize{L"Size", cmdmodeFolderSize, 0, 0, OPT_INITALL | OPT_NEEDFOLDER};
	option switchPST{L"PST", cmdmodePST, 0, 0, OPT_NEEDINPUTFILE};
	option switchProfileSection{L"ProfileSection", cmdmodeProfile, 1, 1, OPT_PROFILE | OPT_NEEDMAPIINIT | OPT_INITMFC};
	option switchByteSwapped{L"ByteSwapped", cmdmodeProfile, 0, 0, OPT_PROFILE | OPT_NEEDMAPIINIT | OPT_INITMFC};
	option switchReceiveFolder{L"ReceiveFolder", cmdmodeReceiveFolder, 0, 1, OPT_INITALL | OPT_NEEDSTORE | OPT_NEEDNUM};
	option switchSkip{L"Skip", cmdmodeUnknown, 0, 0, OPT_NOOPT};
	option switchSearchState{L"SearchState", cmdmodeSearchState, 0, 1, OPT_INITALL | OPT_NEEDFOLDER};

	// If we want to add aliases for any switches, add them here
	option switchHelpAlias{L"Help", cmdmodeHelpFull, 0, 0, OPT_INITMFC};
#pragma warning(pop)

	std::vector<option*> g_options = {
		&switchHelp,
		&switchVerbose,
		&switchSearch,
		&switchDecimal,
		&switchFolder,
		&switchOutput,
		&switchDispid,
		&switchType,
		&switchGuid,
		&switchError,
		&switchParser,
		&switchInput,
		&switchBinary,
		&switchAcl,
		&switchRule,
		&switchContents,
		&switchAssociatedContents,
		&switchMoreProperties,
		&switchNoAddins,
		&switchOnline,
		&switchMAPI,
		&switchMIME,
		&switchCCSFFlags,
		&switchRFC822,
		&switchWrap,
		&switchEncoding,
		&switchCharset,
		&switchAddressBook,
		&switchUnicode,
		&switchProfile,
		&switchXML,
		&switchSubject,
		&switchMessageClass,
		&switchMSG,
		&switchList,
		&switchChildFolders,
		&switchFid,
		&switchMid,
		&switchFlag,
		&switchRecent,
		&switchStore,
		&switchVersion,
		&switchSize,
		&switchPST,
		&switchProfileSection,
		&switchByteSwapped,
		&switchReceiveFolder,
		&switchSkip,
		&switchSearchState,
		// If we want to add aliases for any switches, add them here
		&switchHelpAlias,
	};

	void DisplayUsage(BOOL bFull)
	{
		printf("MAPI data collection and parsing tool. Supports property tag lookup, error translation,\n");
		printf("   smart view processing, rule tables, ACL tables, contents tables, and MAPI<->MIME conversion.\n");
		if (bFull)
		{
			if (!g_lpMyAddins.empty())
			{
				printf("Addins Loaded:\n");
				for (const auto& addIn : g_lpMyAddins)
				{
					printf("   %ws\n", addIn.szName);
				}
			}

			printf("MrMAPI currently knows:\n");
			printf("%6u property tags\n", static_cast<int>(PropTagArray.size()));
			printf("%6u dispids\n", static_cast<int>(NameIDArray.size()));
			printf("%6u types\n", static_cast<int>(PropTypeArray.size()));
			printf("%6u guids\n", static_cast<int>(PropGuidArray.size()));
			printf("%6lu errors\n", error::g_ulErrorArray);
			printf("%6u smart view parsers\n", static_cast<int>(SmartViewParserTypeArray.size()) - 1);
			printf("\n");
		}

		printf("Usage:\n");
		printf("   MrMAPI -%ws\n", switchHelp.name());
		printf(
			"   MrMAPI [-%ws] [-%ws] [-%ws] [-%ws <type>] <property number>|<property name>\n",
			switchSearch.name(),
			switchDispid.name(),
			switchDecimal.name(),
			switchType.name());
		printf("   MrMAPI -%ws\n", switchGuid.name());
		printf("   MrMAPI -%ws <error>\n", switchError.name());
		printf(
			"   MrMAPI -%ws <type> -%ws <input file> [-%ws] [-%ws <output file>]\n",
			switchParser.name(),
			switchInput.name(),
			switchBinary.name(),
			switchOutput.name());
		printf(
			"   MrMAPI -%ws <flag value> [-%ws] [-%ws] <property number>|<property name>\n",
			switchFlag.name(),
			switchDispid.name(),
			switchDecimal.name());
		printf("   MrMAPI -%ws <flag name>\n", switchFlag.name());
		printf(
			"   MrMAPI -%ws [-%ws <profile>] [-%ws <folder>]\n",
			switchRule.name(),
			switchProfile.name(),
			switchFolder.name());
		printf(
			"   MrMAPI -%ws [-%ws <profile>] [-%ws <folder>]\n",
			switchAcl.name(),
			switchProfile.name(),
			switchFolder.name());
		printf(
			"   MrMAPI -%ws | -%ws [-%ws <profile>] [-%ws <folder>] [-%ws <output directory>]\n",
			switchContents.name(),
			switchAssociatedContents.name(),
			switchProfile.name(),
			switchFolder.name(),
			switchOutput.name());
		printf(
			"          [-%ws <subject>] [-%ws <message class>] [-%ws] [-%ws] [-%ws <count>] [-%ws]\n",
			switchSubject.name(),
			switchMessageClass.name(),
			switchMSG.name(),
			switchList.name(),
			switchRecent.name(),
			switchSkip.name());
		printf(
			"   MrMAPI -%ws [-%ws <profile>] [-%ws <folder>]\n",
			switchChildFolders.name(),
			switchProfile.name(),
			switchFolder.name());
		printf(
			"   MrMAPI -%ws -%ws <path to input file> -%ws <path to output file> [-%ws]\n",
			switchXML.name(),
			switchInput.name(),
			switchOutput.name(),
			switchSkip.name());
		printf(
			"   MrMAPI -%ws [fid] [-%ws [mid]] [-%ws <profile>]\n",
			switchFid.name(),
			switchMid.name(),
			switchProfile.name());
		printf(
			"   MrMAPI [<property number>|<property name>] -%ws [<store num>] [-%ws <profile>]\n",
			switchStore.name(),
			switchProfile.name());
		printf(
			"   MrMAPI [<property number>|<property name>] -%ws <folder> [-%ws <profile>]\n",
			switchFolder.name(),
			switchProfile.name());
		printf(
			"   MrMAPI -%ws -%ws <folder> [-%ws <profile>]\n",
			switchSize.name(),
			switchFolder.name(),
			switchProfile.name());
		printf(
			"   MrMAPI -%ws | -%ws -%ws <path to input file> -%ws <path to output file> [-%ws <conversion flags>]\n",
			switchMAPI.name(),
			switchMIME.name(),
			switchInput.name(),
			switchOutput.name(),
			switchCCSFFlags.name());
		printf(
			"          [-%ws] [-%ws <Decimal number of characters>] [-%ws <Decimal number indicating encoding>]\n",
			switchRFC822.name(),
			switchWrap.name(),
			switchEncoding.name());
		printf(
			"          [-%ws] [-%ws] [-%ws CodePage CharSetType CharSetApplyType]\n",
			switchAddressBook.name(),
			switchUnicode.name(),
			switchCharset.name());
		printf("   MrMAPI -%ws -%ws <path to input file>\n", switchPST.name(), switchInput.name());
		printf(
			"   MrMAPI -%ws [<profile> [-%ws <profilesection> [-%ws]] -%ws <output file>]\n",
			switchProfile.name(),
			switchProfileSection.name(),
			switchByteSwapped.name(),
			switchOutput.name());
		printf("   MrMAPI -%ws [<store num>] [-%ws <profile>]\n", switchReceiveFolder.name(), switchProfile.name());
		printf(
			"   MrMAPI -%ws -%ws <folder> [-%ws <profile>]\n",
			switchSearchState.name(),
			switchFolder.name(),
			switchProfile.name());

		if (bFull)
		{
			printf("\n");
			printf("All switches may be shortened if the intended switch is unambiguous.\n");
			printf("For example, -T may be used instead of -%ws.\n", switchType.name());
		}
		printf("\n");
		printf("   Help:\n");
		printf("   -%ws   Display expanded help.\n", switchHelp.name());
		if (bFull)
		{
			printf("\n");
			printf("   Property Tag Lookup:\n");
			printf("   -S   (or -%ws) Perform substring search.\n", switchSearch.name());
			printf("           With no parameters prints all known properties.\n");
			printf("   -D   (or -%ws) Search dispids.\n", switchDispid.name());
			printf("   -N   (or -%ws) Number is in decimal. Ignored for non-numbers.\n", switchDecimal.name());
			printf("   -T   (or -%ws) Print information on specified type.\n", switchType.name());
			printf("           With no parameters prints list of known types.\n");
			printf("           When combined with -S, restrict output to given type.\n");
			printf("   -G   (or -%ws) Display list of known guids.\n", switchGuid.name());
			printf("\n");
			printf("   Flag Lookup:\n");
			printf("   -Fl  (or -%ws) Look up flags for specified property.\n", switchFlag.name());
			printf("           May be combined with -D and -N switches, but all flag values must be in hex.\n");
			printf("   -Fl  (or -%ws) Look up flag name and output its value.\n", switchFlag.name());
			printf("\n");
			printf("   Error Parsing:\n");
			printf("   -E   (or -%ws) Map an error code to its name and vice versa.\n", switchError.name());
			printf("           May be combined with -S and -N switches.\n");
			printf("\n");
			printf("   Smart View Parsing:\n");
			printf(
				"   -P   (or -%ws) Parser type (number). See list below for supported parsers.\n", switchParser.name());
			printf("   -B   (or -%ws) Input file is binary. Default is hex encoded text.\n", switchBinary.name());
			printf("\n");
			printf("   Rules Table:\n");
			printf("   -R   (or -%ws) Output rules table. Profile optional.\n", switchRule.name());
			printf("\n");
			printf("   ACL Table:\n");
			printf("   -A   (or -%ws) Output ACL table. Profile optional.\n", switchAcl.name());
			printf("\n");
			printf("   Contents Table:\n");
			printf(
				"   -C   (or -%ws) Output contents table. May be combined with -H. Profile optional.\n",
				switchContents.name());
			printf(
				"   -H   (or -%ws) Output associated contents table. May be combined with -C. Profile optional\n",
				switchAssociatedContents.name());
			printf("   -Su  (or -%ws) Subject of messages to output.\n", switchSubject.name());
			printf("   -Me  (or -%ws) Message class of messages to output.\n", switchMessageClass.name());
			printf("   -Ms  (or -%ws) Output as .MSG instead of XML.\n", switchMSG.name());
			printf("   -L   (or -%ws) List details to screen and do not output files.\n", switchList.name());
			printf("   -Re  (or -%ws) Restrict output to the 'count' most recent messages.\n", switchRecent.name());
			printf("\n");
			printf("   Child Folders:\n");
			printf("   -Chi (or -%ws) Display child folders of selected folder.\n", switchChildFolders.name());
			printf("\n");
			printf("   MSG File Properties\n");
			printf("   -X   (or -%ws) Output properties of an MSG file as XML.\n", switchXML.name());
			printf("\n");
			printf("   MID/FID Lookup\n");
			printf("   -Fi  (or -%ws) Folder ID (FID) to search for.\n", switchFid.name());
			printf("           If -%ws is specified without a FID, search/display all folders\n", switchFid.name());
			printf("   -Mid (or -%ws) Message ID (MID) to search for.\n", switchMid.name());
			printf(
				"           If -%ws is specified without a MID, display all messages in folders specified by the FID "
				"parameter.\n",
				switchMid.name());
			printf("\n");
			printf("   Store Properties\n");
			printf("   -St  (or -%ws) Output properties of stores as XML.\n", switchStore.name());
			printf("           If store number is specified, outputs properties of a single store.\n");
			printf("           If a property is specified, outputs only that property.\n");
			printf("\n");
			printf("   Folder Properties\n");
			printf("   -F   (or -%ws) Output properties of a folder as XML.\n", switchFolder.name());
			printf("           If a property is specified, outputs only that property.\n");
			printf("   -Size         Output size of a folder and all subfolders.\n");
			printf("           Use -%ws to specify which folder to scan.\n", switchFolder.name());
			printf("   -SearchState  Output search folder state.\n");
			printf("           Use -%ws to specify which folder to scan.\n", switchFolder.name());
			printf("\n");
			printf("   MAPI <-> MIME Conversion:\n");
			printf("   -Ma  (or -%ws) Convert an EML file to MAPI format (MSG file).\n", switchMAPI.name());
			printf("   -Mi  (or -%ws) Convert an MSG file to MIME format (EML file).\n", switchMIME.name());
			printf(
				"   -I   (or -%ws) Indicates the input file for conversion, either a MIME-formatted EML file or an MSG "
				"file.\n",
				switchInput.name());
			printf("   -O   (or -%ws) Indicates the output file for the conversion.\n", switchOutput.name());
			printf("   -Cc  (or -%ws) Indicates specific flags to pass to the converter.\n", switchCCSFFlags.name());
			printf("           Available values (these may be OR'ed together):\n");
			printf("              MIME -> MAPI:\n");
			printf("                CCSF_SMTP:        0x02\n");
			printf("                CCSF_INCLUDE_BCC: 0x20\n");
			printf("                CCSF_USE_RTF:     0x80\n");
			printf("              MAPI -> MIME:\n");
			printf("                CCSF_NOHEADERS:        0x000004\n");
			printf("                CCSF_USE_TNEF:         0x000010\n");
			printf("                CCSF_8BITHEADERS:      0x000040\n");
			printf("                CCSF_PLAIN_TEXT_ONLY:  0x001000\n");
			printf("                CCSF_NO_MSGID:         0x004000\n");
			printf("                CCSF_EMBEDDED_MESSAGE: 0x008000\n");
			printf("                CCSF_PRESERVE_SOURCE:  0x040000\n");
			printf("                CCSF_GLOBAL_MESSAGE:   0x200000\n");
			printf(
				"   -Rf  (or -%ws) (MAPI->MIME only) Indicates the EML should be generated in RFC822 format.\n",
				switchRFC822.name());
			printf("           If not present, RFC1521 is used instead.\n");
			printf(
				"   -W   (or -%ws) (MAPI->MIME only) Indicates the maximum number of characters in each line in the\n",
				switchWrap.name());
			printf("           generated EML. Default value is 74. A value of 0 indicates no wrapping.\n");
			printf(
				"   -En  (or -%ws) (MAPI->MIME only) Indicates the encoding type to use. Supported values are:\n",
				switchEncoding.name());
			printf("              1 - Base64\n");
			printf("              2 - UUENCODE\n");
			printf("              3 - Quoted-Printable\n");
			printf("              4 - 7bit (DEFAULT)\n");
			printf("              5 - 8bit\n");
			printf(
				"   -Ad  (or -%ws) Pass MAPI Address Book into converter. Profile optional.\n",
				switchAddressBook.name());
			printf(
				"   -U   (or -%ws) (MIME->MAPI only) The resulting MSG file should be unicode.\n",
				switchUnicode.name());
			printf(
				"   -Ch  (or -%ws) (MIME->MAPI only) Character set - three required parameters:\n",
				switchCharset.name());
			printf("           CodePage - common values (others supported)\n");
			printf("              1252  - CP_USASCII      - Indicates the USASCII character set, Windows code page "
				   "1252\n");
			printf("              1200  - CP_UNICODE      - Indicates the Unicode character set, Windows code page "
				   "1200\n");
			printf("              50932 - CP_JAUTODETECT  - Indicates Japanese auto-detect (50932)\n");
			printf("              50949 - CP_KAUTODETECT  - Indicates Korean auto-detect (50949)\n");
			printf("              50221 - CP_ISO2022JPESC - Indicates the Internet character set ISO-2022-JP-ESC\n");
			printf("              50222 - CP_ISO2022JPSIO - Indicates the Internet character set ISO-2022-JP-SIO\n");
			printf("           CharSetType - supported values (see CHARSETTYPE)\n");
			printf("              0 - CHARSET_BODY\n");
			printf("              1 - CHARSET_HEADER\n");
			printf("              2 - CHARSET_WEB\n");
			printf("           CharSetApplyType - supported values (see CSETAPPLYTYPE)\n");
			printf("              0 - CSET_APPLY_UNTAGGED\n");
			printf("              1 - CSET_APPLY_ALL\n");
			printf("              2 - CSET_APPLY_TAG_ALL\n");
			printf("\n");
			printf("   PST Analysis\n");
			printf("   -PST Output statistics of a PST file.\n");
			printf("           If a property is specified, outputs only that property.\n");
			printf("   -I   (or -%ws) PST file to be analyzed.\n", switchInput.name());
			printf("\n");
			printf("   Profiles\n");
			printf("   -Pr  (or -%ws) Output list of profiles\n", switchProfile.name());
			printf("           If a profile is specified, exports that profile.\n");
			printf("   -ProfileSection If specified, output specific profile section.\n");
			printf("   -B   (or -%ws) If specified, profile section guid is byte swapped.\n", switchByteSwapped.name());
			printf("   -O   (or -%ws) Indicates the output file for profile export.\n", switchOutput.name());
			printf("           Required if a profile is specified.\n");
			printf("\n");
			printf("   Receive Folder Table\n");
			printf("   -%ws Displays Receive Folder Table for the specified store\n", switchReceiveFolder.name());
			printf("\n");
			printf("   Universal Options:\n");
			printf("   -I   (or -%ws) Input file.\n", switchInput.name());
			printf("   -O   (or -%ws) Output file or directory.\n", switchOutput.name());
			printf(
				"   -F   (or -%ws) Folder to scan. Default is Inbox. See list below for supported folders.\n",
				switchFolder.name());
			printf("           Folders may also be specified by path:\n");
			printf("              \"Top of Information Store\\Calendar\"\n");
			printf("           Path may be preceeded by entry IDs for special folders using @ notation:\n");
			printf("              \"@PR_IPM_SUBTREE_ENTRYID\\Calendar\"\n");
			printf("           Path may further be preceeded by store number using # notation, which may either use a "
				   "store number:\n");
			printf("              \"#0\\@PR_IPM_SUBTREE_ENTRYID\\Calendar\"\n");
			printf("           Or an entry ID:\n");
			printf("              \"#00112233445566...778899AABBCC\\@PR_IPM_SUBTREE_ENTRYID\\Calendar\"\n");
			printf("           MrMAPI's special folder constants may also be used:\n");
			printf("              \"@12\\Calendar\"\n");
			printf("              \"@1\"\n");
			printf("   -Pr  (or -%ws) Profile for MAPILogonEx.\n", switchProfile.name());
			printf(
				"   -M   (or -%ws) More properties. Tries harder to get stream properties. May take longer.\n",
				switchMoreProperties.name());
			printf("   -Sk  (or -%ws) Skip embedded message attachments on export.\n", switchSkip.name());
			printf("   -No  (or -%ws) No Addins. Don't load any add-ins.\n", switchNoAddins.name());
			printf("   -On  (or -%ws) Online mode. Bypass cached mode.\n", switchOnline.name());
			printf("   -V   (or -%ws) Verbose. Turn on all debug output.\n", switchVerbose.name());
			printf("\n");
			printf("   MAPI Implementation Options:\n");
			printf("   -%ws MAPI Version to load - supported values\n", switchVersion.name());
			printf("           Supported values\n");
			printf("              0  - List all available MAPI binaries\n");
			printf("              1  - System MAPI\n");
			printf("              11 - Outlook 2003 (11)\n");
			printf("              12 - Outlook 2007 (12)\n");
			printf("              14 - Outlook 2010 (14)\n");
			printf("              15 - Outlook 2013 (15)\n");
			printf("              16 - Outlook 2016 (16)\n");
			printf("           You can also pass a string, which will load the first MAPI whose path contains the "
				   "string.\n");
			printf("\n");
			printf("Smart View Parsers:\n");
			// Print smart view options
			for (ULONG i = 1; i < SmartViewParserTypeArray.size(); i++)
			{
				_tprintf(_T("   %2lu %ws\n"), i, SmartViewParserTypeArray[i].lpszName);
			}

			printf("\n");
			printf("Folders:\n");
			// Print Folders
			for (ULONG i = 1; i < mapi::NUM_DEFAULT_PROPS; i++)
			{
				printf("   %2lu %ws\n", i, mapi::FolderNames[i]);
			}

			printf("\n");
			printf("Examples:\n");
			printf("   MrMAPI PR_DISPLAY_NAME\n");
			printf("\n");
			printf("   MrMAPI 0x3001001e\n");
			printf("   MrMAPI 3001001e\n");
			printf("   MrMAPI 3001\n");
			printf("\n");
			printf("   MrMAPI -n 12289\n");
			printf("\n");
			printf("   MrMAPI -t PT_LONG\n");
			printf("   MrMAPI -t 3102\n");
			printf("   MrMAPI -t\n");
			printf("\n");
			printf("   MrMAPI -s display\n");
			printf("   MrMAPI -s display -t PT_LONG\n");
			printf("   MrMAPI -t 102 -s display\n");
			printf("\n");
			printf("   MrMAPI -d dispidReminderTime\n");
			printf("   MrMAPI -d 0x8502\n");
			printf("   MrMAPI -d -s reminder\n");
			printf("   MrMAPI -d -n 34050\n");
			printf("\n");
			printf("   MrMAPI -p 17 -i webview.txt -o parsed.txt");
		}
	}

	void PostParseCheck(OPTIONS& options)
	{
		// Having processed the command line, we may not have determined a mode.
		// Some modes can be presumed by the switches we did see.

		if (switchFlag.isSet())
		{
			if (!bSetMode(options.mode, switchFlag.hasULONG(0, 16) ? cmdmodePropTag : cmdmodeFlagSearch))
			{
				options.mode = cmdmodeHelp;
			}
		}

		// If we didn't get a mode set but we saw OPT_NEEDFOLDER, assume we're in folder dumping mode
		if (cmdmodeUnknown == options.mode && options.flags & OPT_NEEDFOLDER) options.mode = cmdmodeFolderProps;

		// If we didn't get a mode set, but we saw OPT_PROFILE, assume we're in profile dumping mode
		if (cmdmodeUnknown == options.mode && options.flags & OPT_PROFILE)
		{
			options.mode = cmdmodeProfile;
#pragma warning(push)
#pragma warning(disable : 5054) // warning C5054: operator '|': deprecated between enumerations of different types
			options.flags |= OPT_NEEDMAPIINIT | OPT_INITMFC;
#pragma warning(pop)
		}

		// If we didn't get a mode set, assume we're in prop tag mode
		if (cmdmodeUnknown == options.mode) options.mode = cmdmodePropTag;

		// If we weren't passed an output file/directory, remember the current directory
		if (switchOutput.empty() && options.mode != cmdmodeSmartView && options.mode != cmdmodeProfile)
		{
			WCHAR strPath[_MAX_PATH];
			GetCurrentDirectoryW(_MAX_PATH, strPath);

			// Trick switchOutput into scanning this path as an argument.
			auto args = std::deque<std::wstring>{strPath};
			(void) switchOutput.scanArgs(args, options, g_options);
		}

		if (switchFolder.empty())
		{
			// Trick switchFolder into scanning this path as an argument.
			OPTIONS fakeOptions{};
			auto args = std::deque<std::wstring>{L"13"};
			(void) switchFolder.scanArgs(args, fakeOptions, g_options);
		}

		// Validate that we have bare minimum to run
		if (options.flags & OPT_NEEDINPUTFILE && switchInput.empty())
			options.mode = cmdmodeHelp;
		else if (options.flags & OPT_NEEDOUTPUTFILE && switchOutput.empty())
			options.mode = cmdmodeHelp;

		switch (options.mode)
		{
		case cmdmodePropTag:
			if (!(switchType.isSet()) && !(switchSearch.isSet()) && cli::switchUnswitched.empty())
				options.mode = cmdmodeHelp;
			else if (
				switchSearch.isSet() && switchType.isSet() &&
				proptype::PropTypeNameToPropType(switchType[0]) == PT_UNSPECIFIED)
				options.mode = cmdmodeHelp;
			else if (switchFlag.hasULONG(0) && (switchSearch.isSet() || switchType.isSet()))
				options.mode = cmdmodeHelp;

			break;
		case cmdmodeSmartView:
			if (switchParser.atULONG(0) == 0) options.mode = cmdmodeHelp;

			break;
		case cmdmodeContents:
			if (!(switchContents.isSet()) && !(switchAssociatedContents.isSet())) options.mode = cmdmodeHelp;

			break;
		case cmdmodeMAPIMIME:
			// Can't convert both ways at once
			if (switchMAPI.isSet() && switchMIME.isSet()) options.mode = cmdmodeHelp;
			// Make sure there's no MIME-only options specified in a MIME->MAPI conversion
			else if (switchMAPI.isSet() && (switchRFC822.isSet() || switchEncoding.isSet() || switchWrap.isSet()))
				options.mode = cmdmodeHelp;
			// Make sure there's no MAPI-only options specified in a MAPI->MIME conversion
			else if (switchMIME.isSet() && (switchCharset.isSet() || switchUnicode.isSet()))
				options.mode = cmdmodeHelp;

			break;
		case cmdmodeProfile:
			if (!switchProfile.empty() && switchOutput.empty())
				options.mode = cmdmodeHelp;
			else if (switchProfile.empty() && !switchOutput.empty())
				options.mode = cmdmodeHelp;
			else if (switchProfileSection.empty() && switchProfile.empty())
				options.mode = cmdmodeHelp;

			break;
		default:
			break;
		}
	}
} // namespace cli