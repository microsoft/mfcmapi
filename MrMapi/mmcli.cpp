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

namespace cli
{
	enum __CommandLineSwitch
	{
		switchSearch = switchFirstSwitch, // '-s'
		switchDecimal, // '-n'
		switchFolder, // '-f'
		switchOutput, // '-o'
		switchDispid, // '-d'
		switchType, // '-t'
		switchGuid, // '-g'
		switchError, // '-e'
		switchParser, // '-p'
		switchInput, // '-i'
		switchBinary, // '-b'
		switchAcl, // '-a'
		switchRule, // '-r'
		switchContents, // '-c'
		switchAssociatedContents, // '-h'
		switchMoreProperties, // '-m'
		switchNoAddins, // '-no'
		switchOnline, // '-on'
		switchMAPI, // '-ma'
		switchMIME, // '-mi'
		switchCCSFFlags, // '-cc'
		switchRFC822, // '-rf'
		switchWrap, // '-w'
		switchEncoding, // '-en'
		switchCharset, // '-ch'
		switchAddressBook, // '-ad'
		switchUnicode, // '-u'
		switchProfile, // '-pr'
		switchXML, // '-x'
		switchSubject, // '-su'
		switchMessageClass, // '-me'
		switchMSG, // '-ms'
		switchList, // '-l'
		switchChildFolders, // '-chi'
		switchFid, // '-fi'
		switchMid, // '-mid'
		switchFlag, // '-fl'
		switchRecent, // '-re'
		switchStore, // '-st'
		switchVersion, // '-vers'
		switchSize, // '-si'
		switchPST, // '-pst'
		switchProfileSection, // '-profiles'
		switchByteSwapped, // '-b'
		switchReceiveFolder, // '-receivefolder'
		switchSkip, // '-sk'
		switchSearchState, // '-searchstate'
	};

	std::vector<OptParser> g_Parsers = {
		// clang-format off
		noSwitchParser ,
		helpParser,
		verboseParser,
		{switchSearch, L"Search", cmdmodeUnknown, 0, 0, OPT_DOPARTIALSEARCH},
		{switchDecimal, L"Number", cmdmodeUnknown, 0, 0, OPT_DODECIMAL},
		{switchFolder, L"Folder", cmdmodeUnknown, 1, 1, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_NEEDFOLDER | OPT_INITMFC},
		{switchOutput, L"Output", cmdmodeUnknown, 1, 1, OPT_NOOPT},
		{switchDispid, L"Dispids", cmdmodePropTag, 0, 0, OPT_DODISPID},
		{switchType, L"Type", cmdmodePropTag, 0, 1, OPT_DOTYPE},
		{switchGuid, L"Guids", cmdmodeGuid, 0, 0, OPT_NOOPT},
		{switchError, L"Error", cmdmodeErr, 0, 0, OPT_NOOPT},
		{switchParser, L"ParserType", cmdmodeSmartView, 1, 1, OPT_INITMFC | OPT_NEEDINPUTFILE},
		{switchInput, L"Input", cmdmodeUnknown, 1, 1, OPT_NOOPT},
		{switchBinary, L"Binary", cmdmodeSmartView, 0, 0, OPT_BINARYFILE},
		{switchAcl, L"Acl", cmdmodeAcls, 0, 0, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDFOLDER},
		{switchRule, L"Rules", cmdmodeRules, 0, 0, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDFOLDER},
		{switchContents, L"Contents", cmdmodeContents, 0, 0, OPT_DOCONTENTS | OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC},
		{switchAssociatedContents, L"HiddenContents", cmdmodeContents, 0, 0, OPT_DOASSOCIATEDCONTENTS | OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC},
		{switchMoreProperties, L"MoreProperties", cmdmodeUnknown, 0, 0, OPT_RETRYSTREAMPROPS},
		{switchNoAddins, L"NoAddins", cmdmodeUnknown, 0, 0, OPT_NOADDINS},
		{switchOnline, L"Online", cmdmodeUnknown, 0, 0, OPT_ONLINE},
		{switchMAPI, L"MAPI", cmdmodeMAPIMIME, 0, 0, OPT_NEEDMAPIINIT | OPT_INITMFC | OPT_NEEDINPUTFILE | OPT_NEEDOUTPUTFILE},
		{switchMIME, L"MIME", cmdmodeMAPIMIME, 0, 0, OPT_NEEDMAPIINIT | OPT_INITMFC | OPT_NEEDINPUTFILE | OPT_NEEDOUTPUTFILE},
		{switchCCSFFlags, L"CCSFFlags", cmdmodeMAPIMIME, 1, 1, OPT_NOOPT},
		{switchRFC822, L"RFC822", cmdmodeMAPIMIME, 0, 0, OPT_NOOPT},
		{switchWrap, L"Wrap", cmdmodeMAPIMIME, 1, 1, OPT_NOOPT},
		{switchEncoding, L"Encoding", cmdmodeMAPIMIME, 1, 1, OPT_NOOPT},
		{switchCharset, L"Charset", cmdmodeMAPIMIME, 3, 3, OPT_NOOPT},
		{switchAddressBook, L"AddressBook", cmdmodeMAPIMIME, 0, 0, OPT_NEEDMAPILOGON}, // special case which needs a logon
		{switchUnicode, L"Unicode", cmdmodeMAPIMIME, 0, 0, OPT_NOOPT},
		{switchProfile, L"Profile", cmdmodeUnknown, 0, 1, OPT_PROFILE},
		{switchXML, L"XML", cmdmodeXML, 0, 0, OPT_NEEDMAPIINIT | OPT_INITMFC | OPT_NEEDINPUTFILE},
		{switchSubject, L"Subject", cmdmodeContents, 0, 0, OPT_NOOPT},
		{switchMessageClass, L"MessageClass", cmdmodeContents, 1, 1, OPT_NOOPT},
		{switchMSG, L"MSG", cmdmodeContents, 0, 0, OPT_MSG},
		{switchList, L"List", cmdmodeContents, 0, 0, OPT_LIST},
		{switchChildFolders, L"ChildFolders", cmdmodeChildFolders, 0, 1, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDFOLDER},
		{switchFid, L"FID", cmdmodeFidMid, 0, 1, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDSTORE},
		{switchMid, L"MID", cmdmodeFidMid, 0, 1, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_MID},
		{switchFlag, L"Flag", cmdmodeUnknown, 1, 1, OPT_NOOPT}, // can't know until we parse the argument
		{switchRecent, L"Recent", cmdmodeContents, 1, 1, OPT_NOOPT},
		{switchStore, L"Store", cmdmodeStoreProperties, 0, 1, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC},
		{switchVersion, L"Version", cmdmodeUnknown, 1, 1, OPT_NOOPT},
		{switchSize, L"Size", cmdmodeFolderSize, 0, 0, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDFOLDER},
		{switchPST, L"PST", cmdmodePST, 0, 0, OPT_NEEDINPUTFILE},
		{switchProfileSection, L"ProfileSection", cmdmodeProfile, 1, 1, OPT_PROFILE | OPT_NEEDMAPIINIT | OPT_INITMFC},
		{switchByteSwapped, L"ByteSwapped", cmdmodeProfile, 0, 0, OPT_PROFILE | OPT_NEEDMAPIINIT | OPT_INITMFC},
		{switchReceiveFolder, L"ReceiveFolder", cmdmodeReceiveFolder, 0, 1, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_NEEDSTORE | OPT_INITMFC},
		{switchSkip, L"Skip", cmdmodeUnknown, 0, 0, OPT_SKIPATTACHMENTS},
		{switchSearchState, L"SearchState", cmdmodeSearchState, 0, 1, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDFOLDER},

		// If we want to add aliases for any switches, add them here
		{switchHelp, L"Help", cmdmodeHelpFull, 0, 0, OPT_INITMFC},
		// clang-format on
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
		printf("   MrMAPI -%ws\n", g_Parsers[switchHelp].szSwitch);
		printf(
			"   MrMAPI [-%ws] [-%ws] [-%ws] [-%ws <type>] <property number>|<property name>\n",
			g_Parsers[switchSearch].szSwitch,
			g_Parsers[switchDispid].szSwitch,
			g_Parsers[switchDecimal].szSwitch,
			g_Parsers[switchType].szSwitch);
		printf("   MrMAPI -%ws\n", g_Parsers[switchGuid].szSwitch);
		printf("   MrMAPI -%ws <error>\n", g_Parsers[switchError].szSwitch);
		printf(
			"   MrMAPI -%ws <type> -%ws <input file> [-%ws] [-%ws <output file>]\n",
			g_Parsers[switchParser].szSwitch,
			g_Parsers[switchInput].szSwitch,
			g_Parsers[switchBinary].szSwitch,
			g_Parsers[switchOutput].szSwitch);
		printf(
			"   MrMAPI -%ws <flag value> [-%ws] [-%ws] <property number>|<property name>\n",
			g_Parsers[switchFlag].szSwitch,
			g_Parsers[switchDispid].szSwitch,
			g_Parsers[switchDecimal].szSwitch);
		printf("   MrMAPI -%ws <flag name>\n", g_Parsers[switchFlag].szSwitch);
		printf(
			"   MrMAPI -%ws [-%ws <profile>] [-%ws <folder>]\n",
			g_Parsers[switchRule].szSwitch,
			g_Parsers[switchProfile].szSwitch,
			g_Parsers[switchFolder].szSwitch);
		printf(
			"   MrMAPI -%ws [-%ws <profile>] [-%ws <folder>]\n",
			g_Parsers[switchAcl].szSwitch,
			g_Parsers[switchProfile].szSwitch,
			g_Parsers[switchFolder].szSwitch);
		printf(
			"   MrMAPI -%ws | -%ws [-%ws <profile>] [-%ws <folder>] [-%ws <output directory>]\n",
			g_Parsers[switchContents].szSwitch,
			g_Parsers[switchAssociatedContents].szSwitch,
			g_Parsers[switchProfile].szSwitch,
			g_Parsers[switchFolder].szSwitch,
			g_Parsers[switchOutput].szSwitch);
		printf(
			"          [-%ws <subject>] [-%ws <message class>] [-%ws] [-%ws] [-%ws <count>] [-%ws]\n",
			g_Parsers[switchSubject].szSwitch,
			g_Parsers[switchMessageClass].szSwitch,
			g_Parsers[switchMSG].szSwitch,
			g_Parsers[switchList].szSwitch,
			g_Parsers[switchRecent].szSwitch,
			g_Parsers[switchSkip].szSwitch);
		printf(
			"   MrMAPI -%ws [-%ws <profile>] [-%ws <folder>]\n",
			g_Parsers[switchChildFolders].szSwitch,
			g_Parsers[switchProfile].szSwitch,
			g_Parsers[switchFolder].szSwitch);
		printf(
			"   MrMAPI -%ws -%ws <path to input file> -%ws <path to output file> [-%ws]\n",
			g_Parsers[switchXML].szSwitch,
			g_Parsers[switchInput].szSwitch,
			g_Parsers[switchOutput].szSwitch,
			g_Parsers[switchSkip].szSwitch);
		printf(
			"   MrMAPI -%ws [fid] [-%ws [mid]] [-%ws <profile>]\n",
			g_Parsers[switchFid].szSwitch,
			g_Parsers[switchMid].szSwitch,
			g_Parsers[switchProfile].szSwitch);
		printf(
			"   MrMAPI [<property number>|<property name>] -%ws [<store num>] [-%ws <profile>]\n",
			g_Parsers[switchStore].szSwitch,
			g_Parsers[switchProfile].szSwitch);
		printf(
			"   MrMAPI [<property number>|<property name>] -%ws <folder> [-%ws <profile>]\n",
			g_Parsers[switchFolder].szSwitch,
			g_Parsers[switchProfile].szSwitch);
		printf(
			"   MrMAPI -%ws -%ws <folder> [-%ws <profile>]\n",
			g_Parsers[switchSize].szSwitch,
			g_Parsers[switchFolder].szSwitch,
			g_Parsers[switchProfile].szSwitch);
		printf(
			"   MrMAPI -%ws | -%ws -%ws <path to input file> -%ws <path to output file> [-%ws <conversion flags>]\n",
			g_Parsers[switchMAPI].szSwitch,
			g_Parsers[switchMIME].szSwitch,
			g_Parsers[switchInput].szSwitch,
			g_Parsers[switchOutput].szSwitch,
			g_Parsers[switchCCSFFlags].szSwitch);
		printf(
			"          [-%ws] [-%ws <Decimal number of characters>] [-%ws <Decimal number indicating encoding>]\n",
			g_Parsers[switchRFC822].szSwitch,
			g_Parsers[switchWrap].szSwitch,
			g_Parsers[switchEncoding].szSwitch);
		printf(
			"          [-%ws] [-%ws] [-%ws CodePage CharSetType CharSetApplyType]\n",
			g_Parsers[switchAddressBook].szSwitch,
			g_Parsers[switchUnicode].szSwitch,
			g_Parsers[switchCharset].szSwitch);
		printf(
			"   MrMAPI -%ws -%ws <path to input file>\n",
			g_Parsers[switchPST].szSwitch,
			g_Parsers[switchInput].szSwitch);
		printf(
			"   MrMAPI -%ws [<profile> [-%ws <profilesection> [-%ws]] -%ws <output file>]\n",
			g_Parsers[switchProfile].szSwitch,
			g_Parsers[switchProfileSection].szSwitch,
			g_Parsers[switchByteSwapped].szSwitch,
			g_Parsers[switchOutput].szSwitch);
		printf(
			"   MrMAPI -%ws [<store num>] [-%ws <profile>]\n",
			g_Parsers[switchReceiveFolder].szSwitch,
			g_Parsers[switchProfile].szSwitch);
		printf(
			"   MrMAPI -%ws -%ws <folder> [-%ws <profile>]\n",
			g_Parsers[switchSearchState].szSwitch,
			g_Parsers[switchFolder].szSwitch,
			g_Parsers[switchProfile].szSwitch);

		if (bFull)
		{
			printf("\n");
			printf("All switches may be shortened if the intended switch is unambiguous.\n");
			printf("For example, -T may be used instead of -%ws.\n", g_Parsers[switchType].szSwitch);
		}
		printf("\n");
		printf("   Help:\n");
		printf("   -%ws   Display expanded help.\n", g_Parsers[switchHelp].szSwitch);
		if (bFull)
		{
			printf("\n");
			printf("   Property Tag Lookup:\n");
			printf("   -S   (or -%ws) Perform substring search.\n", g_Parsers[switchSearch].szSwitch);
			printf("           With no parameters prints all known properties.\n");
			printf("   -D   (or -%ws) Search dispids.\n", g_Parsers[switchDispid].szSwitch);
			printf(
				"   -N   (or -%ws) Number is in decimal. Ignored for non-numbers.\n",
				g_Parsers[switchDecimal].szSwitch);
			printf("   -T   (or -%ws) Print information on specified type.\n", g_Parsers[switchType].szSwitch);
			printf("           With no parameters prints list of known types.\n");
			printf("           When combined with -S, restrict output to given type.\n");
			printf("   -G   (or -%ws) Display list of known guids.\n", g_Parsers[switchGuid].szSwitch);
			printf("\n");
			printf("   Flag Lookup:\n");
			printf("   -Fl  (or -%ws) Look up flags for specified property.\n", g_Parsers[switchFlag].szSwitch);
			printf("           May be combined with -D and -N switches, but all flag values must be in hex.\n");
			printf("   -Fl  (or -%ws) Look up flag name and output its value.\n", g_Parsers[switchFlag].szSwitch);
			printf("\n");
			printf("   Error Parsing:\n");
			printf(
				"   -E   (or -%ws) Map an error code to its name and vice versa.\n", g_Parsers[switchError].szSwitch);
			printf("           May be combined with -S and -N switches.\n");
			printf("\n");
			printf("   Smart View Parsing:\n");
			printf(
				"   -P   (or -%ws) Parser type (number). See list below for supported parsers.\n",
				g_Parsers[switchParser].szSwitch);
			printf(
				"   -B   (or -%ws) Input file is binary. Default is hex encoded text.\n",
				g_Parsers[switchBinary].szSwitch);
			printf("\n");
			printf("   Rules Table:\n");
			printf("   -R   (or -%ws) Output rules table. Profile optional.\n", g_Parsers[switchRule].szSwitch);
			printf("\n");
			printf("   ACL Table:\n");
			printf("   -A   (or -%ws) Output ACL table. Profile optional.\n", g_Parsers[switchAcl].szSwitch);
			printf("\n");
			printf("   Contents Table:\n");
			printf(
				"   -C   (or -%ws) Output contents table. May be combined with -H. Profile optional.\n",
				g_Parsers[switchContents].szSwitch);
			printf(
				"   -H   (or -%ws) Output associated contents table. May be combined with -C. Profile optional\n",
				g_Parsers[switchAssociatedContents].szSwitch);
			printf("   -Su  (or -%ws) Subject of messages to output.\n", g_Parsers[switchSubject].szSwitch);
			printf("   -Me  (or -%ws) Message class of messages to output.\n", g_Parsers[switchMessageClass].szSwitch);
			printf("   -Ms  (or -%ws) Output as .MSG instead of XML.\n", g_Parsers[switchMSG].szSwitch);
			printf(
				"   -L   (or -%ws) List details to screen and do not output files.\n", g_Parsers[switchList].szSwitch);
			printf(
				"   -Re  (or -%ws) Restrict output to the 'count' most recent messages.\n",
				g_Parsers[switchRecent].szSwitch);
			printf("\n");
			printf("   Child Folders:\n");
			printf(
				"   -Chi (or -%ws) Display child folders of selected folder.\n",
				g_Parsers[switchChildFolders].szSwitch);
			printf("\n");
			printf("   MSG File Properties\n");
			printf("   -X   (or -%ws) Output properties of an MSG file as XML.\n", g_Parsers[switchXML].szSwitch);
			printf("\n");
			printf("   MID/FID Lookup\n");
			printf("   -Fi  (or -%ws) Folder ID (FID) to search for.\n", g_Parsers[switchFid].szSwitch);
			printf(
				"           If -%ws is specified without a FID, search/display all folders\n",
				g_Parsers[switchFid].szSwitch);
			printf("   -Mid (or -%ws) Message ID (MID) to search for.\n", g_Parsers[switchMid].szSwitch);
			printf(
				"           If -%ws is specified without a MID, display all messages in folders specified by the FID "
				"parameter.\n",
				g_Parsers[switchMid].szSwitch);
			printf("\n");
			printf("   Store Properties\n");
			printf("   -St  (or -%ws) Output properties of stores as XML.\n", g_Parsers[switchStore].szSwitch);
			printf("           If store number is specified, outputs properties of a single store.\n");
			printf("           If a property is specified, outputs only that property.\n");
			printf("\n");
			printf("   Folder Properties\n");
			printf("   -F   (or -%ws) Output properties of a folder as XML.\n", g_Parsers[switchFolder].szSwitch);
			printf("           If a property is specified, outputs only that property.\n");
			printf("   -Size         Output size of a folder and all subfolders.\n");
			printf("           Use -%ws to specify which folder to scan.\n", g_Parsers[switchFolder].szSwitch);
			printf("   -SearchState  Output search folder state.\n");
			printf("           Use -%ws to specify which folder to scan.\n", g_Parsers[switchFolder].szSwitch);
			printf("\n");
			printf("   MAPI <-> MIME Conversion:\n");
			printf(
				"   -Ma  (or -%ws) Convert an EML file to MAPI format (MSG file).\n", g_Parsers[switchMAPI].szSwitch);
			printf(
				"   -Mi  (or -%ws) Convert an MSG file to MIME format (EML file).\n", g_Parsers[switchMIME].szSwitch);
			printf(
				"   -I   (or -%ws) Indicates the input file for conversion, either a MIME-formatted EML file or an MSG "
				"file.\n",
				g_Parsers[switchInput].szSwitch);
			printf(
				"   -O   (or -%ws) Indicates the output file for the conversion.\n", g_Parsers[switchOutput].szSwitch);
			printf(
				"   -Cc  (or -%ws) Indicates specific flags to pass to the converter.\n",
				g_Parsers[switchCCSFFlags].szSwitch);
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
				g_Parsers[switchRFC822].szSwitch);
			printf("           If not present, RFC1521 is used instead.\n");
			printf(
				"   -W   (or -%ws) (MAPI->MIME only) Indicates the maximum number of characters in each line in the\n",
				g_Parsers[switchWrap].szSwitch);
			printf("           generated EML. Default value is 74. A value of 0 indicates no wrapping.\n");
			printf(
				"   -En  (or -%ws) (MAPI->MIME only) Indicates the encoding type to use. Supported values are:\n",
				g_Parsers[switchEncoding].szSwitch);
			printf("              1 - Base64\n");
			printf("              2 - UUENCODE\n");
			printf("              3 - Quoted-Printable\n");
			printf("              4 - 7bit (DEFAULT)\n");
			printf("              5 - 8bit\n");
			printf(
				"   -Ad  (or -%ws) Pass MAPI Address Book into converter. Profile optional.\n",
				g_Parsers[switchAddressBook].szSwitch);
			printf(
				"   -U   (or -%ws) (MIME->MAPI only) The resulting MSG file should be unicode.\n",
				g_Parsers[switchUnicode].szSwitch);
			printf(
				"   -Ch  (or -%ws) (MIME->MAPI only) Character set - three required parameters:\n",
				g_Parsers[switchCharset].szSwitch);
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
			printf("   -I   (or -%ws) PST file to be analyzed.\n", g_Parsers[switchInput].szSwitch);
			printf("\n");
			printf("   Profiles\n");
			printf("   -Pr  (or -%ws) Output list of profiles\n", g_Parsers[switchProfile].szSwitch);
			printf("           If a profile is specified, exports that profile.\n");
			printf("   -ProfileSection If specified, output specific profile section.\n");
			printf(
				"   -B   (or -%ws) If specified, profile section guid is byte swapped.\n",
				g_Parsers[switchByteSwapped].szSwitch);
			printf(
				"   -O   (or -%ws) Indicates the output file for profile export.\n", g_Parsers[switchOutput].szSwitch);
			printf("           Required if a profile is specified.\n");
			printf("\n");
			printf("   Receive Folder Table\n");
			printf(
				"   -%ws Displays Receive Folder Table for the specified store\n",
				g_Parsers[switchReceiveFolder].szSwitch);
			printf("\n");
			printf("   Universal Options:\n");
			printf("   -I   (or -%ws) Input file.\n", g_Parsers[switchInput].szSwitch);
			printf("   -O   (or -%ws) Output file or directory.\n", g_Parsers[switchOutput].szSwitch);
			printf(
				"   -F   (or -%ws) Folder to scan. Default is Inbox. See list below for supported folders.\n",
				g_Parsers[switchFolder].szSwitch);
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
			printf("   -Pr  (or -%ws) Profile for MAPILogonEx.\n", g_Parsers[switchProfile].szSwitch);
			printf(
				"   -M   (or -%ws) More properties. Tries harder to get stream properties. May take longer.\n",
				g_Parsers[switchMoreProperties].szSwitch);
			printf("   -Sk  (or -%ws) Skip embedded message attachments on export.\n", g_Parsers[switchSkip].szSwitch);
			printf("   -No  (or -%ws) No Addins. Don't load any add-ins.\n", g_Parsers[switchNoAddins].szSwitch);
			printf("   -On  (or -%ws) Online mode. Bypass cached mode.\n", g_Parsers[switchOnline].szSwitch);
			printf("   -V   (or -%ws) Verbose. Turn on all debug output.\n", g_Parsers[switchVerbose].szSwitch);
			printf("\n");
			printf("   MAPI Implementation Options:\n");
			printf("   -%ws MAPI Version to load - supported values\n", g_Parsers[switchVersion].szSwitch);
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

	// Return true if we succesfully peeled off a switch.
	// Return false on an error.
	_Check_return_ bool DoSwitch(OPTIONS* _options, int iSwitch, std::deque<std::wstring>& args)
	{
		MYOPTIONS* options = (MYOPTIONS*) _options;
		LPWSTR szEndPtr = nullptr;
		const auto arg0 = args.front();
		args.pop_front();

		switch (iSwitch)
		{
			// Global flags
		case switchFolder:
			options->ulFolder = strings::wstringToUlong(args.front(), 10);
			if (options->ulFolder)
			{
				options->lpszFolderPath = args.front();
				options->ulFolder = mapi::DEFAULT_INBOX;
			}

			args.pop_front();
			break;
		case switchInput:
			options->lpszInput = args.front();
			args.pop_front();
			break;
		case switchOutput:
			options->lpszOutput = args.front();
			args.pop_front();
			break;
		case switchProfile:
			// If we have a next argument and it's not an option, parse it as a profile name
			if (!args.empty() && switchNoSwitch == ParseArgument(args.front(), g_Parsers))
			{
				options->lpszProfile = args.front();
				args.pop_front();
			}

			break;
		case switchProfileSection:
			options->lpszProfileSection = args.front();
			args.pop_front();
			break;
		case switchByteSwapped:
			options->bByteSwapped = true;
			break;
		case switchVersion:
			options->lpszVersion = args.front();
			args.pop_front();
			break;
			// Proptag parsing
		case switchType:
			// If we have a next argument and it's not an option, parse it as a type
			if (!args.empty() && switchNoSwitch == ParseArgument(args.front(), g_Parsers))
			{
				options->ulTypeNum = proptype::PropTypeNameToPropType(args.front());
				args.pop_front();
			}
			break;
		case switchFlag:
			// We must have a next argument, but it could be a string or a number
			options->lpszFlagName = args.front();
			options->ulFlagValue = wcstoul(args.front().c_str(), &szEndPtr, 16);

			// Set mode based on whether the flag string was completely parsed as a number
			if (NULL == szEndPtr[0])
			{
				if (!bSetMode(options->mode, cmdmodePropTag))
				{
					return false;
				}

				options->options |= OPT_DOFLAG;
			}
			else
			{
				if (!bSetMode(options->mode, cmdmodeFlagSearch))
				{
					return false;
				}
			}

			args.pop_front();
			break;
			// Smart View parsing
		case switchParser:
			options->ulSVParser = strings::wstringToUlong(args.front(), 10);
			args.pop_front();
			break;
			// Contents tables
		case switchSubject:
			options->lpszSubject = args.front();
			args.pop_front();
			break;
		case switchMessageClass:
			options->lpszMessageClass = args.front();
			args.pop_front();
			break;
		case switchRecent:
			options->ulCount = strings::wstringToUlong(args.front(), 10);
			args.pop_front();
			break;
			// FID / MID
		case switchFid:
			if (!args.empty() && switchNoSwitch == ParseArgument(args.front(), g_Parsers))
			{
				options->lpszFid = args.front();
				args.pop_front();
			}

			break;
		case switchMid:
			if (!args.empty() && switchNoSwitch == ParseArgument(args.front(), g_Parsers))
			{
				options->lpszMid = args.front();
				args.pop_front();
			}
			else
			{
				// We use the blank string to remember the -mid parameter was passed and save having an extra flag
				// TODO: Check if this works
				options->lpszMid = L"";
			}

			break;
			// Store Properties / Receive Folder:
		case switchStore:
		case switchReceiveFolder:
			if (!args.empty() && switchNoSwitch == ParseArgument(args.front(), g_Parsers))
			{
				options->ulStore = wcstoul(args.front().c_str(), &szEndPtr, 10);

				// If we parsed completely, this was a store number
				if (NULL == szEndPtr[0])
				{
					// Increment ulStore so we can use to distinguish an unset value
					options->ulStore++;
					args.pop_front();
				}
				// Else it was a naked option - leave it on the stack
			}

			break;
			// MAPIMIME
		case switchMAPI:
			options->MAPIMIMEFlags |= MAPIMIME_TOMAPI;
			break;
		case switchMIME:
			options->MAPIMIMEFlags |= MAPIMIME_TOMIME;
			break;
		case switchCCSFFlags:
			options->convertFlags = static_cast<CCSFLAGS>(strings::wstringToUlong(args.front(), 10));
			args.pop_front();
			break;
		case switchRFC822:
			options->MAPIMIMEFlags |= MAPIMIME_RFC822;
			break;
		case switchWrap:
			options->ulWrapLines = strings::wstringToUlong(args.front(), 10);
			options->MAPIMIMEFlags |= MAPIMIME_WRAP;
			args.pop_front();
			break;
		case switchEncoding:
			options->ulEncodingType = strings::wstringToUlong(args.front(), 10);
			options->MAPIMIMEFlags |= MAPIMIME_ENCODING;
			args.pop_front();
			break;
		case switchCharset:
			options->ulCodePage = strings::wstringToUlong(args.front(), 10);
			args.pop_front();
			options->cSetType = static_cast<CHARSETTYPE>(strings::wstringToUlong(args.front(), 10));
			args.pop_front();
			if (options->cSetType > CHARSET_WEB)
			{
				return false;
			}

			options->cSetApplyType = static_cast<CSETAPPLYTYPE>(strings::wstringToUlong(args.front(), 10));
			args.pop_front();
			if (options->cSetApplyType > CSET_APPLY_TAG_ALL)
			{
				return false;
			}

			options->MAPIMIMEFlags |= MAPIMIME_CHARSET;
			break;
		case switchAddressBook:
			options->MAPIMIMEFlags |= MAPIMIME_ADDRESSBOOK;
			break;
		case switchUnicode:
			options->MAPIMIMEFlags |= MAPIMIME_UNICODE;
			break;

		case switchNoSwitch:
			// naked option without a flag - we only allow one of these
			if (!options->lpszUnswitchedOption.empty())
			{
				return false;
			} // He's already got one, you see.

			options->lpszUnswitchedOption = arg0;
			break;
		case switchUnknown:
			// display help
			return false;
		case switchHelp:
			break;
		case switchVerbose:
			break;
		case switchSearch:
			break;
		case switchDecimal:
			break;
		case switchDispid:
			break;
		case switchGuid:
			break;
		case switchError:
			break;
		case switchBinary:
			break;
		case switchAcl:
			break;
		case switchRule:
			break;
		case switchContents:
			break;
		case switchAssociatedContents:
			break;
		case switchMoreProperties:
			break;
		case switchNoAddins:
			break;
		case switchOnline:
			break;
		case switchXML:
			break;
		case switchMSG:
			break;
		case switchList:
			break;
		case switchChildFolders:
			break;
		case switchSize:
			break;
		case switchPST:
			break;
		case switchSkip:
			break;
		case switchSearchState:
			break;
		default:
			break;
		}

		return true;
	}

	void PostParseCheck(OPTIONS* _options)
	{
		MYOPTIONS* options = (MYOPTIONS*) _options;
		// Having processed the command line, we may not have determined a mode.
		// Some modes can be presumed by the switches we did see.

		// If we didn't get a mode set but we saw OPT_NEEDFOLDER, assume we're in folder dumping mode
		if (cmdmodeUnknown == options->mode && options->options & OPT_NEEDFOLDER) options->mode = cmdmodeFolderProps;

		// If we didn't get a mode set, but we saw OPT_PROFILE, assume we're in profile dumping mode
		if (cmdmodeUnknown == options->mode && options->options & OPT_PROFILE)
		{
			options->mode = cmdmodeProfile;
			options->options |= OPT_NEEDMAPIINIT | OPT_INITMFC;
		}

		// If we didn't get a mode set, assume we're in prop tag mode
		if (cmdmodeUnknown == options->mode) options->mode = cmdmodePropTag;

		// If we weren't passed an output file/directory, remember the current directory
		if (options->lpszOutput.empty() && options->mode != cmdmodeSmartView && options->mode != cmdmodeProfile)
		{
			char strPath[_MAX_PATH];
			GetCurrentDirectoryA(_MAX_PATH, strPath);

			options->lpszOutput = strings::LPCSTRToWstring(strPath);
		}

		// Validate that we have bare minimum to run
		if (options->options & OPT_NEEDINPUTFILE && options->lpszInput.empty())
			options->mode = cmdmodeHelp;
		else if (options->options & OPT_NEEDOUTPUTFILE && options->lpszOutput.empty())
			options->mode = cmdmodeHelp;

		switch (options->mode)
		{
		case cmdmodePropTag:
			if (!(options->options & OPT_DOTYPE) && !(options->options & OPT_DOPARTIALSEARCH) &&
				options->lpszUnswitchedOption.empty())
				options->mode = cmdmodeHelp;
			else if (
				options->options & OPT_DOPARTIALSEARCH && options->options & OPT_DOTYPE &&
				ulNoMatch == options->ulTypeNum)
				options->mode = cmdmodeHelp;
			else if (
				options->options & OPT_DOFLAG &&
				(options->options & OPT_DOPARTIALSEARCH || options->options & OPT_DOTYPE))
				options->mode = cmdmodeHelp;

			break;
		case cmdmodeSmartView:
			if (!options->ulSVParser) options->mode = cmdmodeHelp;

			break;
		case cmdmodeContents:
			if (!(options->options & OPT_DOCONTENTS) && !(options->options & OPT_DOASSOCIATEDCONTENTS))
				options->mode = cmdmodeHelp;

			break;
		case cmdmodeMAPIMIME:
#define CHECKFLAG(__flag) ((options->MAPIMIMEFlags & (__flag)) == (__flag))
			// Can't convert both ways at once
			if (CHECKFLAG(MAPIMIME_TOMAPI) && CHECKFLAG(MAPIMIME_TOMIME)) options->mode = cmdmodeHelp;
			// Make sure there's no MIME-only options specified in a MIME->MAPI conversion
			else if (
				CHECKFLAG(MAPIMIME_TOMAPI) &&
				(CHECKFLAG(MAPIMIME_RFC822) || CHECKFLAG(MAPIMIME_ENCODING) || CHECKFLAG(MAPIMIME_WRAP)))
				options->mode = cmdmodeHelp;
			// Make sure there's no MAPI-only options specified in a MAPI->MIME conversion
			else if (CHECKFLAG(MAPIMIME_TOMIME) && (CHECKFLAG(MAPIMIME_CHARSET) || CHECKFLAG(MAPIMIME_UNICODE)))
				options->mode = cmdmodeHelp;

			break;
		case cmdmodeProfile:
			if (!options->lpszProfile.empty() && options->lpszOutput.empty())
				options->mode = cmdmodeHelp;
			else if (options->lpszProfile.empty() && !options->lpszOutput.empty())
				options->mode = cmdmodeHelp;
			else if (!options->lpszProfileSection.empty() && options->lpszProfile.empty())
				options->mode = cmdmodeHelp;

			break;
		default:
			break;
		}
	}

	void PrintArgs(_In_ const MYOPTIONS& ProgOpts)
	{
		output::DebugPrint(DBGGeneric, L"Mode = %d\n", ProgOpts.mode);
		output::DebugPrint(DBGGeneric, L"options = 0x%08X\n", ProgOpts.options);
		output::DebugPrint(DBGGeneric, L"ulTypeNum = 0x%08X\n", ProgOpts.ulTypeNum);
		if (!ProgOpts.lpszUnswitchedOption.empty())
			output::DebugPrint(DBGGeneric, L"lpszUnswitchedOption = %ws\n", ProgOpts.lpszUnswitchedOption.c_str());
		if (!ProgOpts.lpszFlagName.empty())
			output::DebugPrint(DBGGeneric, L"lpszFlagName = %ws\n", ProgOpts.lpszFlagName.c_str());
		if (!ProgOpts.lpszFolderPath.empty())
			output::DebugPrint(DBGGeneric, L"lpszFolderPath = %ws\n", ProgOpts.lpszFolderPath.c_str());
		if (!ProgOpts.lpszInput.empty())
			output::DebugPrint(DBGGeneric, L"lpszInput = %ws\n", ProgOpts.lpszInput.c_str());
		if (!ProgOpts.lpszMessageClass.empty())
			output::DebugPrint(DBGGeneric, L"lpszMessageClass = %ws\n", ProgOpts.lpszMessageClass.c_str());
		if (!ProgOpts.lpszMid.empty()) output::DebugPrint(DBGGeneric, L"lpszMid = %ws\n", ProgOpts.lpszMid.c_str());
		if (!ProgOpts.lpszOutput.empty())
			output::DebugPrint(DBGGeneric, L"lpszOutput = %ws\n", ProgOpts.lpszOutput.c_str());
		if (!ProgOpts.lpszProfile.empty())
			output::DebugPrint(DBGGeneric, L"lpszProfile = %ws\n", ProgOpts.lpszProfile.c_str());
		if (!ProgOpts.lpszProfileSection.empty())
			output::DebugPrint(DBGGeneric, L"lpszProfileSection = %ws\n", ProgOpts.lpszProfileSection.c_str());
		if (!ProgOpts.lpszSubject.empty())
			output::DebugPrint(DBGGeneric, L"lpszSubject = %ws\n", ProgOpts.lpszSubject.c_str());
		if (!ProgOpts.lpszVersion.empty())
			output::DebugPrint(DBGGeneric, L"lpszVersion = %ws\n", ProgOpts.lpszVersion.c_str());
	}
} // namespace cli