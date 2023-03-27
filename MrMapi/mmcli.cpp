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
	option switchNamedProps{L"NamedProps", cmdmodeNamedProps, 0, 0, OPT_INITALL | OPT_NEEDSTORE};
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
	option switchAccounts{L"Accounts", cmdmodeEnumAccounts, 0, 0, OPT_INITALL | OPT_PROFILE};
	option switchIterate{L"Iterate", cmdmodeEnumAccounts, 0, 0, OPT_NOOPT};
	option switchWizard{L"Wizard", cmdmodeEnumAccounts, 0, 0, OPT_NOOPT};
	option switchFindProperty{L"FindProperty", cmdmodeContents, 1, USHRT_MAX, OPT_INITALL};
	option switchFindNamedProperty{L"FindNamedProperty", cmdmodeContents, 1, USHRT_MAX, OPT_INITALL};

	// If we want to add aliases for any switches, add them here
	option switchHelpAlias{L"Help", cmdmodeHelpFull, 0, 0, OPT_INITMFC};
#pragma warning(pop)

	std::vector<option*> g_options = {
		&switchHelp,
		&switchVerbose,
		&switchSearch,
		&switchDecimal,
		&switchNamedProps,
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
		&switchAccounts,
		&switchIterate,
		&switchWizard,
		&switchFindProperty,
		&switchFindNamedProperty,
		// If we want to add aliases for any switches, add them here
		&switchHelpAlias,
	};

	void DisplayUsage(BOOL bFull) noexcept
	{
		wprintf(L"MAPI data collection and parsing tool. Supports property tag lookup, error translation,\n");
		wprintf(L"   smart view processing, rule tables, ACL tables, contents tables, and MAPI<->MIME conversion.\n");
		if (bFull)
		{
			if (!g_lpMyAddins.empty())
			{
				wprintf(L"Addins Loaded:\n");
				for (const auto& addIn : g_lpMyAddins)
				{
					wprintf(L"   %ws\n", addIn.szName);
				}
			}

			wprintf(L"MrMAPI currently knows:\n");
			wprintf(L"%6u property tags\n", static_cast<UINT>(PropTagArray.size()));
			wprintf(L"%6u dispids\n", static_cast<UINT>(NameIDArray.size()));
			wprintf(L"%6u types\n", static_cast<UINT>(PropTypeArray.size()));
			wprintf(L"%6u guids\n", static_cast<UINT>(PropGuidArray.size()));
			wprintf(L"%6lu errors\n", error::g_ulErrorArray);
			wprintf(L"%6u smart view parsers\n", static_cast<UINT>(SmartViewParserTypeArray.size()) - 1);
			wprintf(L"\n");
		}

		wprintf(L"Usage:\n");
		wprintf(L"   MrMAPI -%ws\n", switchHelp.name());
		wprintf(
			L"   MrMAPI [-%ws] [-%ws] [-%ws] [-%ws <type>] <property number>|<property name>\n",
			switchSearch.name(),
			switchDispid.name(),
			switchDecimal.name(),
			switchType.name());
		wprintf(L"   MrMAPI -%ws\n", switchGuid.name());
		wprintf(L"   MrMAPI -%ws <error>\n", switchError.name());
		wprintf(
			L"   MrMAPI -%ws <type> -%ws <input file> [-%ws] [-%ws <output file>]\n",
			switchParser.name(),
			switchInput.name(),
			switchBinary.name(),
			switchOutput.name());
		wprintf(
			L"   MrMAPI -%ws <flag value> [-%ws] [-%ws] <property number>|<property name>\n",
			switchFlag.name(),
			switchDispid.name(),
			switchDecimal.name());
		wprintf(L"   MrMAPI -%ws <flag name>\n", switchFlag.name());
		wprintf(
			L"   MrMAPI -%ws [-%ws <profile>] [-%ws <folder>]\n",
			switchRule.name(),
			switchProfile.name(),
			switchFolder.name());
		wprintf(
			L"   MrMAPI -%ws [-%ws <profile>] [-%ws <folder>]\n",
			switchAcl.name(),
			switchProfile.name(),
			switchFolder.name());
		wprintf(
			L"   MrMAPI -%ws | -%ws [-%ws <profile>] [-%ws <folder>] [-%ws <output directory>]\n",
			switchContents.name(),
			switchAssociatedContents.name(),
			switchProfile.name(),
			switchFolder.name(),
			switchOutput.name());
		wprintf(
			L"          [-%ws <subject>] [-%ws <message class>] [-%ws] [-%ws] [-%ws <count>] [-%ws]\n",
			switchSubject.name(),
			switchMessageClass.name(),
			switchMSG.name(),
			switchList.name(),
			switchRecent.name(),
			switchSkip.name());
		wprintf(
			L"   MrMAPI -%ws [-%ws <profile>] [-%ws <folder>]\n",
			switchChildFolders.name(),
			switchProfile.name(),
			switchFolder.name());
		wprintf(
			L"   MrMAPI -%ws -%ws <path to input file> -%ws <path to output file> [-%ws]\n",
			switchXML.name(),
			switchInput.name(),
			switchOutput.name(),
			switchSkip.name());
		wprintf(
			L"   MrMAPI -%ws [fid] [-%ws [mid]] [-%ws <profile>]\n",
			switchFid.name(),
			switchMid.name(),
			switchProfile.name());
		wprintf(
			L"   MrMAPI [<property number>|<property name>] -%ws [<store num>] [-%ws <profile>]\n",
			switchStore.name(),
			switchProfile.name());
		wprintf(
			L"   MrMAPI [<property number>|<property name>] -%ws <folder> [-%ws <profile>]\n",
			switchFolder.name(),
			switchProfile.name());
		wprintf(
			L"   MrMAPI -%ws -%ws <folder> [-%ws <profile>]\n",
			switchSize.name(),
			switchFolder.name(),
			switchProfile.name());
		wprintf(
			L"   MrMAPI -%ws | -%ws -%ws <path to input file> -%ws <path to output file> [-%ws <conversion flags>]\n",
			switchMAPI.name(),
			switchMIME.name(),
			switchInput.name(),
			switchOutput.name(),
			switchCCSFFlags.name());
		wprintf(
			L"          [-%ws] [-%ws <Decimal number of characters>] [-%ws <Decimal number indicating encoding>]\n",
			switchRFC822.name(),
			switchWrap.name(),
			switchEncoding.name());
		wprintf(
			L"          [-%ws] [-%ws] [-%ws CodePage CharSetType CharSetApplyType]\n",
			switchAddressBook.name(),
			switchUnicode.name(),
			switchCharset.name());
		wprintf(L"   MrMAPI -%ws -%ws <path to input file>\n", switchPST.name(), switchInput.name());
		wprintf(
			L"   MrMAPI -%ws [<profile> [-%ws <profilesection> [-%ws]] -%ws <output file>]\n",
			switchProfile.name(),
			switchProfileSection.name(),
			switchByteSwapped.name(),
			switchOutput.name());
		wprintf(L"   MrMAPI -%ws [<store num>] [-%ws <profile>]\n", switchReceiveFolder.name(), switchProfile.name());
		wprintf(
			L"   MrMAPI -%ws -%ws <folder> [-%ws <profile>]\n",
			switchSearchState.name(),
			switchFolder.name(),
			switchProfile.name());
		wprintf(
			L"   MrMAPI -%ws [-%ws <profile>] [-%ws <path to output file>]\n",
			switchNamedProps.name(),
			switchProfile.name(),
			switchOutput.name());
		wprintf(
			L"   MrMAPI -%ws [-%ws] [-%ws <profile>]\n",
			switchAccounts.name(),
			switchIterate.name(),
			switchProfile.name());
		wprintf(
			L"   MrMAPI -%ws -%ws [-%ws <flags>] [-%ws <profile>]\n",
			switchAccounts.name(),
			switchWizard.name(),
			switchFlag.name(),
			switchProfile.name());
		wprintf(
			L"   MrMAPI -%ws [-%ws <property names>] [-%ws <dispid names>] [-%ws <folder>] [-%ws <output directory>]\n",
			switchContents.name(),
			switchFindProperty.name(),
			switchFindNamedProperty.name(),
			switchFolder.name(),
			switchOutput.name());

		if (bFull)
		{
			wprintf(L"\n");
			wprintf(L"All switches may be shortened if the intended switch is unambiguous.\n");
			wprintf(L"For example, -T may be used instead of -%ws.\n", switchType.name());
		}
		wprintf(L"\n");
		wprintf(L"   Help:\n");
		wprintf(L"   -%ws   Display expanded help.\n", switchHelp.name());
		if (bFull)
		{
			wprintf(L"\n");
			wprintf(L"   Property Tag Lookup:\n");
			wprintf(L"   -S   (or -%ws) Perform substring search.\n", switchSearch.name());
			wprintf(L"           With no parameters prints all known properties.\n");
			wprintf(L"   -D   (or -%ws) Search dispids.\n", switchDispid.name());
			wprintf(L"   -N   (or -%ws) Number is in decimal. Ignored for non-numbers.\n", switchDecimal.name());
			wprintf(L"   -T   (or -%ws) Print information on specified type.\n", switchType.name());
			wprintf(L"           With no parameters prints list of known types.\n");
			wprintf(L"           When combined with -S, restrict output to given type.\n");
			wprintf(L"   -G   (or -%ws) Display list of known guids.\n", switchGuid.name());
			wprintf(L"\n");
			wprintf(L"   Flag Lookup:\n");
			wprintf(L"   -Fl  (or -%ws) Look up flags for specified property.\n", switchFlag.name());
			wprintf(L"           May be combined with -D and -N switches, but all flag values must be in hex.\n");
			wprintf(L"   -Fl  (or -%ws) Look up flag name and output its value.\n", switchFlag.name());
			wprintf(L"\n");
			wprintf(L"   Error Parsing:\n");
			wprintf(L"   -E   (or -%ws) Map an error code to its name and vice versa.\n", switchError.name());
			wprintf(L"           May be combined with -S and -N switches.\n");
			wprintf(L"\n");
			wprintf(L"   Smart View Parsing:\n");
			wprintf(
				L"   -P   (or -%ws) Parser type (number). See list below for supported parsers.\n",
				switchParser.name());
			wprintf(L"   -B   (or -%ws) Input file is binary. Default is hex encoded text.\n", switchBinary.name());
			wprintf(L"\n");
			wprintf(L"   Rules Table:\n");
			wprintf(L"   -R   (or -%ws) Output rules table. Profile optional.\n", switchRule.name());
			wprintf(L"\n");
			wprintf(L"   ACL Table:\n");
			wprintf(L"   -A   (or -%ws) Output ACL table. Profile optional.\n", switchAcl.name());
			wprintf(L"\n");
			wprintf(L"   Contents Table:\n");
			wprintf(
				L"   -C   (or -%ws) Output contents table. May be combined with -H. Profile optional.\n",
				switchContents.name());
			wprintf(
				L"   -H   (or -%ws) Output associated contents table. May be combined with -C. Profile optional\n",
				switchAssociatedContents.name());
			wprintf(L"   -Su  (or -%ws) Subject of messages to output.\n", switchSubject.name());
			wprintf(L"   -Me  (or -%ws) Message class of messages to output.\n", switchMessageClass.name());
			wprintf(L"   -Ms  (or -%ws) Output as .MSG instead of XML.\n", switchMSG.name());
			wprintf(L"   -L   (or -%ws) List details to screen and do not output files.\n", switchList.name());
			wprintf(L"   -Re  (or -%ws) Restrict output to the 'count' most recent messages.\n", switchRecent.name());
			wprintf(
				L"   -FindP  (or -%ws) Restrict output to messages which contain given properties.\n",
				switchFindProperty.name());
			wprintf(
				L"   -FindN  (or -%ws) Restrict output to messages which contain given named properties.\n",
				switchFindNamedProperty.name());
			wprintf(L"\n");
			wprintf(L"   Child Folders:\n");
			wprintf(L"   -Chi (or -%ws) Display child folders of selected folder.\n", switchChildFolders.name());
			wprintf(L"\n");
			wprintf(L"   MSG File Properties\n");
			wprintf(L"   -X   (or -%ws) Output properties of an MSG file as XML.\n", switchXML.name());
			wprintf(L"\n");
			wprintf(L"   MID/FID Lookup\n");
			wprintf(L"   -Fi  (or -%ws) Folder ID (FID) to search for.\n", switchFid.name());
			wprintf(L"           If -%ws is specified without a FID, search/display all folders\n", switchFid.name());
			wprintf(L"   -Mid (or -%ws) Message ID (MID) to search for.\n", switchMid.name());
			wprintf(
				L"           If -%ws is specified without a MID, display all messages in folders specified by the FID "
				"parameter.\n",
				switchMid.name());
			wprintf(L"\n");
			wprintf(L"   Store Properties\n");
			wprintf(L"   -St  (or -%ws) Output properties of stores as XML.\n", switchStore.name());
			wprintf(L"           If store number is specified, outputs properties of a single store.\n");
			wprintf(L"           If a property is specified, outputs only that property.\n");
			wprintf(L"\n");
			wprintf(L"   Folder Properties\n");
			wprintf(L"   -F   (or -%ws) Output properties of a folder as XML.\n", switchFolder.name());
			wprintf(L"           If a property is specified, outputs only that property.\n");
			wprintf(L"   -Size         Output size of a folder and all subfolders.\n");
			wprintf(L"           Use -%ws to specify which folder to scan.\n", switchFolder.name());
			wprintf(L"   -SearchState  Output search folder state.\n");
			wprintf(L"           Use -%ws to specify which folder to scan.\n", switchFolder.name());
			wprintf(L"\n");
			wprintf(L"   MAPI <-> MIME Conversion:\n");
			wprintf(L"   -Ma  (or -%ws) Convert an EML file to MAPI format (MSG file).\n", switchMAPI.name());
			wprintf(L"   -Mi  (or -%ws) Convert an MSG file to MIME format (EML file).\n", switchMIME.name());
			wprintf(
				L"   -I   (or -%ws) Indicates the input file for conversion, either a MIME-formatted EML file or an "
				L"MSG "
				"file.\n",
				switchInput.name());
			wprintf(L"   -O   (or -%ws) Indicates the output file for the conversion.\n", switchOutput.name());
			wprintf(L"   -Cc  (or -%ws) Indicates specific flags to pass to the converter.\n", switchCCSFFlags.name());
			wprintf(L"           Available values (these may be OR'ed together):\n");
			wprintf(L"              MIME -> MAPI:\n");
			wprintf(L"                CCSF_SMTP:        0x02\n");
			wprintf(L"                CCSF_INCLUDE_BCC: 0x20\n");
			wprintf(L"                CCSF_USE_RTF:     0x80\n");
			wprintf(L"              MAPI -> MIME:\n");
			wprintf(L"                CCSF_NOHEADERS:        0x000004\n");
			wprintf(L"                CCSF_USE_TNEF:         0x000010\n");
			wprintf(L"                CCSF_8BITHEADERS:      0x000040\n");
			wprintf(L"                CCSF_PLAIN_TEXT_ONLY:  0x001000\n");
			wprintf(L"                CCSF_NO_MSGID:         0x004000\n");
			wprintf(L"                CCSF_EMBEDDED_MESSAGE: 0x008000\n");
			wprintf(L"                CCSF_PRESERVE_SOURCE:  0x040000\n");
			wprintf(L"                CCSF_GLOBAL_MESSAGE:   0x200000\n");
			wprintf(
				L"   -Rf  (or -%ws) (MAPI->MIME only) Indicates the EML should be generated in RFC822 format.\n",
				switchRFC822.name());
			wprintf(L"           If not present, RFC1521 is used instead.\n");
			wprintf(
				L"   -W   (or -%ws) (MAPI->MIME only) Indicates the maximum number of characters in each line in the\n",
				switchWrap.name());
			wprintf(L"           generated EML. Default value is 74. A value of 0 indicates no wrapping.\n");
			wprintf(
				L"   -En  (or -%ws) (MAPI->MIME only) Indicates the encoding type to use. Supported values are:\n",
				switchEncoding.name());
			wprintf(L"              1 - Base64\n");
			wprintf(L"              2 - UUENCODE\n");
			wprintf(L"              3 - Quoted-Printable\n");
			wprintf(L"              4 - 7bit (DEFAULT)\n");
			wprintf(L"              5 - 8bit\n");
			wprintf(
				L"   -Ad  (or -%ws) Pass MAPI Address Book into converter. Profile optional.\n",
				switchAddressBook.name());
			wprintf(
				L"   -U   (or -%ws) (MIME->MAPI only) The resulting MSG file should be unicode.\n",
				switchUnicode.name());
			wprintf(
				L"   -Ch  (or -%ws) (MIME->MAPI only) Character set - three required parameters:\n",
				switchCharset.name());
			wprintf(L"           CodePage - common values (others supported)\n");
			wprintf(L"              1252  - CP_USASCII      - Indicates the USASCII character set, Windows code page "
					"1252\n");
			wprintf(L"              1200  - CP_UNICODE      - Indicates the Unicode character set, Windows code page "
					"1200\n");
			wprintf(L"              50932 - CP_JAUTODETECT  - Indicates Japanese auto-detect (50932)\n");
			wprintf(L"              50949 - CP_KAUTODETECT  - Indicates Korean auto-detect (50949)\n");
			wprintf(L"              50221 - CP_ISO2022JPESC - Indicates the Internet character set ISO-2022-JP-ESC\n");
			wprintf(L"              50222 - CP_ISO2022JPSIO - Indicates the Internet character set ISO-2022-JP-SIO\n");
			wprintf(L"           CharSetType - supported values (see CHARSETTYPE)\n");
			wprintf(L"              0 - CHARSET_BODY\n");
			wprintf(L"              1 - CHARSET_HEADER\n");
			wprintf(L"              2 - CHARSET_WEB\n");
			wprintf(L"           CharSetApplyType - supported values (see CSETAPPLYTYPE)\n");
			wprintf(L"              0 - CSET_APPLY_UNTAGGED\n");
			wprintf(L"              1 - CSET_APPLY_ALL\n");
			wprintf(L"              2 - CSET_APPLY_TAG_ALL\n");
			wprintf(L"\n");
			wprintf(L"   PST Analysis\n");
			wprintf(L"   -PST Output statistics of a PST file.\n");
			wprintf(L"           If a property is specified, outputs only that property.\n");
			wprintf(L"   -I   (or -%ws) PST file to be analyzed.\n", switchInput.name());
			wprintf(L"\n");
			wprintf(L"   Profiles\n");
			wprintf(L"   -Pr  (or -%ws) Output list of profiles\n", switchProfile.name());
			wprintf(L"           If a profile is specified, exports that profile.\n");
			wprintf(L"   -ProfileSection If specified, output specific profile section.\n");
			wprintf(
				L"   -B   (or -%ws) If specified, profile section guid is byte swapped.\n", switchByteSwapped.name());
			wprintf(L"   -O   (or -%ws) Indicates the output file for profile export.\n", switchOutput.name());
			wprintf(L"           Required if a profile is specified.\n");
			wprintf(L"\n");
			wprintf(L"   Receive Folder Table\n");
			wprintf(L"   -%ws Displays Receive Folder Table for the specified store\n", switchReceiveFolder.name());
			wprintf(L"\n");
			wprintf(L"   Universal Options:\n");
			wprintf(L"   -I   (or -%ws) Input file.\n", switchInput.name());
			wprintf(L"   -O   (or -%ws) Output file or directory.\n", switchOutput.name());
			wprintf(
				L"   -F   (or -%ws) Folder to scan. Default is Inbox. See list below for supported folders.\n",
				switchFolder.name());
			wprintf(L"           Folders may also be specified by path:\n");
			wprintf(L"              \"Top of Information Store\\Calendar\"\n");
			wprintf(L"           Path may be preceeded by entry IDs for special folders using @ notation:\n");
			wprintf(L"              \"@PR_IPM_SUBTREE_ENTRYID\\Calendar\"\n");
			wprintf(
				L"           Path may further be preceeded by store number using # notation, which may either use a "
				"store number:\n");
			wprintf(L"              \"#0\\@PR_IPM_SUBTREE_ENTRYID\\Calendar\"\n");
			wprintf(L"           Or an entry ID:\n");
			wprintf(L"              \"#00112233445566...778899AABBCC\\@PR_IPM_SUBTREE_ENTRYID\\Calendar\"\n");
			wprintf(L"           MrMAPI's special folder constants may also be used:\n");
			wprintf(L"              \"@12\\Calendar\"\n");
			wprintf(L"              \"@1\"\n");
			wprintf(L"   -Pr  (or -%ws) Profile for MAPILogonEx.\n", switchProfile.name());
			wprintf(
				L"   -M   (or -%ws) More properties. Tries harder to get stream properties. May take longer.\n",
				switchMoreProperties.name());
			wprintf(L"   -Sk  (or -%ws) Skip embedded message attachments on export.\n", switchSkip.name());
			wprintf(L"   -No  (or -%ws) No Addins. Don't load any add-ins.\n", switchNoAddins.name());
			wprintf(L"   -On  (or -%ws) Online mode. Bypass cached mode.\n", switchOnline.name());
			wprintf(L"   -V   (or -%ws) Verbose. Turn on all debug output.\n", switchVerbose.name());
			wprintf(L"\n");
			wprintf(L"   MAPI Implementation Options:\n");
			wprintf(L"   -%ws MAPI Version to load - supported values\n", switchVersion.name());
			wprintf(L"           Supported values\n");
			wprintf(L"              0  - List all available MAPI binaries\n");
			wprintf(L"              1  - System MAPI\n");
			wprintf(L"              11 - Outlook 2003 (11)\n");
			wprintf(L"              12 - Outlook 2007 (12)\n");
			wprintf(L"              14 - Outlook 2010 (14)\n");
			wprintf(L"              15 - Outlook 2013 (15)\n");
			wprintf(L"              16 - Outlook 2016 (16)\n");
			wprintf(L"           You can also pass a string, which will load the first MAPI whose path contains the "
					"string.\n");
			wprintf(L"\n");
			wprintf(L"Smart View Parsers:\n");
			// Print smart view options
			for (ULONG i = 1; i < SmartViewParserTypeArray.size(); i++)
			{
				wprintf(L"   %2lu %ws\n", i, SmartViewParserTypeArray[i].lpszName);
			}

			wprintf(L"\n");
			wprintf(L"Folders:\n");
			// Print Folders
			for (ULONG i = 1; i < mapi::NUM_DEFAULT_PROPS; i++)
			{
				wprintf(L"   %2lu %ws\n", i, mapi::FolderNames[i]);
			}

			wprintf(L"\n");
			wprintf(L"Examples:\n");
			wprintf(L"   MrMAPI PR_DISPLAY_NAME\n");
			wprintf(L"\n");
			wprintf(L"   MrMAPI 0x3001001e\n");
			wprintf(L"   MrMAPI 3001001e\n");
			wprintf(L"   MrMAPI 3001\n");
			wprintf(L"\n");
			wprintf(L"   MrMAPI -n 12289\n");
			wprintf(L"\n");
			wprintf(L"   MrMAPI -t PT_LONG\n");
			wprintf(L"   MrMAPI -t 3102\n");
			wprintf(L"   MrMAPI -t\n");
			wprintf(L"\n");
			wprintf(L"   MrMAPI -s display\n");
			wprintf(L"   MrMAPI -s display -t PT_LONG\n");
			wprintf(L"   MrMAPI -t 102 -s display\n");
			wprintf(L"\n");
			wprintf(L"   MrMAPI -d dispidReminderTime\n");
			wprintf(L"   MrMAPI -d 0x8502\n");
			wprintf(L"   MrMAPI -d -s reminder\n");
			wprintf(L"   MrMAPI -d -n 34050\n");
			wprintf(L"\n");
			wprintf(L"   MrMAPI -p 17 -i webview.txt -o parsed.txt");
		}
	}

	void PostParseCheck(OPTIONS& options)
	{
		// Having processed the command line, we may not have determined a mode.
		// Some modes can be presumed by the switches we did see.

		if (switchFlag.isSet() && options.mode != cmdmodePropTag && options.mode != cmdmodeEnumAccounts)
		{
			if (!bSetMode(options.mode, cmdmodeFlagSearch))
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
		if (switchOutput.empty() && options.mode != cmdmodeSmartView && options.mode != cmdmodeProfile &&
			options.mode != cmdmodeNamedProps)
		{
			WCHAR strPath[_MAX_PATH];
			GetCurrentDirectoryW(_MAX_PATH, strPath);

			// Trick switchOutput into scanning this path as an argument.
			auto args = std::deque<std::wstring>{strPath};
			static_cast<void>(switchOutput.scanArgs(args, options, g_options));
		}

		if (switchFolder.empty())
		{
			// Trick switchFolder into scanning this path as an argument.
			OPTIONS fakeOptions{};
			auto args = std::deque<std::wstring>{L"13"};
			static_cast<void>(switchFolder.scanArgs(args, fakeOptions, g_options));
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
			else if (switchProfileSection.isSet() && switchProfile.empty())
				options.mode = cmdmodeHelp;

			break;
		default:
			break;
		}
	}
} // namespace cli