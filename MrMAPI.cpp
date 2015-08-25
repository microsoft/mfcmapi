#include "stdafx.h"

#include "MrMAPI\MrMAPI.h"
#include "MAPIFunctions.h"
#include "String.h"
#include "InterpretProp2.h"
#include "Smartview\SmartView.h"
#include "MrMAPI\MMAcls.h"
#include "MrMAPI\MMContents.h"
#include "MrMAPI\MMErr.h"
#include "MrMAPI\MMFidMid.h"
#include "MrMAPI\MMFolder.h"
#include "MrMAPI\MMProfile.h"
#include "MrMAPI\MMPropTag.h"
#include "MrMAPI\MMRules.h"
#include "MrMAPI\MMSmartView.h"
#include "MrMAPI\MMStore.h"
#include "MrMAPI\MMMapiMime.h"
#include <shlwapi.h>
#include "ImportProcs.h"
#include "MAPIStoreFunctions.h"
#include "MrMAPI\MMPst.h"
#include "MrMAPI\MMReceiveFolder.h"
#include "NamedPropCache.h"

// Initialize MFC for LoadString support later on
void InitMFC()
{
#pragma warning(push)
#pragma warning(disable:6309)
#pragma warning(disable:6387)
	if (!AfxWinInit(::GetModuleHandle(NULL), NULL, ::GetCommandLine(), 0)) return;
#pragma warning(pop)
}

_Check_return_ HRESULT MrMAPILogonEx(_In_ wstring lpszProfile, _Deref_out_opt_ LPMAPISESSION* lppSession)
{
	HRESULT hRes = S_OK;
	ULONG ulFlags = MAPI_EXTENDED | MAPI_NO_MAIL | MAPI_UNICODE | MAPI_NEW_SESSION;
	if (lpszProfile.empty()) ulFlags |= MAPI_USE_DEFAULT;

	WC_MAPI(MAPILogonEx(NULL, (LPTSTR)(lpszProfile.empty() ? NULL : lpszProfile.c_str()), NULL,
		ulFlags,
		lppSession));
	return hRes;
}

_Check_return_ HRESULT OpenExchangeOrDefaultMessageStore(
	_In_ LPMAPISESSION lpMAPISession,
	_Deref_out_opt_ LPMDB* lppMDB)
{
	if (!lpMAPISession || !lppMDB) return MAPI_E_INVALID_PARAMETER;
	HRESULT hRes = S_OK;
	LPMDB lpMDB = NULL;
	*lppMDB = NULL;

	WC_H(OpenMessageStoreGUID(lpMAPISession, pbExchangeProviderPrimaryUserGuid, &lpMDB));
	if (FAILED(hRes) || !lpMDB)
	{
		hRes = S_OK;
		WC_H(OpenDefaultMessageStore(lpMAPISession, &lpMDB));
	}

	if (SUCCEEDED(hRes) && lpMDB)
	{
		*lppMDB = lpMDB;
	}
	else
	{
		if (lpMDB) lpMDB->Release();
		if (SUCCEEDED(hRes)) hRes = MAPI_E_CALL_FAILED;
	}
	return hRes;
} // OpenExchangeOrDefaultMessageStore

enum __CommandLineSwitch
{
	switchNoSwitch = 0,       // not a switch
	switchUnknown,            // unknown switch
	switchHelp,               // '-h'
	switchVerbose,            // '-v'
	switchSearch,             // '-s'
	switchDecimal,            // '-n'
	switchFolder,             // '-f'
	switchOutput,             // '-o'
	switchDispid,             // '-d'
	switchType,               // '-t'
	switchGuid,               // '-g'
	switchError,              // '-e'
	switchParser,             // '-p'
	switchInput,              // '-i'
	switchBinary,             // '-b'
	switchAcl,                // '-a'
	switchRule,               // '-r'
	switchContents,           // '-c'
	switchAssociatedContents, // '-h'
	switchMoreProperties,     // '-m'
	switchNoAddins,           // '-no'
	switchOnline,             // '-on'
	switchMAPI,               // '-ma'
	switchMIME,               // '-mi'
	switchCCSFFlags,          // '-cc'
	switchRFC822,             // '-rf'
	switchWrap,               // '-w'
	switchEncoding,           // '-en'
	switchCharset,            // '-ch'
	switchAddressBook,        // '-ad'
	switchUnicode,            // '-u'
	switchProfile,            // '-pr'
	switchXML,                // '-x'
	switchSubject,            // '-su'
	switchMessageClass,       // '-me'
	switchMSG,                // '-ms'
	switchList,               // '-l'
	switchChildFolders,       // '-chi'
	switchFid,                // '-fi'
	switchMid,                // '-mid'
	switchFlag,               // '-fl'
	switchRecent,             // '-re'
	switchStore,              // '-st'
	switchVersion,            // '-vers'
	switchSize,               // '-si'
	switchPST,                // '-pst'
	switchProfileSection,     // '-profiles'
	switchByteSwapped,        // '-b'
	switchReceiveFolder,      // '-receivefolder'
	switchSkip,               // '-sk'
	switchSearchState,        // '-searchstate'
};

struct COMMANDLINE_SWITCH
{
	__CommandLineSwitch iSwitch;
	LPCSTR szSwitch;
};

// All entries before the aliases must be in the
// same order as the __CommandLineSwitch enum.
COMMANDLINE_SWITCH g_Switches[] =
{
	{ switchNoSwitch, "" },
	{ switchUnknown, "" },
	{ switchHelp, "?" },
	{ switchVerbose, "Verbose" },
	{ switchSearch, "Search" },
	{ switchDecimal, "Number" },
	{ switchFolder, "Folder" },
	{ switchOutput, "Output" },
	{ switchDispid, "Dispids" },
	{ switchType, "Type" },
	{ switchGuid, "Guids" },
	{ switchError, "Error" },
	{ switchParser, "ParserType" },
	{ switchInput, "Input" },
	{ switchBinary, "Binary" },
	{ switchAcl, "Acl" },
	{ switchRule, "Rules" },
	{ switchContents, "Contents" },
	{ switchAssociatedContents, "HiddenContents" },
	{ switchMoreProperties, "MoreProperties" },
	{ switchNoAddins, "NoAddins" },
	{ switchOnline, "Online" },
	{ switchMAPI, "MAPI" },
	{ switchMIME, "MIME" },
	{ switchCCSFFlags, "CCSFFlags" },
	{ switchRFC822, "RFC822" },
	{ switchWrap, "Wrap" },
	{ switchEncoding, "Encoding" },
	{ switchCharset, "Charset" },
	{ switchAddressBook, "AddressBook" },
	{ switchUnicode, "Unicode" },
	{ switchProfile, "Profile" },
	{ switchXML, "XML" },
	{ switchSubject, "Subject" },
	{ switchMessageClass, "MessageClass" },
	{ switchMSG, "MSG" },
	{ switchList, "List" },
	{ switchChildFolders, "ChildFolders" },
	{ switchFid, "FID" },
	{ switchMid, "MID" },
	{ switchFlag, "Flag" },
	{ switchRecent, "Recent" },
	{ switchStore, "Store" },
	{ switchVersion, "Version" },
	{ switchSize, "Size" },
	{ switchPST, "PST" },
	{ switchProfileSection, "ProfileSection" },
	{ switchByteSwapped, "ByteSwapped" },
	{ switchReceiveFolder, "ReceiveFolder" },
	{ switchSkip, "Skip" },
	{ switchSearchState, "SearchState" },
	// If we want to add aliases for any switches, add them here
	{ switchHelp, "Help" },
};
ULONG g_ulSwitches = _countof(g_Switches);

extern LPADDIN g_lpMyAddins;

void DisplayUsage(BOOL bFull)
{
	printf("MAPI data collection and parsing tool. Supports property tag lookup, error translation,\n");
	printf("   smart view processing, rule tables, ACL tables, contents tables, and MAPI<->MIME conversion.\n");
	LPADDIN lpCurAddIn = g_lpMyAddins;
	if (bFull)
	{
		if (lpCurAddIn)
		{
			printf("Addins Loaded:\n");
			while (lpCurAddIn)
			{
				printf("   %ws\n", lpCurAddIn->szName);
				lpCurAddIn = lpCurAddIn->lpNextAddIn;
			}
		}
		printf("MrMAPI currently knows:\n");
		printf("%6u property tags\n", ulPropTagArray);
		printf("%6u dispids\n", ulNameIDArray);
		printf("%6u types\n", ulPropTypeArray);
		printf("%6u guids\n", ulPropGuidArray);
		printf("%6u errors\n", g_ulErrorArray);
		printf("%6u smart view parsers\n", ulSmartViewParserTypeArray - 1);
		printf("\n");
	}
	printf("Usage:\n");
	printf("   MrMAPI -%s\n", g_Switches[switchHelp].szSwitch);
	printf("   MrMAPI [-%s] [-%s] [-%s] [-%s <type>] <property number>|<property name>\n",
		g_Switches[switchSearch].szSwitch, g_Switches[switchDispid].szSwitch, g_Switches[switchDecimal].szSwitch, g_Switches[switchType].szSwitch);
	printf("   MrMAPI -%s\n", g_Switches[switchGuid].szSwitch);
	printf("   MrMAPI -%s <error>\n", g_Switches[switchError].szSwitch);
	printf("   MrMAPI -%s <type> -%s <input file> [-%s] [-%s <output file>]\n",
		g_Switches[switchParser].szSwitch, g_Switches[switchInput].szSwitch, g_Switches[switchBinary].szSwitch, g_Switches[switchOutput].szSwitch);
	printf("   MrMAPI -%s <flag value> [-%s] [-%s] <property number>|<property name>\n",
		g_Switches[switchFlag].szSwitch, g_Switches[switchDispid].szSwitch, g_Switches[switchDecimal].szSwitch);
	printf("   MrMAPI -%s <flag name>\n",
		g_Switches[switchFlag].szSwitch);
	printf("   MrMAPI -%s [-%s <profile>] [-%s <folder>]\n", g_Switches[switchRule].szSwitch, g_Switches[switchProfile].szSwitch, g_Switches[switchFolder].szSwitch);
	printf("   MrMAPI -%s [-%s <profile>] [-%s <folder>]\n", g_Switches[switchAcl].szSwitch, g_Switches[switchProfile].szSwitch, g_Switches[switchFolder].szSwitch);
	printf("   MrMAPI -%s | -%s [-%s <profile>] [-%s <folder>] [-%s <output directory>]\n",
		g_Switches[switchContents].szSwitch, g_Switches[switchAssociatedContents].szSwitch, g_Switches[switchProfile].szSwitch, g_Switches[switchFolder].szSwitch, g_Switches[switchOutput].szSwitch);
	printf("          [-%s <subject>] [-%s <message class>] [-%s] [-%s] [-%s <count>] [-%s]\n",
		g_Switches[switchSubject].szSwitch, g_Switches[switchMessageClass].szSwitch, g_Switches[switchMSG].szSwitch, g_Switches[switchList].szSwitch, g_Switches[switchRecent].szSwitch, g_Switches[switchSkip].szSwitch);
	printf("   MrMAPI -%s [-%s <profile>] [-%s <folder>]\n", g_Switches[switchChildFolders].szSwitch, g_Switches[switchProfile].szSwitch, g_Switches[switchFolder].szSwitch);
	printf("   MrMAPI -%s -%s <path to input file> -%s <path to output file> [-%s]\n",
		g_Switches[switchXML].szSwitch, g_Switches[switchInput].szSwitch, g_Switches[switchOutput].szSwitch, g_Switches[switchSkip].szSwitch);
	printf("   MrMAPI -%s [fid] [-%s [mid]] [-%s <profile>]\n",
		g_Switches[switchFid].szSwitch, g_Switches[switchMid].szSwitch, g_Switches[switchProfile].szSwitch);
	printf("   MrMAPI [<property number>|<property name>] -%s [<store num>] [-%s <profile>]\n",
		g_Switches[switchStore].szSwitch, g_Switches[switchProfile].szSwitch);
	printf("   MrMAPI [<property number>|<property name>] -%s <folder> [-%s <profile>]\n",
		g_Switches[switchFolder].szSwitch, g_Switches[switchProfile].szSwitch);
	printf("   MrMAPI -%s -%s <folder> [-%s <profile>]\n",
		g_Switches[switchSize].szSwitch, g_Switches[switchFolder].szSwitch, g_Switches[switchProfile].szSwitch);
	printf("   MrMAPI -%s | -%s -%s <path to input file> -%s <path to output file> [-%s <conversion flags>]\n",
		g_Switches[switchMAPI].szSwitch, g_Switches[switchMIME].szSwitch, g_Switches[switchInput].szSwitch, g_Switches[switchOutput].szSwitch, g_Switches[switchCCSFFlags].szSwitch);
	printf("          [-%s] [-%s <Decimal number of characters>] [-%s <Decimal number indicating encoding>]\n",
		g_Switches[switchRFC822].szSwitch, g_Switches[switchWrap].szSwitch, g_Switches[switchEncoding].szSwitch);
	printf("          [-%s] [-%s] [-%s CodePage CharSetType CharSetApplyType]\n",
		g_Switches[switchAddressBook].szSwitch, g_Switches[switchUnicode].szSwitch, g_Switches[switchCharset].szSwitch);
	printf("   MrMAPI -%s -%s <path to input file>\n",
		g_Switches[switchPST].szSwitch, g_Switches[switchInput].szSwitch);
	printf("   MrMAPI -%s [<profile> [-%s <profilesection> [-%s]] -%s <output file>]\n",
		g_Switches[switchProfile].szSwitch, g_Switches[switchProfileSection].szSwitch, g_Switches[switchByteSwapped].szSwitch, g_Switches[switchOutput].szSwitch);
	printf("   MrMAPI -%s [<store num>] [-%s <profile>]\n", g_Switches[switchReceiveFolder].szSwitch, g_Switches[switchProfile].szSwitch);
	printf("   MrMAPI -%s -%s <folder> [-%s <profile>]\n",
		g_Switches[switchSearchState].szSwitch, g_Switches[switchFolder].szSwitch, g_Switches[switchProfile].szSwitch);

	if (bFull)
	{
		printf("\n");
		printf("All switches may be shortened if the intended switch is unambiguous.\n");
		printf("For example, -T may be used instead of -%s.\n", g_Switches[switchType].szSwitch);
	}
	printf("\n");
	printf("   Help:\n");
	printf("   -%s   Display expanded help.\n", g_Switches[switchHelp].szSwitch);
	if (bFull)
	{
		printf("\n");
		printf("   Property Tag Lookup:\n");
		printf("   -S   (or -%s) Perform substring search.\n", g_Switches[switchSearch].szSwitch);
		printf("           With no parameters prints all known properties.\n");
		printf("   -D   (or -%s) Search dispids.\n", g_Switches[switchDispid].szSwitch);
		printf("   -N   (or -%s) Number is in decimal. Ignored for non-numbers.\n", g_Switches[switchDecimal].szSwitch);
		printf("   -T   (or -%s) Print information on specified type.\n", g_Switches[switchType].szSwitch);
		printf("           With no parameters prints list of known types.\n");
		printf("           When combined with -S, restrict output to given type.\n");
		printf("   -G   (or -%s) Display list of known guids.\n", g_Switches[switchGuid].szSwitch);
		printf("\n");
		printf("   Flag Lookup:\n");
		printf("   -Fl  (or -%s) Look up flags for specified property.\n", g_Switches[switchFlag].szSwitch);
		printf("           May be combined with -D and -N switches, but all flag values must be in hex.\n");
		printf("   -Fl  (or -%s) Look up flag name and output its value.\n", g_Switches[switchFlag].szSwitch);
		printf("\n");
		printf("   Error Parsing:\n");
		printf("   -E   (or -%s) Map an error code to its name and vice versa.\n", g_Switches[switchError].szSwitch);
		printf("           May be combined with -S and -N switches.\n");
		printf("\n");
		printf("   Smart View Parsing:\n");
		printf("   -P   (or -%s) Parser type (number). See list below for supported parsers.\n", g_Switches[switchParser].szSwitch);
		printf("   -B   (or -%s) Input file is binary. Default is hex encoded text.\n", g_Switches[switchBinary].szSwitch);
		printf("\n");
		printf("   Rules Table:\n");
		printf("   -R   (or -%s) Output rules table. Profile optional.\n", g_Switches[switchRule].szSwitch);
		printf("\n");
		printf("   ACL Table:\n");
		printf("   -A   (or -%s) Output ACL table. Profile optional.\n", g_Switches[switchAcl].szSwitch);
		printf("\n");
		printf("   Contents Table:\n");
		printf("   -C   (or -%s) Output contents table. May be combined with -H. Profile optional.\n", g_Switches[switchContents].szSwitch);
		printf("   -H   (or -%s) Output associated contents table. May be combined with -C. Profile optional\n", g_Switches[switchAssociatedContents].szSwitch);
		printf("   -Su  (or -%s) Subject of messages to output.\n", g_Switches[switchSubject].szSwitch);
		printf("   -Me  (or -%s) Message class of messages to output.\n", g_Switches[switchMessageClass].szSwitch);
		printf("   -Ms  (or -%s) Output as .MSG instead of XML.\n", g_Switches[switchMSG].szSwitch);
		printf("   -L   (or -%s) List details to screen and do not output files.\n", g_Switches[switchList].szSwitch);
		printf("   -Re  (or -%s) Restrict output to the 'count' most recent messages.\n", g_Switches[switchRecent].szSwitch);
		printf("\n");
		printf("   Child Folders:\n");
		printf("   -Chi (or -%s) Display child folders of selected folder.\n", g_Switches[switchChildFolders].szSwitch);
		printf("\n");
		printf("   MSG File Properties\n");
		printf("   -X   (or -%s) Output properties of an MSG file as XML.\n", g_Switches[switchXML].szSwitch);
		printf("\n");
		printf("   MID/FID Lookup\n");
		printf("   -Fi  (or -%s) Folder ID (FID) to search for.\n", g_Switches[switchFid].szSwitch);
		printf("           If -%s is specified without a FID, search/display all folders\n", g_Switches[switchFid].szSwitch);
		printf("   -Mid (or -%s) Message ID (MID) to search for.\n", g_Switches[switchMid].szSwitch);
		printf("           If -%s is specified without a MID, display all messages in folders specified by the FID parameter.\n", g_Switches[switchMid].szSwitch);
		printf("\n");
		printf("   Store Properties\n");
		printf("   -St  (or -%s) Output properties of stores as XML.\n", g_Switches[switchStore].szSwitch);
		printf("           If store number is specified, outputs properties of a single store.\n");
		printf("           If a property is specified, outputs only that property.\n");
		printf("\n");
		printf("   Folder Properties\n");
		printf("   -F   (or -%s) Output properties of a folder as XML.\n", g_Switches[switchFolder].szSwitch);
		printf("           If a property is specified, outputs only that property.\n");
		printf("   -Size         Output size of a folder and all subfolders.\n");
		printf("           Use -%s to specify which folder to scan.\n", g_Switches[switchFolder].szSwitch);
		printf("   -SearchState  Output search folder state.\n");
		printf("           Use -%s to specify which folder to scan.\n", g_Switches[switchFolder].szSwitch);
		printf("\n");
		printf("   MAPI <-> MIME Conversion:\n");
		printf("   -Ma  (or -%s) Convert an EML file to MAPI format (MSG file).\n", g_Switches[switchMAPI].szSwitch);
		printf("   -Mi  (or -%s) Convert an MSG file to MIME format (EML file).\n", g_Switches[switchMIME].szSwitch);
		printf("   -I   (or -%s) Indicates the input file for conversion, either a MIME-formatted EML file or an MSG file.\n", g_Switches[switchInput].szSwitch);
		printf("   -O   (or -%s) Indicates the output file for the conversion.\n", g_Switches[switchOutput].szSwitch);
		printf("   -Cc  (or -%s) Indicates specific flags to pass to the converter.\n", g_Switches[switchCCSFFlags].szSwitch);
		printf("           Available values (these may be OR'ed together):\n");
		printf("              MIME -> MAPI:\n");
		printf("                CCSF_SMTP:        0x02\n");
		printf("                CCSF_INCLUDE_BCC: 0x20\n");
		printf("                CCSF_USE_RTF:     0x80\n");
		printf("              MAPI -> MIME:\n");
		printf("                CCSF_NOHEADERS:        0x00004\n");
		printf("                CCSF_USE_TNEF:         0x00010\n");
		printf("                CCSF_8BITHEADERS:      0x00040\n");
		printf("                CCSF_PLAIN_TEXT_ONLY:  0x01000\n");
		printf("                CCSF_NO_MSGID:         0x04000\n");
		printf("                CCSF_EMBEDDED_MESSAGE: 0x08000\n");
		printf("                CCSF_PRESERVE_SOURCE:  0x40000\n");
		printf("   -Rf  (or -%s) (MAPI->MIME only) Indicates the EML should be generated in RFC822 format.\n", g_Switches[switchRFC822].szSwitch);
		printf("           If not present, RFC1521 is used instead.\n");
		printf("   -W   (or -%s) (MAPI->MIME only) Indicates the maximum number of characters in each line in the\n", g_Switches[switchWrap].szSwitch);
		printf("           generated EML. Default value is 74. A value of 0 indicates no wrapping.\n");
		printf("   -En  (or -%s) (MAPI->MIME only) Indicates the encoding type to use. Supported values are:\n", g_Switches[switchEncoding].szSwitch);
		printf("              1 - Base64\n");
		printf("              2 - UUENCODE\n");
		printf("              3 - Quoted-Printable\n");
		printf("              4 - 7bit (DEFAULT)\n");
		printf("              5 - 8bit\n");
		printf("   -Ad  (or -%s) Pass MAPI Address Book into converter. Profile optional.\n", g_Switches[switchAddressBook].szSwitch);
		printf("   -U   (or -%s) (MIME->MAPI only) The resulting MSG file should be unicode.\n", g_Switches[switchUnicode].szSwitch);
		printf("   -Ch  (or -%s) (MIME->MAPI only) Character set - three required parameters:\n", g_Switches[switchCharset].szSwitch);
		printf("           CodePage - common values (others supported)\n");
		printf("              1252  - CP_USASCII      - Indicates the USASCII character set, Windows code page 1252\n");
		printf("              1200  - CP_UNICODE      - Indicates the Unicode character set, Windows code page 1200\n");
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
		printf("   -I   (or -%s) PST file to be analyzed.\n", g_Switches[switchInput].szSwitch);
		printf("\n");
		printf("   Profiles\n");
		printf("   -Pr  (or -%s) Output list of profiles\n", g_Switches[switchProfile].szSwitch);
		printf("           If a profile is specified, exports that profile.\n");
		printf("   -ProfileSection If specified, output specific profile section.\n");
		printf("   -B   (or -%s) If specified, profile section guid is byte swapped.\n", g_Switches[switchByteSwapped].szSwitch);
		printf("   -O   (or -%s) Indicates the output file for profile export.\n", g_Switches[switchOutput].szSwitch);
		printf("           Required if a profile is specified.\n");
		printf("\n");
		printf("   Receive Folder Table\n");
		printf("   -%s Displays Receive Folder Table for the specified store\n", g_Switches[switchReceiveFolder].szSwitch);
		printf("\n");
		printf("   Universal Options:\n");
		printf("   -I   (or -%s) Input file.\n", g_Switches[switchInput].szSwitch);
		printf("   -O   (or -%s) Output file or directory.\n", g_Switches[switchOutput].szSwitch);
		printf("   -F   (or -%s) Folder to scan. Default is Inbox. See list below for supported folders.\n", g_Switches[switchFolder].szSwitch);
		printf("           Folders may also be specified by path:\n");
		printf("              \"Top of Information Store\\Calendar\"\n");
		printf("           Path may be preceeded by entry IDs for special folders using @ notation:\n");
		printf("              \"@PR_IPM_SUBTREE_ENTRYID\\Calendar\"\n");
		printf("           Path may further be preceeded by store number using # notation, which may either use a store number:\n");
		printf("              \"#0\\@PR_IPM_SUBTREE_ENTRYID\\Calendar\"\n");
		printf("           Or an entry ID:\n");
		printf("              \"#00112233445566...778899AABBCC\\@PR_IPM_SUBTREE_ENTRYID\\Calendar\"\n");
		printf("           MrMAPI's special folder constants may also be used:\n");
		printf("              \"@12\\Calendar\"\n");
		printf("              \"@1\"\n");
		printf("   -Pr  (or -%s) Profile for MAPILogonEx.\n", g_Switches[switchProfile].szSwitch);
		printf("   -M   (or -%s) More properties. Tries harder to get stream properties. May take longer.\n", g_Switches[switchMoreProperties].szSwitch);
		printf("   -Sk  (or -%s) Skip embedded message attachments on export.\n", g_Switches[switchSkip].szSwitch);
		printf("   -No  (or -%s) No Addins. Don't load any add-ins.\n", g_Switches[switchNoAddins].szSwitch);
		printf("   -On  (or -%s) Online mode. Bypass cached mode.\n", g_Switches[switchOnline].szSwitch);
		printf("   -V   (or -%s) Verbose. Turn on all debug output.\n", g_Switches[switchVerbose].szSwitch);
		printf("\n");
		printf("   MAPI Implementation Options:\n");
		printf("   -%s MAPI Version to load - supported values\n", g_Switches[switchVersion].szSwitch);
		printf("           Supported values\n");
		printf("              0  - List all available MAPI binaries\n");
		printf("              1  - System MAPI\n");
		printf("              11 - Outlook 2003 (11)\n");
		printf("              12 - Outlook 2007 (12)\n");
		printf("              14 - Outlook 2010 (14)\n");
		printf("              15 - Outlook 2013 (15)\n");
		printf("           You can also pass a string, which will load the first MAPI whose path contains the string.\n");
		printf("\n");
		printf("Smart View Parsers:\n");
		// Print smart view options
		ULONG i = 1;
		for (i = 1; i < ulSmartViewParserTypeArray; i++)
		{
			_tprintf(_T("   %2u %ws\n"), i, SmartViewParserTypeArray[i].lpszName);
		}
		printf("\n");
		printf("Folders:\n");
		// Print Folders
		for (i = 1; i < NUM_DEFAULT_PROPS; i++)
		{
			printf("   %2u %hs\n", i, FolderNames[i]);
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

bool bSetMode(_In_ CmdMode* pMode, _In_ CmdMode TargetMode)
{
	if (pMode && ((cmdmodeUnknown == *pMode) || (TargetMode == *pMode)))
	{
		*pMode = TargetMode;
		return true;
	}
	return false;
}

struct OptParser
{
	__CommandLineSwitch Switch;
	CmdMode Mode;
	int MinArgs;
	int MaxArgs;
	ULONG ulOpt;
};

OptParser g_Parsers[] =
{
	{ switchHelp, cmdmodeHelp, 0, 0, OPT_INITMFC },
	{ switchVerbose, cmdmodeUnknown, 0, 0, OPT_VERBOSE | OPT_INITMFC },
	{ switchNoAddins, cmdmodeUnknown, 0, 0, OPT_NOADDINS },
	{ switchOnline, cmdmodeUnknown, 0, 0, OPT_ONLINE },
	{ switchSearch, cmdmodeUnknown, 0, 0, OPT_DOPARTIALSEARCH },
	{ switchDecimal, cmdmodeUnknown, 0, 0, OPT_DODECIMAL },
	{ switchFolder, cmdmodeUnknown, 1, 1, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_NEEDFOLDER | OPT_INITMFC },
	{ switchInput, cmdmodeUnknown, 1, 1, 0 },
	{ switchOutput, cmdmodeUnknown, 1, 1, 0 },
	{ switchProfile, cmdmodeUnknown, 0, 1, OPT_PROFILE },
	{ switchMoreProperties, cmdmodeUnknown, 0, 0, OPT_RETRYSTREAMPROPS },
	{ switchSkip, cmdmodeUnknown, 0, 0, OPT_SKIPATTACHMENTS },
	{ switchDispid, cmdmodePropTag, 0, 0, OPT_DODISPID },
	{ switchType, cmdmodePropTag, 0, 1, OPT_DOTYPE },
	{ switchFlag, cmdmodeUnknown, 1, 1, 0 }, // can't know until we parse the argument
	{ switchGuid, cmdmodeGuid, 0, 0, 0 },
	{ switchError, cmdmodeErr, 0, 0, 0 },
	{ switchParser, cmdmodeSmartView, 1, 1, OPT_INITMFC | OPT_NEEDINPUTFILE },
	{ switchBinary, cmdmodeSmartView, 0, 0, OPT_BINARYFILE },
	{ switchAcl, cmdmodeAcls, 0, 0, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDFOLDER },
	{ switchRule, cmdmodeRules, 0, 0, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDFOLDER },
	{ switchContents, cmdmodeContents, 0, 0, OPT_DOCONTENTS | OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC },
	{ switchAssociatedContents, cmdmodeContents, 0, 0, OPT_DOASSOCIATEDCONTENTS | OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC },
	{ switchSubject, cmdmodeContents, 0, 0, 0 },
	{ switchMSG, cmdmodeContents, 0, 0, OPT_MSG },
	{ switchList, cmdmodeContents, 0, 0, OPT_LIST },
	{ switchRecent, cmdmodeContents, 1, 1, 0 },
	{ switchXML, cmdmodeXML, 0, 0, OPT_NEEDMAPIINIT | OPT_INITMFC | OPT_NEEDINPUTFILE },
	{ switchFid, cmdmodeFidMid, 0, 1, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDSTORE },
	{ switchMid, cmdmodeFidMid, 0, 1, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC },
	{ switchStore, cmdmodeStoreProperties, 0, 1, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC },
	{ switchChildFolders, cmdmodeChildFolders, 0, 1, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDFOLDER },
	{ switchMAPI, cmdmodeMAPIMIME, 0, 0, OPT_NEEDMAPIINIT | OPT_INITMFC | OPT_NEEDINPUTFILE | OPT_NEEDOUTPUTFILE },
	{ switchMIME, cmdmodeMAPIMIME, 0, 0, OPT_NEEDMAPIINIT | OPT_INITMFC | OPT_NEEDINPUTFILE | OPT_NEEDOUTPUTFILE },
	{ switchCCSFFlags, cmdmodeMAPIMIME, 1, 1, 0 },
	{ switchRFC822, cmdmodeMAPIMIME, 1, 1, 0 },
	{ switchWrap, cmdmodeMAPIMIME, 1, 1, 0 },
	{ switchEncoding, cmdmodeMAPIMIME, 1, 1, 0 },
	{ switchCharset, cmdmodeMAPIMIME, 3, 3, 0 },
	{ switchAddressBook, cmdmodeMAPIMIME, 0, 0, OPT_NEEDMAPILOGON }, // special case which needs a logon
	{ switchUnicode, cmdmodeMAPIMIME, 0, 0, 0 },
	{ switchSize, cmdmodeFolderSize, 0, 0, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDFOLDER },
	{ switchPST, cmdmodePST, 0, 0, OPT_NEEDINPUTFILE },
	{ switchVersion, cmdmodeUnknown, 1, 1, 0 },
	{ switchProfileSection, cmdmodeProfile, 1, 1, OPT_PROFILE | OPT_NEEDMAPIINIT | OPT_INITMFC },
	{ switchByteSwapped, cmdmodeProfile, 0, 0, OPT_PROFILE | OPT_NEEDMAPIINIT | OPT_INITMFC },
	{ switchReceiveFolder, cmdmodeReceiveFolder, 0, 1, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_NEEDSTORE | OPT_INITMFC },
	{ switchSearchState, cmdmodeSearchState, 0, 1, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDFOLDER },
	{ switchNoSwitch, cmdmodeUnknown, 0, 0, 0 },
};

MYOPTIONS::MYOPTIONS()
{
	Mode = cmdmodeUnknown;
	ulOptions = 0;
	ulTypeNum = 0;
	ulSVParser = 0;
	ulStore = 0;
	ulFolder = 0;
	ulMAPIMIMEFlags = 0;
	ulConvertFlags = 0;
	ulWrapLines = 0;
	ulEncodingType = 0;
	ulCodePage = 0;
	ulFlagValue = 0;
	ulCount = 0;
	bByteSwapped = false;
	cSetType = CHARSET_BODY;
	cSetApplyType = CSET_APPLY_UNTAGGED;
	lpMAPISession = NULL;
	lpMDB = NULL;
	lpFolder = NULL;
}

OptParser* GetParser(__CommandLineSwitch Switch)
{
	int i = 0;
	for (i = 0; i < _countof(g_Parsers); i++)
	{
		if (Switch == g_Parsers[i].Switch) return &g_Parsers[i];
	}

	return NULL;
}

// Checks if szArg is an option, and if it is, returns which option it is
// We return the first match, so g_Switches should be ordered appropriately
__CommandLineSwitch ParseArgument(_In_z_ LPCSTR szArg)
{
	if (!szArg || !szArg[0]) return switchNoSwitch;

	ULONG i = 0;
	LPCSTR szSwitch = NULL;

	// Check if this is a switch at all
	switch (szArg[0])
	{
	case '-':
	case '/':
	case '\\':
		if (szArg[1] != 0) szSwitch = &szArg[1];
		break;
	default:
		return switchNoSwitch;
		break;
	}

	for (i = 0; i < g_ulSwitches; i++)
	{
		// If we have a match
		if (StrStrIA(g_Switches[i].szSwitch, szSwitch) == g_Switches[i].szSwitch)
		{
			return g_Switches[i].iSwitch;
		}
	}

	return switchUnknown;
}

// Parses command line arguments and fills out MYOPTIONS
bool ParseArgs(_In_ int argc, _In_count_(argc) char * argv[], _Out_ MYOPTIONS * pRunOpts)
{
	LPSTR szEndPtr = NULL;

	pRunOpts->ulTypeNum = ulNoMatch;
	pRunOpts->ulFolder = DEFAULT_INBOX;

	if (!pRunOpts) return false;
	if (1 == argc) return false;

	bool bHitError = false;

	for (int i = 1; i < argc; i++)
	{
		__CommandLineSwitch iSwitch = ParseArgument(argv[i]);
		OptParser* opt = GetParser(iSwitch);

		if (opt)
		{
			pRunOpts->ulOptions |= opt->ulOpt;
			if (cmdmodeUnknown != opt->Mode && cmdmodeHelp != pRunOpts->Mode)
			{
				if (!bSetMode(&pRunOpts->Mode, opt->Mode))
				{
					// resetting our mode here, switch to help
					pRunOpts->Mode = cmdmodeHelp;
					bHitError = true;
				}
			}

			// Make sure we have the minimum number of args
			// Commands with variable argument counts can special case themselves
			if (opt->MinArgs > 0)
			{
				int iArg = 0;
				for (iArg = 1; iArg <= opt->MinArgs; iArg++)
				{
					if (argc <= i + iArg || switchNoSwitch != ParseArgument(argv[i + iArg]))
					{
						// resetting our mode here, switch to help
						pRunOpts->Mode = cmdmodeHelp;
						bHitError = true;
					}
				}
			}
		}

		if (!bHitError) switch (iSwitch)
		{
			// Global flags
		case switchFolder:
			pRunOpts->ulFolder = strtoul(argv[i + 1], &szEndPtr, 10);
			if (!pRunOpts->ulFolder)
			{
				pRunOpts->lpszFolderPath = LPCSTRToWstring(argv[i + 1]);
				pRunOpts->ulFolder = DEFAULT_INBOX;
			}
			i++;
			break;
		case switchInput:
			pRunOpts->lpszInput = LPCSTRToWstring(argv[i + 1]);
			i++;
			break;
		case switchOutput:
			pRunOpts->lpszOutput = LPCSTRToWstring(argv[i + 1]);
			i++;
			break;
		case switchProfile:
			// If we have a next argument and it's not an option, parse it as a profile name
			if (i + 1 < argc && switchNoSwitch == ParseArgument(argv[i + 1]))
			{
				pRunOpts->lpszProfile = LPCSTRToWstring(argv[i + 1]);
				i++;
			}
			break;
		case switchProfileSection:
			pRunOpts->lpszProfileSection = LPCSTRToWstring(argv[i + 1]);
			i++;
			break;
		case switchByteSwapped:
			pRunOpts->bByteSwapped = true;
			break;
		case switchVersion:
			pRunOpts->lpszVersion = LPCSTRToWstring(argv[i + 1]);
			i++;
			break;
			// Proptag parsing
		case switchType:
			// If we have a next argument and it's not an option, parse it as a type
			if (i + 1 < argc && switchNoSwitch == ParseArgument(argv[i + 1]))
			{
				pRunOpts->ulTypeNum = PropTypeNameToPropTypeA(argv[i + 1]);
				i++;
			}
			break;
		case switchFlag:
			// If we have a next argument and it's not an option, parse it as a flag
			pRunOpts->lpszFlagName = LPCSTRToWstring(argv[i + 1]);
			pRunOpts->ulFlagValue = strtoul(argv[i + 1], &szEndPtr, 16);

			// Set mode based on whether the flag string was completely parsed as a number
			if (NULL == szEndPtr[0])
			{
				if (!bSetMode(&pRunOpts->Mode, cmdmodePropTag)) { bHitError = true; break; }
				pRunOpts->ulOptions |= OPT_DOFLAG;
			}
			else
			{
				if (!bSetMode(&pRunOpts->Mode, cmdmodeFlagSearch)) { bHitError = true; break; }
			}
			i++;
			break;
			// Smart View parsing
		case switchParser:
			pRunOpts->ulSVParser = strtoul(argv[i + 1], &szEndPtr, 10);
			i++;
			break;
			// Contents tables
		case switchSubject:
			pRunOpts->lpszSubject = LPCSTRToWstring(argv[i + 1]);
			i++;
			break;
		case switchMessageClass:
			pRunOpts->lpszMessageClass = LPCSTRToWstring(argv[i + 1]);
			i++;
			break;
		case switchRecent:
			pRunOpts->ulCount = strtoul(argv[i + 1], &szEndPtr, 10);
			i++;
			break;
			// FID / MID
		case switchFid:
			if (i + 1 < argc  && switchNoSwitch == ParseArgument(argv[i + 1]))
			{
				pRunOpts->lpszFid = LPCSTRToWstring(argv[i + 1]);
				i++;
			}
			break;
		case switchMid:
			if (i + 1 < argc  && switchNoSwitch == ParseArgument(argv[i + 1]))
			{
				pRunOpts->lpszMid = LPCSTRToWstring(argv[i + 1]);
				i++;
			}
			else
			{
				// We use the blank string to remember the -mid parameter was passed and save having an extra flag
				// TODO: Check if this works
				pRunOpts->lpszMid = L"";
			}
			break;
			// Store Properties / Receive Folder:
		case switchStore:
		case switchReceiveFolder:
			if (i + 1 < argc  && switchNoSwitch == ParseArgument(argv[i + 1]))
			{
				pRunOpts->ulStore = strtoul(argv[i + 1], &szEndPtr, 10);

				// If we parsed completely, this was a store number
				if (NULL == szEndPtr[0])
				{
					// Increment ulStore so we can use to distinguish an unset value
					pRunOpts->ulStore++;
					i++;
				}
				// Else it was a naked option - leave it on the stack
			}
			break;
			// MAPIMIME
		case switchMAPI:
			pRunOpts->ulMAPIMIMEFlags |= MAPIMIME_TOMAPI;
			break;
		case switchMIME:
			pRunOpts->ulMAPIMIMEFlags |= MAPIMIME_TOMIME;
			break;
		case switchCCSFFlags:
			pRunOpts->ulConvertFlags = strtoul(argv[i + 1], &szEndPtr, 10);
			i++;
			break;
		case switchRFC822:
			pRunOpts->ulMAPIMIMEFlags |= MAPIMIME_RFC822;
			break;
		case switchWrap:
			pRunOpts->ulWrapLines = strtoul(argv[i + 1], &szEndPtr, 10);
			pRunOpts->ulMAPIMIMEFlags |= MAPIMIME_WRAP;
			i++;
			break;
		case switchEncoding:
			pRunOpts->ulEncodingType = strtoul(argv[i + 1], &szEndPtr, 10);
			pRunOpts->ulMAPIMIMEFlags |= MAPIMIME_ENCODING;
			i++;
			break;
		case switchCharset:
			pRunOpts->ulCodePage = strtoul(argv[i + 1], &szEndPtr, 10);
			pRunOpts->cSetType = (CHARSETTYPE)strtoul(argv[i + 2], &szEndPtr, 10);
			if (pRunOpts->cSetType > CHARSET_WEB) { bHitError = true; break; }
			pRunOpts->cSetApplyType = (CSETAPPLYTYPE)strtoul(argv[i + 3], &szEndPtr, 10);
			if (pRunOpts->cSetApplyType > CSET_APPLY_TAG_ALL)  { bHitError = true; break; }
			pRunOpts->ulMAPIMIMEFlags |= MAPIMIME_CHARSET;
			i += 3;
			break;
		case switchAddressBook:
			pRunOpts->ulMAPIMIMEFlags |= MAPIMIME_ADDRESSBOOK;
			break;
		case switchUnicode:
			pRunOpts->ulMAPIMIMEFlags |= MAPIMIME_UNICODE;
			break;

		case switchNoSwitch:
			// naked option without a flag - we only allow one of these
			if (!pRunOpts->lpszUnswitchedOption.empty()) { bHitError = true; break; } // He's already got one, you see.
			pRunOpts->lpszUnswitchedOption = LPCSTRToWstring(argv[i]);
			break;
		case switchUnknown:
			// display help
			bHitError = true;
			break;
		}
	}

	if (bHitError)
	{
		pRunOpts->Mode = cmdmodeHelp;
		return false;
	}

	// Having processed the command line, we may not have determined a mode.
	// Some modes can be presumed by the switches we did see.

	// If we didn't get a mode set but we saw OPT_NEEDFOLDER, assume we're in folder dumping mode
	if (cmdmodeUnknown == pRunOpts->Mode && pRunOpts->ulOptions & OPT_NEEDFOLDER) pRunOpts->Mode = cmdmodeFolderProps;

	// If we didn't get a mode set, but we saw OPT_PROFILE, assume we're in profile dumping mode
	if (cmdmodeUnknown == pRunOpts->Mode && pRunOpts->ulOptions & OPT_PROFILE)
	{
		pRunOpts->Mode = cmdmodeProfile;
		pRunOpts->ulOptions |= OPT_NEEDMAPIINIT | OPT_INITMFC;
	}

	// If we didn't get a mode set, assume we're in prop tag mode
	if (cmdmodeUnknown == pRunOpts->Mode) pRunOpts->Mode = cmdmodePropTag;

	// If we weren't passed an output file/directory, remember the current directory
	if (pRunOpts->lpszOutput.empty() && pRunOpts->Mode != cmdmodeSmartView && pRunOpts->Mode != cmdmodeProfile)
	{
		char strPath[_MAX_PATH];
		GetCurrentDirectoryA(_MAX_PATH, strPath);

		pRunOpts->lpszOutput = LPCSTRToWstring(strPath);
	}

	// Validate that we have bare minimum to run
	if (pRunOpts->ulOptions & OPT_NEEDINPUTFILE && pRunOpts->lpszInput.empty()) return false;
	if (pRunOpts->ulOptions & OPT_NEEDOUTPUTFILE && pRunOpts->lpszOutput.empty()) return false;

	switch (pRunOpts->Mode)
	{
	case cmdmodePropTag:
		if (!(pRunOpts->ulOptions & OPT_DOTYPE) && !(pRunOpts->ulOptions & OPT_DOPARTIALSEARCH) && (pRunOpts->lpszUnswitchedOption.empty())) return false;
		if ((pRunOpts->ulOptions & OPT_DOPARTIALSEARCH) && (pRunOpts->ulOptions & OPT_DOTYPE) && ulNoMatch == pRunOpts->ulTypeNum) return false;
		if ((pRunOpts->ulOptions & OPT_DOFLAG) && ((pRunOpts->ulOptions & OPT_DOPARTIALSEARCH) || (pRunOpts->ulOptions & OPT_DOTYPE))) return false;
		break;
	case cmdmodeSmartView:
		if (!pRunOpts->ulSVParser) return false;
		break;
	case cmdmodeContents:
		if (!(pRunOpts->ulOptions & OPT_DOCONTENTS) && !(pRunOpts->ulOptions & OPT_DOASSOCIATEDCONTENTS)) return false;
		break;
	case cmdmodeMAPIMIME:
#define CHECKFLAG(__flag) ((pRunOpts->ulMAPIMIMEFlags & (__flag)) == (__flag))
		// Can't convert both ways at once
		if (CHECKFLAG(MAPIMIME_TOMAPI) && CHECKFLAG(MAPIMIME_TOMIME)) return false;

		// Make sure there's no MIME-only options specified in a MIME->MAPI conversion
		if (CHECKFLAG(MAPIMIME_TOMAPI) &&
			(CHECKFLAG(MAPIMIME_RFC822) ||
			CHECKFLAG(MAPIMIME_ENCODING) ||
			CHECKFLAG(MAPIMIME_WRAP)))
			return false;

		// Make sure there's no MAPI-only options specified in a MAPI->MIME conversion
		if (CHECKFLAG(MAPIMIME_TOMIME) &&
			(CHECKFLAG(MAPIMIME_CHARSET) ||
			CHECKFLAG(MAPIMIME_UNICODE)))
			return false;
		break;
	case cmdmodeProfile:
		if (!pRunOpts->lpszProfile.empty() && pRunOpts->lpszOutput.empty()) return false;
		if (pRunOpts->lpszProfile.empty() && !pRunOpts->lpszOutput.empty()) return false;
		if (!pRunOpts->lpszProfileSection.empty() && pRunOpts->lpszProfile.empty()) return false;
		break;
	default:
		break;
	}

	// Didn't fail - return true
	return true;
}

void PrintArgs(_In_ MYOPTIONS ProgOpts)
{
	DebugPrint(DBGGeneric, L"Mode = %d\n", ProgOpts.Mode);
	DebugPrint(DBGGeneric, L"ulOptions = 0x%08X\n", ProgOpts.ulOptions);
	DebugPrint(DBGGeneric, L"ulTypeNum = 0x%08X\n", ProgOpts.ulTypeNum);
	if (!ProgOpts.lpszUnswitchedOption.empty()) DebugPrint(DBGGeneric, L"lpszUnswitchedOption = %ws\n", ProgOpts.lpszUnswitchedOption.c_str());
	if (!ProgOpts.lpszFlagName.empty()) DebugPrint(DBGGeneric, L"lpszFlagName = %ws\n", ProgOpts.lpszFlagName.c_str());
	if (!ProgOpts.lpszFolderPath.empty()) DebugPrint(DBGGeneric, L"lpszFolderPath = %ws\n", ProgOpts.lpszFolderPath.c_str());
	if (!ProgOpts.lpszInput.empty()) DebugPrint(DBGGeneric, L"lpszInput = %ws\n", ProgOpts.lpszInput.c_str());
	if (!ProgOpts.lpszMessageClass.empty()) DebugPrint(DBGGeneric, L"lpszMessageClass = %ws\n", ProgOpts.lpszMessageClass.c_str());
	if (!ProgOpts.lpszMid.empty()) DebugPrint(DBGGeneric, L"lpszMid = %ws\n", ProgOpts.lpszMid.c_str());
	if (!ProgOpts.lpszOutput.empty()) DebugPrint(DBGGeneric, L"lpszOutput = %ws\n", ProgOpts.lpszOutput.c_str());
	if (!ProgOpts.lpszProfile.empty()) DebugPrint(DBGGeneric, L"lpszProfile = %ws\n", ProgOpts.lpszProfile.c_str());
	if (!ProgOpts.lpszProfileSection.empty()) DebugPrint(DBGGeneric, L"lpszProfileSection = %ws\n", ProgOpts.lpszProfileSection.c_str());
	if (!ProgOpts.lpszSubject.empty()) DebugPrint(DBGGeneric, L"lpszSubject = %ws\n", ProgOpts.lpszSubject.c_str());
	if (!ProgOpts.lpszVersion.empty()) DebugPrint(DBGGeneric, L"lpszVersion = %ws\n", ProgOpts.lpszVersion.c_str());
}

// Returns true if we've done everything we need to do and can exit the program.
// Returns false to continue work.
bool LoadMAPIVersion(wstring lpszVersion)
{
	// Load DLLS and get functions from them
	ImportProcs();
	DebugPrint(DBGGeneric, L"LoadMAPIVersion(%ws)\n", lpszVersion);

	LPWSTR szPath = NULL;

	LPWSTR szEndPtr = NULL;
	ULONG ulVersion = wcstoul(lpszVersion.c_str(), &szEndPtr, 10);
	MAPIPathIterator* mpi = new MAPIPathIterator(true);
	if (mpi)
	{
		if (szEndPtr[0])
		{
			DebugPrint(DBGGeneric, L"Got a string\n");

			wstringToLower(lpszVersion);
			for (;;)
			{
				szPath = mpi->GetNextMAPIPath();
				if (!szPath) break;
				_wcslwr(szPath);

				if (wcsstr(szPath, lpszVersion.c_str()))
				{
					break;
				}

				delete[] szPath;
				szPath = NULL;
			}
		}
		else if (0 == ulVersion)
		{
			DebugPrint(DBGGeneric, L"Listing MAPI\n");
			for (;;)
			{
				szPath = mpi->GetNextMAPIPath();
				if (!szPath) break;
				_wcslwr(szPath);

				printf("MAPI path: %ws\n", szPath);
				delete[] szPath;
				szPath = NULL;
			}
			return true;
		}
		else
		{
			DebugPrint(DBGGeneric, L"Got a number %u\n", ulVersion);
			switch (ulVersion)
			{
			case 1: // system
				szPath = mpi->GetMAPISystemDir();
				break;
			case 11: // Outlook 2003 (11)
				szPath = mpi->GetInstalledOutlookMAPI(oqcOffice11);
				break;
			case 12: // Outlook 2007 (12)
				szPath = mpi->GetInstalledOutlookMAPI(oqcOffice12);
				break;
			case 14: // Outlook 2010 (14)
				szPath = mpi->GetInstalledOutlookMAPI(oqcOffice14);
				break;
			case 15: // Outlook 2013 (15)
				szPath = mpi->GetInstalledOutlookMAPI(oqcOffice15);
				break;
			}
		}
	}

	if (szPath)
	{
		DebugPrint(DBGGeneric, L"Found MAPI path %ws\n", szPath);
		HMODULE hMAPI = NULL;
		HRESULT hRes = S_OK;
		WC_D(hMAPI, MyLoadLibraryW(szPath));
		SetMAPIHandle(hMAPI);
		delete[] szPath;
	}

	delete mpi;
	return false;
}

void main(_In_ int argc, _In_count_(argc) char * argv[])
{
	HRESULT hRes = S_OK;
	bool bMAPIInit = false;

	SetDllDirectory("");
	MyHeapSetInformation(NULL, HeapEnableTerminationOnCorruption, NULL, 0);

	// Set up our property arrays or nothing works
	MergeAddInArrays();

	RegKeys[regkeyDO_SMART_VIEW].ulCurDWORD = 1;
	RegKeys[regkeyUSE_GETPROPLIST].ulCurDWORD = 1;
	RegKeys[regkeyPARSED_NAMED_PROPS].ulCurDWORD = 1;
	RegKeys[regkeyCACHE_NAME_DPROPS].ulCurDWORD = 1;

	MYOPTIONS ProgOpts;
	bool bGoodCommandLine = ParseArgs(argc, argv, &ProgOpts);

	// Must be first after ParseArgs
	if (ProgOpts.ulOptions & OPT_INITMFC)
	{
		InitMFC();
	}

	if (ProgOpts.ulOptions & OPT_VERBOSE)
	{
		RegKeys[regkeyDEBUG_TAG].ulCurDWORD = 0xFFFFFFFF;
		PrintArgs(ProgOpts);
	}

	if (!(ProgOpts.ulOptions & OPT_NOADDINS))
	{
		RegKeys[regkeyLOADADDINS].ulCurDWORD = true;
		LoadAddIns();
	}

	if (!ProgOpts.lpszVersion.empty())
	{
		if (LoadMAPIVersion(ProgOpts.lpszVersion)) return;
	}

	if (cmdmodeHelp == ProgOpts.Mode || !bGoodCommandLine)
	{
		DisplayUsage(cmdmodeHelp == ProgOpts.Mode || bGoodCommandLine);
	}
	else
	{
		if (ProgOpts.ulOptions & OPT_ONLINE)
		{
			RegKeys[regKeyMAPI_NO_CACHE].ulCurDWORD = true;
			RegKeys[regkeyMDB_ONLINE].ulCurDWORD = true;
		}

		// Log on to MAPI if needed
		if (ProgOpts.ulOptions & OPT_NEEDMAPIINIT)
		{
			WC_MAPI(MAPIInitialize(NULL));
			if (FAILED(hRes))
			{
				printf("Error initializing MAPI: 0x%08x\n", hRes);
			}
			else
			{
				bMAPIInit = true;
			}
		}

		if (bMAPIInit && ProgOpts.ulOptions & OPT_NEEDMAPILOGON)
		{
			WC_H(MrMAPILogonEx(ProgOpts.lpszProfile, &ProgOpts.lpMAPISession));
			if (FAILED(hRes)) printf("MAPILogonEx returned an error: 0x%08x\n", hRes);
		}

		// If they need a folder get it and store at the same time from the folder id
		if (ProgOpts.lpMAPISession && ProgOpts.ulOptions & OPT_NEEDFOLDER)
		{
			WC_H(HrMAPIOpenStoreAndFolder(ProgOpts.lpMAPISession, ProgOpts.ulFolder, ProgOpts.lpszFolderPath, &ProgOpts.lpMDB, &ProgOpts.lpFolder));
			if (FAILED(hRes)) printf("HrMAPIOpenStoreAndFolder returned an error: 0x%08x\n", hRes);
		}
		else if (ProgOpts.lpMAPISession && ProgOpts.ulOptions & OPT_NEEDSTORE)
		{
			// They asked us for a store, if they passed a store index give them that one
			if (ProgOpts.ulStore != 0)
			{
				// Decrement by one here on the index since we incremented during parameter parsing
				// This is so zero indicate they did not specify a store
				WC_H(OpenStore(ProgOpts.lpMAPISession, ProgOpts.ulStore - 1, &ProgOpts.lpMDB));
				if (FAILED(hRes)) printf("OpenStore returned an error: 0x%08x\n", hRes);
			}
			else
			{
				// If they needed a store but didn't specify, get the default one
				WC_H(OpenExchangeOrDefaultMessageStore(ProgOpts.lpMAPISession, &ProgOpts.lpMDB));
				if (FAILED(hRes)) printf("OpenExchangeOrDefaultMessageStore returned an error: 0x%08x\n", hRes);
			}
		}

		switch (ProgOpts.Mode)
		{
		case cmdmodePropTag:
			DoPropTags(ProgOpts);
			break;
		case cmdmodeGuid:
			DoGUIDs(ProgOpts);
			break;
		case cmdmodeSmartView:
			DoSmartView(ProgOpts);
			break;
		case cmdmodeAcls:
			DoAcls(ProgOpts);
			break;
		case cmdmodeRules:
			DoRules(ProgOpts);
			break;
		case cmdmodeErr:
			DoErrorParse(ProgOpts);
			break;
		case cmdmodeContents:
			DoContents(ProgOpts);
			break;
		case cmdmodeXML:
			DoMSG(ProgOpts);
			break;
		case cmdmodeFidMid:
			DoFidMid(ProgOpts);
			break;
		case cmdmodeStoreProperties:
			DoStore(ProgOpts);
			break;
		case cmdmodeMAPIMIME:
			DoMAPIMIME(ProgOpts);
			break;
		case cmdmodeChildFolders:
			DoChildFolders(ProgOpts);
			break;
		case cmdmodeFlagSearch:
			DoFlagSearch(ProgOpts);
			break;
		case cmdmodeFolderProps:
			DoFolderProps(ProgOpts);
			break;
		case cmdmodeFolderSize:
			DoFolderSize(ProgOpts);
			break;
		case cmdmodePST:
			DoPST(ProgOpts);
			break;
		case cmdmodeProfile:
			DoProfile(ProgOpts);
			break;
		case cmdmodeReceiveFolder:
			DoReceiveFolder(ProgOpts);
			break;
		case cmdmodeSearchState:
			DoSearchState(ProgOpts);
			break;
		}
	}

	UninitializeNamedPropCache();

	if (bMAPIInit)
	{
		if (ProgOpts.lpFolder) ProgOpts.lpFolder->Release();
		if (ProgOpts.lpMDB) ProgOpts.lpMDB->Release();
		if (ProgOpts.lpMAPISession) ProgOpts.lpMAPISession->Release();
		MAPIUninitialize();
	}

	if (!(ProgOpts.ulOptions & OPT_NOADDINS))
	{
		UnloadAddIns();
	}
}