#include "stdafx.h"

#include "MrMAPI.h"
#include "MAPIFunctions.h"
#include "InterpretProp2.h"
#include "SmartView.h"
#include "MMAcls.h"
#include "MMContents.h"
#include "MMErr.h"
#include "MMFidMid.h"
#include "MMFolder.h"
#include "MMPropTag.h"
#include "MMRules.h"
#include "MMSmartView.h"
#include "MMMapiMime.h"
#include <shlwapi.h>
#include "ImportProcs.h"
#include "MAPIStoreFunctions.h"

// Initialize MFC for LoadString support later on
void InitMFC()
{
#pragma warning(push)
#pragma warning(disable:6309)
#pragma warning(disable:6387)
	if (!AfxWinInit(::GetModuleHandle(NULL), NULL, ::GetCommandLine(), 0)) return;
#pragma warning(pop)
} // InitMFC

_Check_return_ HRESULT MrMAPILogonEx(_In_opt_z_ LPCWSTR lpszProfile, _Deref_out_opt_ LPMAPISESSION* lppSession)
{
	HRESULT hRes = S_OK;
	ULONG ulFlags = MAPI_EXTENDED | MAPI_NO_MAIL | MAPI_UNICODE | MAPI_NEW_SESSION;
	if (!lpszProfile) ulFlags |= MAPI_USE_DEFAULT;

	WC_H(MAPILogonEx(NULL, (LPTSTR) lpszProfile, NULL,
		ulFlags,
		lppSession));
	return hRes;
} // MrMAPILogonEx

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
		WC_H(OpenDefaultMessageStore(lpMAPISession,&lpMDB));
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
	switchFlag,               // '-flag'
};

struct COMMANDLINE_SWITCH
{
	__CommandLineSwitch	iSwitch;
	LPCSTR szSwitch;
};

// All entries before the aliases must be in the
// same order as the __CommandLineSwitch enum.
COMMANDLINE_SWITCH g_Switches[] =
{
	{switchNoSwitch,           ""},
	{switchUnknown,            ""},
	{switchHelp,               "?"},
	{switchVerbose,            "Verbose"},
	{switchSearch,             "Search"},
	{switchDecimal,            "Number"},
	{switchFolder,             "Folder"},
	{switchOutput,             "Output"},
	{switchDispid,             "Dispids"},
	{switchType,               "Type"},
	{switchGuid,               "Guids"},
	{switchError,              "Error"},
	{switchParser,             "ParserType"},
	{switchInput,              "Input"},
	{switchBinary,             "Binary"},
	{switchAcl,                "Acl"},
	{switchRule,               "Rules"},
	{switchContents,           "Contents"},
	{switchAssociatedContents, "HiddenContents"},
	{switchMoreProperties,     "MoreProperties"},
	{switchNoAddins,           "NoAddins"},
	{switchOnline,             "Online"},
	{switchMAPI,               "MAPI"},
	{switchMIME,               "MIME"},
	{switchCCSFFlags,          "CCSFFlags"},
	{switchRFC822,             "RFC822"},
	{switchWrap,               "Wrap"},
	{switchEncoding,           "Encoding"},
	{switchCharset,            "Charset"},
	{switchAddressBook,        "AddressBook"},
	{switchUnicode,            "Unicode"},
	{switchProfile,            "Profile"},
	{switchXML,                "XML"},
	{switchSubject,            "Subject"},
	{switchMessageClass,       "MessageClass"},
	{switchMSG,                "MSG"},
	{switchList,               "List"},
	{switchChildFolders,       "ChildFolders"},
	{switchFid,                "FID"},
	{switchMid,                "MID"},
	{switchFlag,               "Flag"},
// If we want to add aliases for any switches, add them here
};
ULONG g_ulSwitches = _countof(g_Switches);

extern LPADDIN g_lpMyAddins;

void DisplayUsage(BOOL bFull)
{
	InitMFC();
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
				printf("   %ws\n",lpCurAddIn->szName);
				lpCurAddIn = lpCurAddIn->lpNextAddIn;
			}
		}
		printf("MrMAPI currently knows:\n");
		printf("%6d property tags\n",ulPropTagArray);
		printf("%6d dispids\n",ulNameIDArray);
		printf("%6d types\n",ulPropTypeArray);
		printf("%6d guids\n",ulPropGuidArray);
		printf("%6d errors\n",g_ulErrorArray);
		printf("%6d smart view parsers\n",g_cuidParsingTypes-1);
		printf("\n");
	}
	printf("Usage:\n");
	printf("   MrMAPI -%s\n",g_Switches[switchHelp].szSwitch);
	printf("   MrMAPI [-%s] [-%s] [-%s] [-%s <type>] <property number>|<property name>\n",
		g_Switches[switchSearch].szSwitch,g_Switches[switchDispid].szSwitch,g_Switches[switchDecimal].szSwitch,g_Switches[switchType].szSwitch);
	printf("   MrMAPI -%s\n",g_Switches[switchGuid].szSwitch);
	printf("   MrMAPI -%s <error>\n",g_Switches[switchError].szSwitch);
	printf("   MrMAPI -%s <type> -%s <input file> [-%s] [-%s <output file>]\n",
		g_Switches[switchParser].szSwitch,g_Switches[switchInput].szSwitch,g_Switches[switchBinary].szSwitch,g_Switches[switchOutput].szSwitch);
	printf("   MrMAPI -%s <flag value> [-%s] [-%s] <property number>|<property name>\n",
		g_Switches[switchFlag].szSwitch,g_Switches[switchDispid].szSwitch,g_Switches[switchDecimal].szSwitch);
	printf("   MrMAPI -%s [-%s <profile>] [-%s <folder>]\n",g_Switches[switchRule].szSwitch,g_Switches[switchProfile].szSwitch,g_Switches[switchFolder].szSwitch);
	printf("   MrMAPI -%s [-%s <profile>] [-%s <folder>]\n",g_Switches[switchAcl].szSwitch,g_Switches[switchProfile].szSwitch,g_Switches[switchFolder].szSwitch);
	printf("   MrMAPI [-%s | -%s] [-%s <profile>] [-%s <folder>] [-%s <output directory>]\n",
		g_Switches[switchContents].szSwitch,g_Switches[switchAssociatedContents].szSwitch,g_Switches[switchProfile].szSwitch,g_Switches[switchFolder].szSwitch,g_Switches[switchOutput].szSwitch);
	printf("          [-%s <subject>] [-%s <message class>] [-%s] [-%s]\n",
		g_Switches[switchSubject].szSwitch, g_Switches[switchMessageClass].szSwitch, g_Switches[switchMSG].szSwitch, g_Switches[switchList].szSwitch);
	printf("   MrMAPI -%s [-%s <profile>] [-%s <folder>]\n",g_Switches[switchChildFolders].szSwitch,g_Switches[switchProfile].szSwitch,g_Switches[switchFolder].szSwitch);
	printf("   MrMAPI -%s -%s <path to input file> -%s <path to output file>\n",
		g_Switches[switchXML].szSwitch,g_Switches[switchInput].szSwitch,g_Switches[switchOutput].szSwitch);
	printf("   MrMAPI -%s [fid] [-%s [mid]] [-%s <profile>]\n",
		g_Switches[switchFid].szSwitch,g_Switches[switchMid].szSwitch,g_Switches[switchProfile].szSwitch);
	printf("   MrMAPI -%s | -%s -%s <path to input file> -%s <path to output file> [-%s <conversion flags>]\n",
		g_Switches[switchMAPI].szSwitch,g_Switches[switchMIME].szSwitch,g_Switches[switchInput].szSwitch,g_Switches[switchOutput].szSwitch,g_Switches[switchCCSFFlags].szSwitch);
	printf("          [-%s] [-%s <Decimal number of characters>] [-%s <Decimal number indicating encoding>]\n",
		g_Switches[switchRFC822].szSwitch,g_Switches[switchWrap].szSwitch,g_Switches[switchEncoding].szSwitch);
	printf("          [-%s] [-%s] [-%s CodePage CharSetType CharSetApplyType]\n",
		g_Switches[switchAddressBook].szSwitch,g_Switches[switchUnicode].szSwitch,g_Switches[switchCharset].szSwitch);
	if (bFull)
	{
		printf("\n");
		printf("All switches may be shortened if the intended switch is unambiguous.\n");
		printf("For example, -T may be used instead of -%s.\n",g_Switches[switchType].szSwitch);
	}
	printf("\n");
	printf("   Help:\n");
	printf("   -%s   Display expanded help.\n",g_Switches[switchHelp].szSwitch);
	if (bFull)
	{
		printf("\n");
		printf("   Property Tag Lookup:\n");
		printf("   -S   (or -%s) Perform substring search.\n",g_Switches[switchSearch].szSwitch);
		printf("           With no parameters prints all known properties.\n");
		printf("   -D   (or -%s) Search dispids.\n",g_Switches[switchDispid].szSwitch);
		printf("   -N   (or -%s) Number is in decimal. Ignored for non-numbers.\n",g_Switches[switchDecimal].szSwitch);
		printf("   -T   (or -%s) Print information on specified type.\n",g_Switches[switchType].szSwitch);
		printf("           With no parameters prints list of known types.\n");
		printf("           When combined with -S, restrict output to given type.\n");
		printf("   -G   (or -%s) Display list of known guids.\n",g_Switches[switchGuid].szSwitch);
		printf("\n");
		printf("   Flag Lookup:\n");
		printf("   -Fl  (or -%s) Look up flags for specified property.\n",g_Switches[switchFlag].szSwitch);
		printf("           May be combined with -D and -N switches, but all flag values must be in hex.\n");
		printf("\n");
		printf("   Error Parsing:\n");
		printf("   -E   (or -%s) Map an error code to its name and vice versa.\n",g_Switches[switchError].szSwitch);
		printf("           May be combined with -S and -N switches.\n");
		printf("\n");
		printf("   Smart View Parsing:\n");
		printf("   -P   (or -%s) Parser type (number). See list below for supported parsers.\n",g_Switches[switchParser].szSwitch);
		printf("   -B   (or -%s) Input file is binary. Default is hex encoded text.\n",g_Switches[switchBinary].szSwitch);
		printf("\n");
		printf("   Rules Table:\n");
		printf("   -R   (or -%s) Output rules table. Profile optional.\n",g_Switches[switchRule].szSwitch);
		printf("\n");
		printf("   ACL Table:\n");
		printf("   -A   (or -%s) Output ACL table. Profile optional.\n",g_Switches[switchAcl].szSwitch);
		printf("\n");
		printf("   Contents Table:\n");
		printf("   -C   (or -%s) Output contents table. May be combined with -H. Profile optional.\n",g_Switches[switchContents].szSwitch);
		printf("   -H   (or -%s) Output associated contents table. May be combined with -C. Profile optional\n",g_Switches[switchAssociatedContents].szSwitch);
		printf("   -Su  (or -%s) Subject of messages to output.\n",g_Switches[switchSubject].szSwitch);
		printf("   -Me  (or -%s) Message class of messages to output.\n",g_Switches[switchMessageClass].szSwitch);
		printf("   -Ms  (or -%s) Output as .MSG instead of XML.\n",g_Switches[switchMSG].szSwitch);
		printf("   -L   (or -%s) List details to screen and do not output files.\n",g_Switches[switchList].szSwitch);
		printf("\n");
		printf("   Child Folders:\n");
		printf("   -Chi (or -%s) Display child folders of selected folder.\n",g_Switches[switchChildFolders].szSwitch);
		printf("\n");
		printf("   MSG File Properties\n");
		printf("   -X   (or -%s) Output properties of an MSG file as XML.\n",g_Switches[switchXML].szSwitch);
		printf("\n");
		printf("   MID/FID Lookup\n");
		printf("   -Fi  (or -%s) Folder ID (FID) to search for.\n",g_Switches[switchFid].szSwitch);
		printf("           If -%s is specified without a FID, search/display all folders\n",g_Switches[switchFid].szSwitch);
		printf("   -Mid (or -%s) Message ID (MID) to search for.\n",g_Switches[switchMid].szSwitch);
		printf("           If -%s is specified without a MID, display all messages in folders specified by the FID parameter.\n",g_Switches[switchMid].szSwitch);
		printf("\n");
		printf("   MAPI <-> MIME Conversion:\n");
		printf("   -Ma  (or -%s) Convert an EML file to MAPI format (MSG file).\n",g_Switches[switchMAPI].szSwitch);
		printf("   -Mi  (or -%s) Convert an MSG file to MIME format (EML file).\n",g_Switches[switchMIME].szSwitch);
		printf("   -I   (or -%s) Indicates the input file for conversion, either a MIME-formatted EML file or an MSG file.\n",g_Switches[switchInput].szSwitch);
		printf("   -O   (or -%s) Indicates the output file for the convertion.\n",g_Switches[switchOutput].szSwitch);
		printf("   -Cc  (or -%s) Indicates specific flags to pass to the converter.\n",g_Switches[switchCCSFFlags].szSwitch);
		printf("           Available values (these may be OR'ed together):\n");
		printf("              MIME -> MAPI:\n");
		printf("                CCSF_SMTP:        0x02\n");
		printf("                CCSF_INCLUDE_BCC: 0x20\n");
		printf("                CCSF_USE_RTF:     0x80\n");
		printf("              MAPI -> MIME:\n");
		printf("                CCSF_NOHEADERS:        0x0004\n");
		printf("                CCSF_USE_TNEF:         0x0010\n");
		printf("                CCSF_8BITHEADERS:      0x0040\n");
		printf("                CCSF_PLAIN_TEXT_ONLY:  0x1000\n");
		printf("                CCSF_NO_MSGID:         0x4000\n");
		printf("                CCSF_EMBEDDED_MESSAGE: 0x8000\n");
		printf("   -Rf  (or -%s) (MAPI->MIME only) Indicates the EML should be generated in RFC822 format.\n",g_Switches[switchRFC822].szSwitch);
		printf("           If not present, RFC1521 is used instead.\n");
		printf("   -W   (or -%s) (MAPI->MIME only) Indicates the maximum number of characters in each line in the\n",g_Switches[switchWrap].szSwitch);
		printf("           generated EML. Default value is 74. A value of 0 indicates no wrapping.\n");
		printf("   -En  (or -%s) (MAPI->MIME only) Indicates the encoding type to use. Supported values are:\n",g_Switches[switchEncoding].szSwitch);
		printf("              1 - Base64\n");
		printf("              2 - UUENCODE\n");
		printf("              3 - Quoted-Printable\n");
		printf("              4 - 7bit (DEFAULT)\n");
		printf("              5 - 8bit\n");
		printf("   -Ad  (or -%s) Pass MAPI Address Book into converter. Profile optional.\n",g_Switches[switchAddressBook].szSwitch);
		printf("   -U   (or -%s) (MIME->MAPI only) The resulting MSG file should be unicode.\n",g_Switches[switchUnicode].szSwitch);
		printf("   -Ch  (or -%s) (MIME->MAPI only) Character set - three required parameters:\n",g_Switches[switchCharset].szSwitch);
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
		printf("   Universal Options:\n");
		printf("   -I   (or -%s) Input file.\n",g_Switches[switchInput].szSwitch);
		printf("   -O   (or -%s) Output file or directory.\n",g_Switches[switchOutput].szSwitch);
		printf("   -F   (or -%s) Folder to scan. Default is Inbox. See list below for supported folders.\n",g_Switches[switchFolder].szSwitch);
		printf("           Folders may also be specified by path:\n");
		printf("              \"Top of Information Store\\Calendar\"\n");
		printf("           Path may be preceeded by entry IDs for special folders using @ notation:\n");
		printf("              \"@PR_IPM_SUBTREE_ENTRYID\\Calendar\"\n");
		printf("           MrMAPI's special folder constants may also be used:\n");
		printf("              \"@12\\Calendar\"\n");
		printf("              \"@1\"\n");
		printf("   -Pr  (or -%s) Profile for MAPILogonEx.\n",g_Switches[switchProfile].szSwitch);
		printf("   -M   (or -%s) More properties. Tries harder to get stream properties. May take longer.\n",g_Switches[switchMoreProperties].szSwitch);
		printf("   -No  (or -%s) No Addins. Don't load any add-ins.\n",g_Switches[switchNoAddins].szSwitch);
		printf("   -On  (or -%s) Online mode. Bypass cached mode.\n",g_Switches[switchOnline].szSwitch);
		printf("   -V   (or -%s) Verbose. Turn on all debug output.\n",g_Switches[switchVerbose].szSwitch);
		printf("\n");
		printf("Smart View Parsers:\n");
		// Print smart view options
		ULONG i = 1;
		for (i = 1 ; i < g_cuidParsingTypes ; i++)
		{
			_tprintf(_T("   %2d %ws\n"),i,g_uidParsingTypes[i].lpszName);
		}
		printf("\n");
		printf("Folders:\n");
		// Print Folders
		for (i = 1 ; i < NUM_DEFAULT_PROPS ; i++)
		{
			printf("   %2d %hs\n",i,FolderNames[i]);
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
} // DisplayUsage

bool bSetMode(_In_ CmdMode* pMode, _In_ CmdMode TargetMode)
{
	if (pMode && ((cmdmodeUnknown == *pMode) || (TargetMode == *pMode)))
	{
		*pMode = TargetMode;
		return true;
	}
	return false;
} // bSetMode

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

	for (i = 0 ; i < g_ulSwitches ; i++)
	{
		// If we have a match
		if (StrStrIA(g_Switches[i].szSwitch,szSwitch) == g_Switches[i].szSwitch)
		{
			return g_Switches[i].iSwitch;
		}
	}
	return switchUnknown;
} // ParseArgument

// Parses command line arguments and fills out MYOPTIONS
bool ParseArgs(_In_ int argc, _In_count_(argc) char * argv[], _Out_ MYOPTIONS * pRunOpts)
{
	HRESULT hRes = S_OK;
	LPSTR szEndPtr = NULL;

	// Clear our options list
	ZeroMemory(pRunOpts,sizeof(MYOPTIONS));

	pRunOpts->ulTypeNum = ulNoMatch;
	pRunOpts->ulFolder = DEFAULT_INBOX;

	if (!pRunOpts) return false;
	if (1 == argc) return false;

	for (int i = 1; i < argc; i++)
	{
		switch (ParseArgument(argv[i]))
		{
		// Global flags
		case switchHelp:
			pRunOpts->bHelp = true;
			break;
		case switchVerbose:
			pRunOpts->bVerbose = true;
			break;
		case switchNoAddins:
			pRunOpts->bNoAddins = true;
			break;
		case switchOnline:
			pRunOpts->bOnline = true;
			break;
		case switchSearch:
			pRunOpts->bDoPartialSearch = true;
			break;
		case switchDecimal:
			pRunOpts->bDoDecimal = true;
			break;
		case switchFolder:
			if (argc <= i+1 || switchNoSwitch != ParseArgument(argv[i+1])) return false;
			pRunOpts->ulFolder = strtoul(argv[i+1],&szEndPtr,10);
			if (!pRunOpts->ulFolder)
			{
				EC_H(AnsiToUnicode(argv[i+1],&pRunOpts->lpszFolderPath));
			}
			i++;
			break;
		case switchInput:
			if (argc <= i+1 || switchNoSwitch != ParseArgument(argv[i+1])) return false;
			EC_H(AnsiToUnicode(argv[i+1],&pRunOpts->lpszInput));
			i++;
			break;
		case switchOutput:
			if (argc <= i+1 || switchNoSwitch != ParseArgument(argv[i+1])) return false;
			EC_H(AnsiToUnicode(argv[i+1],&pRunOpts->lpszOutput));
			i++;
			break;
		case switchProfile:
			if (argc <= i+1 || switchNoSwitch != ParseArgument(argv[i+1])) return false;
			EC_H(AnsiToUnicode(argv[i+1],&pRunOpts->lpszProfile));
			i++;
			break;
		case switchMoreProperties:
			pRunOpts->bRetryStreamProps = true;
			break;
		// Proptag parsing
		case switchDispid:
			if (!bSetMode(&pRunOpts->Mode,cmdmodePropTag)) return false;
			pRunOpts->bDoDispid = true;
			break;
		case switchType:
			if (!bSetMode(&pRunOpts->Mode,cmdmodePropTag)) return false;
			pRunOpts->bDoType = true;
			// If we have a next argument and it's not an option, parse it as a type
			if (i+1 < argc && switchNoSwitch == ParseArgument(argv[i+1]))
			{
				pRunOpts->ulTypeNum = PropTypeNameToPropTypeA(argv[i+1]);
				i++;
			}
			break;
		case switchFlag:
			if (!bSetMode(&pRunOpts->Mode,cmdmodePropTag)) return false;
			pRunOpts->bDoFlag = true;
			// If we have a next argument and it's not an option, parse it as a flag
			if (argc <= i+1 || switchNoSwitch != ParseArgument(argv[i+1])) return false;
			pRunOpts->ulFlagValue = strtoul(argv[i+1],&szEndPtr,16);
			i++;
			break;
		// GUID
		case switchGuid:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeGuid)) return false;
			break;
		// Error parsing
		case switchError:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeErr)) return false;
			break;
		// Smart View parsing
		case switchParser:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeSmartView)) return false;
			if (argc <= i+1 || switchNoSwitch != ParseArgument(argv[i+1])) return false;
			pRunOpts->ulParser = strtoul(argv[i+1],&szEndPtr,10);
			i++;
			break;
		case switchBinary:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeSmartView)) return false;
			pRunOpts->bBinaryFile = true;
			break;
		// ACL Table
		case switchAcl:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeAcls)) return false;
			break;
		// Rules Table
		case switchRule:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeRules)) return false;
			break;
		// Contents tables
		case switchContents:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeContents)) return false;
			pRunOpts->bDoContents = true;
			break;
		case switchAssociatedContents:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeContents)) return false;
			pRunOpts->bDoAssociatedContents = true;
			break;
		case switchSubject:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeContents)) return false;
			if (argc <= i+1 || switchNoSwitch != ParseArgument(argv[i+1])) return false;
			EC_H(AnsiToUnicode(argv[i+1],&pRunOpts->lpszSubject));
			i++;
			break;
		case switchMessageClass:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeContents)) return false;
			if (argc <= i+1 || switchNoSwitch != ParseArgument(argv[i+1])) return false;
			EC_H(AnsiToUnicode(argv[i+1],&pRunOpts->lpszMessageClass));
			i++;
			break;
		case switchMSG:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeContents)) return false;
			pRunOpts->bMSG = true;
			break;
		case switchList:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeContents)) return false;
			pRunOpts->bList = true;
			break;
		// XML
		case switchXML:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeXML)) return false;
			break;
		// FID / MID
		case switchFid:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeFidMid)) return false;
			if (i+1 < argc  && switchNoSwitch == ParseArgument(argv[i+1]))
			{
				EC_H(AnsiToUnicode(argv[i+1],&pRunOpts->lpszFid));
				i++;
			}
			break;
		case switchMid:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeFidMid)) return false;
			if (i+1 < argc  && switchNoSwitch == ParseArgument(argv[i+1]))
			{
				EC_H(AnsiToUnicode(argv[i+1],&pRunOpts->lpszMid));
				i++;
			}
			else
			{
				// We use the blank string to remember the -mid parameter was passed and save having an extra flag
				EC_H(AnsiToUnicode("",&pRunOpts->lpszMid)); // STRING_OK
			}
			break;
		// Child Folders
		case switchChildFolders:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeChildFolders)) return false;
			break;
		// MAPIMIME
		case switchMAPI:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeMAPIMIME)) return false;
			pRunOpts->ulMAPIMIMEFlags |= MAPIMIME_TOMAPI;
			break;
		case switchMIME:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeMAPIMIME)) return false;
			pRunOpts->ulMAPIMIMEFlags |= MAPIMIME_TOMIME;
			break;
		case switchCCSFFlags:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeMAPIMIME)) return false;
			if (argc <= i+1 || switchNoSwitch != ParseArgument(argv[i+1])) return false;
			pRunOpts->ulConvertFlags = strtoul(argv[i+1],&szEndPtr,10);
			i++;
			break;
		case switchRFC822:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeMAPIMIME)) return false;
			pRunOpts->ulMAPIMIMEFlags |= MAPIMIME_RFC822;
			break;
		case switchWrap:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeMAPIMIME)) return false;
			if (argc <= i+1 || switchNoSwitch != ParseArgument(argv[i+1])) return false;
			pRunOpts->ulWrapLines = strtoul(argv[i+1], &szEndPtr, 10);
			pRunOpts->ulMAPIMIMEFlags |= MAPIMIME_WRAP;
			i++;
			break;
		case switchEncoding:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeMAPIMIME)) return false;
			if (argc <= i+1 || switchNoSwitch != ParseArgument(argv[i+1])) return false;
			pRunOpts->ulEncodingType = strtoul(argv[i+1], &szEndPtr, 10);
			pRunOpts->ulMAPIMIMEFlags |= MAPIMIME_ENCODING;
			i++;
			break;
		case switchCharset:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeMAPIMIME)) return false;
			if (argc <= i+3 ||
				switchNoSwitch != ParseArgument(argv[i+1]) ||
				switchNoSwitch != ParseArgument(argv[i+2]) ||
				switchNoSwitch != ParseArgument(argv[i+3]))
				return false;

			pRunOpts->ulCodePage = strtoul(argv[i+1], &szEndPtr, 10);
			pRunOpts->cSetType = (CHARSETTYPE) strtoul(argv[i+2], &szEndPtr, 10);
			if (pRunOpts->cSetType > CHARSET_WEB) return false;
			pRunOpts->cSetApplyType = (CSETAPPLYTYPE) strtoul(argv[i+3], &szEndPtr, 10);
			if (pRunOpts->cSetApplyType > CSET_APPLY_TAG_ALL) return false;
			pRunOpts->ulMAPIMIMEFlags |= MAPIMIME_CHARSET;
			i += 3;
			break;
		case switchAddressBook:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeMAPIMIME)) return false;
			pRunOpts->ulMAPIMIMEFlags |= MAPIMIME_ADDRESSBOOK;
			break;
		case switchUnicode:
			if (!bSetMode(&pRunOpts->Mode,cmdmodeMAPIMIME)) return false;
			pRunOpts->ulMAPIMIMEFlags |= MAPIMIME_UNICODE;
			break;

		case switchNoSwitch:
			// naked option without a flag - we only allow one of these
			if (pRunOpts->lpszUnswitchedOption) return false; // He's already got one, you see.
			EC_H(AnsiToUnicode(argv[i],&pRunOpts->lpszUnswitchedOption));
			break;
		case switchUnknown:
		default:
			// display help
			return false;
			break;
		}
	}

	// If we didn't get a mode set, assume we're in prop tag mode
	if (cmdmodeUnknown == pRunOpts->Mode) pRunOpts->Mode = cmdmodePropTag;

	// If we weren't passed an output file/directory, remember the current directory
	if (!pRunOpts->lpszOutput && pRunOpts->Mode != cmdmodeSmartView)
	{
		CHAR strPath[_MAX_PATH];
		GetCurrentDirectoryA(_MAX_PATH,strPath);

		EC_H(AnsiToUnicode(strPath,&pRunOpts->lpszOutput));
	}

	// Validate that we have bare minimum to run
	switch (pRunOpts->Mode)
	{
	case cmdmodePropTag:
		if (pRunOpts->bDoType && !pRunOpts->bDoPartialSearch && (pRunOpts->lpszUnswitchedOption != NULL)) return false;
		if (!pRunOpts->bDoType && !pRunOpts->bDoPartialSearch && (pRunOpts->lpszUnswitchedOption == NULL)) return false;
		if (pRunOpts->bDoPartialSearch && pRunOpts->bDoType && ulNoMatch == pRunOpts->ulTypeNum) return false;
		if (pRunOpts->bDoFlag && (pRunOpts->bDoPartialSearch || pRunOpts->bDoType)) return false;
		break;
	case cmdmodeGuid:
		// Nothing to check
		break;
	case cmdmodeSmartView:
		if (!pRunOpts->ulParser || !pRunOpts->lpszInput) return false;
		break;
	case cmdmodeAcls:
		// Nothing to check - both lpszProfile and ulFolder are optional
		break;
	case cmdmodeRules:
		// Nothing to check - both lpszProfile and ulFolder are optional
		break;
	case cmdmodeContents:
		if (!pRunOpts->bDoContents && !pRunOpts->bDoAssociatedContents) return false;
		break;
	case cmdmodeXML:
		if (!pRunOpts->lpszInput) return false;
		break;
	case cmdmodeMAPIMIME:
#define CHECKFLAG(__flag) ((pRunOpts->ulMAPIMIMEFlags & (__flag)) == (__flag))
		// Can't convert both ways at once
		if (CHECKFLAG(MAPIMIME_TOMAPI) && CHECKFLAG(MAPIMIME_TOMIME)) return false;

		// Should have given us source and target files
		if (pRunOpts->lpszInput == NULL || 
			pRunOpts->lpszOutput == NULL)
			return false;

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
	case cmdmodeErr:
		// Nothing to check
		break;
	default:
		break;
	}

	// Didn't fail - return true
	return true;
} // ParseArgs

void main(_In_ int argc, _In_count_(argc) char * argv[])
{
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

	if (ProgOpts.bVerbose)
	{
		InitMFC();
		RegKeys[regkeyDEBUG_TAG].ulCurDWORD = 0xFFFFFFFF;
	}

	if (!ProgOpts.bNoAddins)
	{
		RegKeys[regkeyLOADADDINS].ulCurDWORD = true;
		LoadAddIns();
	}

	if (ProgOpts.bHelp)
	{
		DisplayUsage(true);
	}
	else if (!bGoodCommandLine)
	{
		DisplayUsage(false);
	}
	else
	{
		if (ProgOpts.bOnline)
		{
			RegKeys[regKeyMAPI_NO_CACHE].ulCurDWORD = true;
			RegKeys[regkeyMDB_ONLINE].ulCurDWORD = true;
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
		case cmdmodeMAPIMIME:
			DoMAPIMIME(ProgOpts);
			break;
		case cmdmodeChildFolders:
			DoChildFolders(ProgOpts);
			break;
		}
	}

	delete[] ProgOpts.lpszProfile;
	delete[] ProgOpts.lpszUnswitchedOption;
	delete[] ProgOpts.lpszInput;
	delete[] ProgOpts.lpszOutput;
	delete[] ProgOpts.lpszSubject;
	delete[] ProgOpts.lpszMessageClass;
	delete[] ProgOpts.lpszFolderPath;
	delete[] ProgOpts.lpszFid;
	delete[] ProgOpts.lpszMid;
} // main