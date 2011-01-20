#include "stdafx.h"

#include "MrMAPI.h"
#include "MAPIFunctions.h"
#include "InterpretProp2.h"
#include "SmartView.h"
#include "MMAcls.h"
#include "MMErr.h"
#include "MMPropTag.h"
#include "MMRules.h"
#include "MMSmartView.h"

// Initialize MFC for LoadString support later on
void InitMFC()
{
#pragma warning(push)
#pragma warning(disable:6309)
#pragma warning(disable:6387)
	if (!AfxWinInit(::GetModuleHandle(NULL), NULL, ::GetCommandLine(), 0)) return;
#pragma warning(pop)
} // InitMFC

void DisplayUsage()
{
	InitMFC();
	printf("MAPI data collection and parsing tool. Supports property tag lookup, error translation,\n");
	printf("   smart view processing, rule tables and ACL tables.\n");
	printf("MrMAPI currently knows:\n");
	printf("%6d property tags\n",ulPropTagArray);
	printf("%6d dispids\n",ulNameIDArray);
	printf("%6d types\n",ulPropTypeArray);
	printf("%6d guids\n",ulPropGuidArray);
	printf("%6d errors\n",g_ulErrorArray);
	printf("%6d smart view parsers\n",g_cuidParsingTypes-1);
	printf("\n");
	printf("Usage:\n");
	printf("   MrMAPI [-S] [-D] [-N] [-T <type>] <number>|<name>\n");
	printf("   MrMAPI -G\n");
	printf("   MrMAPI -E <error>\n");
	printf("   MrMAPI -P <type> -I <input file> [-B] [-O <output file>]\n");
	printf("   MrMAPI -R <profile> [-F <folder>]\n");
	printf("   MrMAPI -A <profile> [-F <folder>]\n");
	printf("\n");
	printf("   Property Tag Lookup:\n");
	printf("   -S   Perform substring search.\n");
	printf("           With no parameters prints all known properties.\n");
	printf("   -D   Search dispids.\n");
	printf("   -N   Number is in decimal. Ignored for non-numbers.\n");
	printf("   -T   Print information on specified type.\n");
	printf("           With no parameters prints list of known types.\n");
	printf("           When combined with -S, restrict output to given type.\n");
	printf("   -G   Display list of known guids.\n");
	printf("\n");
	printf("   Error Parsing:\n");
	printf("   -E   Map an error code to its name and vice versa (may be combined with -s and -n switches).\n");
	printf("\n");
	printf("   Smart View Parsing:\n");
	printf("   -P   Parser type (number). See list below for supported parsers.\n");
	printf("   -I   Input file.\n");
	printf("   -B   Input file is binary. Default is hex encoded text.\n");
	printf("   -O   Output file (optional).\n");
	printf("\n");
	printf("   Rules Table:\n");
	printf("   -R   Output rules table. Profile optional\n");
	printf("   -F   Folder to output. Default is Inbox. See list below for supported folders.\n");
	printf("\n");
	printf("   ACL Table:\n");
	printf("   -A   Output ACL table. Profile optional\n");
	printf("   -F   Folder to output. Default is Inbox. See list below for supported folders.\n");
	printf("\n");
	printf("Smart View Parsers:\n");
	// Print smart view options
	ULONG i = 1;
	for (i = 1 ; i < g_cuidParsingTypes ; i++)
	{
		HRESULT hRes = S_OK;
		CString szStruct;
		EC_B(szStruct.LoadString(g_uidParsingTypes[i]));
		_tprintf(_T("   %2d %s\n"),i,(LPCTSTR) szStruct);
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
} // DisplayUsage

BOOL bSetMode(_In_ CmdMode* pMode, _In_ CmdMode TargetMode)
{
	if (pMode && ((cmdmodeUnknown == *pMode) || (TargetMode == *pMode)))
	{
		*pMode = TargetMode;
		return true;
	}
	return false;
} // bSetMode

// Checks if szArg is an option, and if it is, returns first character
char GetOption(_In_z_ LPCSTR szArg)
{
	if (!szArg) return NULL;
	switch (szArg[0])
	{
	case '-':
	case '/':
	case '\\':
		if (szArg[1] != 0) return (char) tolower(szArg[1]);
		break;
	default:
		break;
	}
	return NULL;
} // GetOption

// Parses command line arguments and fills out MYOPTIONS
BOOL ParseArgs(_In_ int argc, _In_count_(argc) char * argv[], _Out_ MYOPTIONS * pRunOpts)
{
	HRESULT hRes = S_OK;
	// Clear our options list
	ZeroMemory(pRunOpts,sizeof(MYOPTIONS));

	pRunOpts->ulTypeNum = ulNoMatch;

	if (!pRunOpts) return false;
	if (1 == argc) return false;

	for (int i = 1; i < argc; i++)
	{
		switch (GetOption(argv[i]))
		{
			// Global flags
		case 's':
			pRunOpts->bDoPartialSearch = true;
			break;
		case 'n':
			pRunOpts->bDoDecimal = true;
			break;
		// Proptag parsing
		case 'd':
			if (!bSetMode(&pRunOpts->Mode,cmdmodePropTag)) return false;
			pRunOpts->bDoDispid = true;
			break;
		case 't':
			if (!bSetMode(&pRunOpts->Mode,cmdmodePropTag)) return false;
			pRunOpts->bDoType = true;
			// If we have a next argument and it's not an option, parse it as a type
			if (i+1 < argc && !GetOption(argv[i+1]))
			{
				pRunOpts->ulTypeNum = PropTypeNameToPropTypeA(argv[i+1]);
				i++;
			}
			break;
			// GUID
		case 'g':
			if (!bSetMode(&pRunOpts->Mode,cmdmodeGuid)) return false;
			break;
		// Error parsing
		case 'e':
			if (!bSetMode(&pRunOpts->Mode,cmdmodeErr)) return false;
			break;
		// Smart View parsing
		case 'p':
			if (!bSetMode(&pRunOpts->Mode,cmdmodeSmartView)) return false;
			if (i+1 < argc)
			{
				LPSTR szEndPtr = NULL;
				pRunOpts->ulParser = strtoul(argv[i+1],&szEndPtr,10);
				i++;
			}
			break;
		case 'i':
			if (!bSetMode(&pRunOpts->Mode,cmdmodeSmartView)) return false;
			if (i+1 < argc)
			{
				EC_H(AnsiToUnicode(argv[i+1],&pRunOpts->lpszInput));
				i++;
			}
			break;
		case 'b':
			if (!bSetMode(&pRunOpts->Mode,cmdmodeSmartView)) return false;
			pRunOpts->bBinaryFile = true;
			break;
		case 'o':
			if (!bSetMode(&pRunOpts->Mode,cmdmodeSmartView)) return false;
			if (i+1 < argc)
			{
				EC_H(AnsiToUnicode(argv[i+1],&pRunOpts->lpszOutput));
				i++;
			}
			break;
		case 'a':
			if (!bSetMode(&pRunOpts->Mode,cmdmodeAcls)) return false;
			if (i+1 < argc && !GetOption(argv[i+1]))
			{
				EC_H(AnsiToUnicode(argv[i+1],&pRunOpts->lpszProfile));
				i++;
			}
			break;
		case 'r':
			if (!bSetMode(&pRunOpts->Mode,cmdmodeRules)) return false;
			if (i+1 < argc && !GetOption(argv[i+1]))
			{
				EC_H(AnsiToUnicode(argv[i+1],&pRunOpts->lpszProfile));
				i++;
			}
			break;
		case 'f':
			if (i+1 < argc && !GetOption(argv[i+1]))
			{
				LPSTR szEndPtr = NULL;
				pRunOpts->ulFolder = strtoul(argv[i+1],&szEndPtr,10);
				i++;
			}
			break;
		case NULL:
			// naked option without a flag - we only allow one of these
			if (pRunOpts->lpszUnswitchedOption) return false; // He's already got one, you see.
			EC_H(AnsiToUnicode(argv[i],&pRunOpts->lpszUnswitchedOption));
			break;
		default:
			// display help
			return false;
			break;
		}
	}

	// If we didn't get a mode set, assume we're in prop tag mode
	if (cmdmodeUnknown == pRunOpts->Mode) pRunOpts->Mode = cmdmodePropTag;

	// Validate that we have bare minimum to run
	switch (pRunOpts->Mode)
	{
	case cmdmodePropTag:
		if (pRunOpts->bDoType && !pRunOpts->bDoPartialSearch && (pRunOpts->lpszUnswitchedOption != NULL)) return false;
		break;
	case cmdmodeGuid:
		// Nothing to check
		break;
	case cmdmodeSmartView:
		if (!pRunOpts->ulParser || !pRunOpts->lpszInput) return false;
		break;
	case cmdmodeAcls:
		// Nothing to check - both lpszProfile and ulFolder are optional
		if (DEFAULT_UNSPECIFIED == pRunOpts->ulFolder) pRunOpts->ulFolder = DEFAULT_INBOX;
		break;
	case cmdmodeRules:
		// Nothing to check - both lpszProfile and ulFolder are optional
		if (DEFAULT_UNSPECIFIED == pRunOpts->ulFolder) pRunOpts->ulFolder = DEFAULT_INBOX;
		break;
	case cmdmodeErr:
		// Nothing to check
	default:
		break;
	}

	// Didn't fail - return true
	return true;
} // ParseArgs

void main(_In_ int argc, _In_count_(argc) char * argv[])
{
	LoadAddIns();

	MYOPTIONS ProgOpts;

	if (!ParseArgs(argc, argv, &ProgOpts))
	{
		DisplayUsage();
		delete[] ProgOpts.lpszUnswitchedOption;
		delete[] ProgOpts.lpszInput;
		delete[] ProgOpts.lpszOutput;
		delete[] ProgOpts.lpszProfile;
		return;
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
	}

	delete[] ProgOpts.lpszUnswitchedOption;
	delete[] ProgOpts.lpszInput;
	delete[] ProgOpts.lpszOutput;
	delete[] ProgOpts.lpszProfile;
} // main