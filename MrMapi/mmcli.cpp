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
	// For some reason, placing this in the lambda causes a compiler error. So we'll make it an inline function
	inline MYOPTIONS* GetMyOptions(OPTIONS* _options) { return dynamic_cast<MYOPTIONS*>(_options); }

	OptParser switchSearchParser{L"Search", cmdmodeUnknown, 0, 0, OPT_DOPARTIALSEARCH};
	OptParser switchDecimalParser{L"Number", cmdmodeUnknown, 0, 0, OPT_DODECIMAL};
	OptParser switchFolderParser{L"Folder",
								 cmdmodeUnknown,
								 1,
								 1,
								 OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_NEEDFOLDER | OPT_INITMFC,
								 [](auto _options) {
									 auto options = GetMyOptions(_options);
									 options->ulFolder = strings::wstringToUlong(switchFolderParser.args.front(), 10);
									 if (options->ulFolder)
									 {
										 options->lpszFolderPath = switchFolderParser.args.front();
										 options->ulFolder = mapi::DEFAULT_INBOX;
									 }

									 return true;
								 }};
	OptParser switchOutputParser{L"Output", cmdmodeUnknown, 1, 1, OPT_NOOPT, [](auto _options) {
									 auto options = GetMyOptions(_options);
									 options->lpszOutput = switchOutputParser.args.front();
									 return true;
								 }};
	OptParser switchDispidParser{L"Dispids", cmdmodePropTag, 0, 0, OPT_DODISPID};
	OptParser switchTypeParser{L"Type", cmdmodePropTag, 0, 1, OPT_DOTYPE, [](auto _options) {
								   if (!switchTypeParser.args.empty())
								   {
									   auto options = GetMyOptions(_options);
									   options->ulTypeNum =
										   proptype::PropTypeNameToPropType(switchTypeParser.args.front());
								   }

								   return true;
							   }};
	OptParser switchGuidParser{L"Guids", cmdmodeGuid, 0, 0, OPT_NOOPT};
	OptParser switchErrorParser{L"Error", cmdmodeErr, 0, 0, OPT_NOOPT};
	OptParser switchParserParser{L"ParserType",
								 cmdmodeSmartView,
								 1,
								 1,
								 OPT_INITMFC | OPT_NEEDINPUTFILE,
								 [](auto _options) {
									 auto options = GetMyOptions(_options);
									 options->ulSVParser = strings::wstringToUlong(switchParserParser.args.front(), 10);
									 return true;
								 }};
	OptParser switchInputParser{L"Input", cmdmodeUnknown, 1, 1, OPT_NOOPT, [](auto _options) {
									auto options = GetMyOptions(_options);
									options->lpszInput = switchInputParser.args.front();
									return true;
								}};
	OptParser switchBinaryParser{L"Binary", cmdmodeSmartView, 0, 0, OPT_BINARYFILE};
	OptParser switchAclParser{L"Acl",
							  cmdmodeAcls,
							  0,
							  0,
							  OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDFOLDER};
	OptParser switchRuleParser{L"Rules",
							   cmdmodeRules,
							   0,
							   0,
							   OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDFOLDER};
	OptParser switchContentsParser{L"Contents",
								   cmdmodeContents,
								   0,
								   0,
								   OPT_DOCONTENTS | OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC};
	OptParser switchAssociatedContentsParser{L"HiddenContents",
											 cmdmodeContents,
											 0,
											 0,
											 OPT_DOASSOCIATEDCONTENTS | OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON |
												 OPT_INITMFC};
	OptParser switchMorePropertiesParser{L"MoreProperties", cmdmodeUnknown, 0, 0, OPT_RETRYSTREAMPROPS};
	OptParser switchNoAddinsParser{L"NoAddins", cmdmodeUnknown, 0, 0, OPT_NOADDINS};
	OptParser switchOnlineParser{L"Online", cmdmodeUnknown, 0, 0, OPT_ONLINE};
	OptParser switchMAPIParser{L"MAPI",
							   cmdmodeMAPIMIME,
							   0,
							   0,
							   OPT_NEEDMAPIINIT | OPT_INITMFC | OPT_NEEDINPUTFILE | OPT_NEEDOUTPUTFILE,
							   [](auto _options) {
								   auto options = GetMyOptions(_options);
								   options->MAPIMIMEFlags |= MAPIMIME_TOMAPI;
								   return true;
							   }};
	OptParser switchMIMEParser{L"MIME",
							   cmdmodeMAPIMIME,
							   0,
							   0,
							   OPT_NEEDMAPIINIT | OPT_INITMFC | OPT_NEEDINPUTFILE | OPT_NEEDOUTPUTFILE,
							   [](auto _options) {
								   auto options = GetMyOptions(_options);
								   options->MAPIMIMEFlags |= MAPIMIME_TOMIME;
								   return true;
							   }};
	OptParser switchCCSFFlagsParser{L"CCSFFlags", cmdmodeMAPIMIME, 1, 1, OPT_NOOPT, [](auto _options) {
										auto options = GetMyOptions(_options);
										options->convertFlags = static_cast<CCSFLAGS>(
											strings::wstringToUlong(switchCCSFFlagsParser.args.front(), 10));
										return true;
									}};
	OptParser switchRFC822Parser{L"RFC822", cmdmodeMAPIMIME, 0, 0, OPT_NOOPT, [](auto _options) {
									 auto options = GetMyOptions(_options);
									 options->MAPIMIMEFlags |= MAPIMIME_RFC822;
									 return true;
								 }};
	OptParser switchWrapParser{L"Wrap", cmdmodeMAPIMIME, 1, 1, OPT_NOOPT, [](auto _options) {
								   auto options = GetMyOptions(_options);
								   options->ulWrapLines = strings::wstringToUlong(switchWrapParser.args.front(), 10);
								   options->MAPIMIMEFlags |= MAPIMIME_WRAP;
								   return true;
							   }};
	OptParser switchEncodingParser{L"Encoding", cmdmodeMAPIMIME, 1, 1, OPT_NOOPT, [](auto _options) {
									   auto options = GetMyOptions(_options);
									   options->ulEncodingType =
										   strings::wstringToUlong(switchEncodingParser.args.front(), 10);
									   options->MAPIMIMEFlags |= MAPIMIME_ENCODING;
									   return true;
								   }};
	OptParser switchCharsetParser{L"Charset", cmdmodeMAPIMIME, 3, 3, OPT_NOOPT, [](auto _options) {
									  auto options = GetMyOptions(_options);
									  options->ulCodePage = strings::wstringToUlong(switchCharsetParser.args[0], 10);
									  options->cSetType = static_cast<CHARSETTYPE>(
										  strings::wstringToUlong(switchCharsetParser.args[1], 10));
									  if (options->cSetType > CHARSET_WEB)
									  {
										  return false;
									  }

									  options->cSetApplyType = static_cast<CSETAPPLYTYPE>(
										  strings::wstringToUlong(switchCharsetParser.args[2], 10));
									  if (options->cSetApplyType > CSET_APPLY_TAG_ALL)
									  {
										  return false;
									  }

									  options->MAPIMIMEFlags |= MAPIMIME_CHARSET;
									  return true;
								  }};
	OptParser switchAddressBookParser{L"AddressBook", cmdmodeMAPIMIME, 0, 0, OPT_NEEDMAPILOGON, [](auto _options) {
										  auto options = GetMyOptions(_options);
										  options->MAPIMIMEFlags |= MAPIMIME_ADDRESSBOOK;
										  return true;
									  }}; // special case which needs a logon
	OptParser switchUnicodeParser{L"Unicode", cmdmodeMAPIMIME, 0, 0, OPT_NOOPT, [](auto _options) {
									  auto options = GetMyOptions(_options);
									  options->MAPIMIMEFlags |= MAPIMIME_UNICODE;
									  return true;
								  }};
	OptParser switchProfileParser{L"Profile", cmdmodeUnknown, 0, 1, OPT_PROFILE, [](auto _options) {
									  auto options = GetMyOptions(_options);
									  if (!switchProfileParser.args.empty())
									  {
										  options->lpszProfile = switchProfileParser.args.front();
									  }

									  return true;
								  }};
	OptParser switchXMLParser{L"XML", cmdmodeXML, 0, 0, OPT_NEEDMAPIINIT | OPT_INITMFC | OPT_NEEDINPUTFILE};
	OptParser switchSubjectParser{L"Subject", cmdmodeContents, 1, 1, OPT_NOOPT, [](auto _options) {
									  auto options = GetMyOptions(_options);
									  options->lpszSubject = switchSubjectParser.args.front();
									  return true;
								  }};
	OptParser switchMessageClassParser{L"MessageClass", cmdmodeContents, 1, 1, OPT_NOOPT, [](auto _options) {
										   auto options = GetMyOptions(_options);
										   options->lpszMessageClass = switchMessageClassParser.args.front();
										   return true;
									   }};
	OptParser switchMSGParser{L"MSG", cmdmodeContents, 0, 0, OPT_MSG};
	OptParser switchListParser{L"List", cmdmodeContents, 0, 0, OPT_LIST};
	OptParser switchChildFoldersParser{L"ChildFolders",
									   cmdmodeChildFolders,
									   0,
									   0,
									   OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDFOLDER};
	OptParser switchFidParser{L"FID",
							  cmdmodeFidMid,
							  0,
							  1,
							  OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDSTORE,
							  [](auto _options) {
								  auto options = GetMyOptions(_options);
								  if (!switchFidParser.args.empty())
								  {
									  options->lpszFid = switchFidParser.args.front();
								  }

								  return true;
							  }};
	OptParser switchMidParser{L"MID",
							  cmdmodeFidMid,
							  0,
							  1,
							  OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_MID,
							  [](auto _options) {
								  auto options = GetMyOptions(_options);
								  if (!switchMidParser.args.empty())
								  {
									  options->lpszMid = switchMidParser.args.front();
								  }
								  else
								  {
									  // We use the blank string to remember the -mid parameter was passed and save having an extra flag
									  // TODO: Just check switchMidParser for seen instead
									  options->lpszMid = L"";
								  }

								  return true;
							  }};
	OptParser switchFlagParser{L"Flag", cmdmodeUnknown, 1, 1, OPT_NOOPT, [](auto _options) {
								   auto options = GetMyOptions(_options);
								   LPWSTR szEndPtr = nullptr;
								   // We must have a next argument, but it could be a string or a number
								   options->lpszFlagName = switchFlagParser.args.front();
								   options->ulFlagValue = wcstoul(switchFlagParser.args.front().c_str(), &szEndPtr, 16);

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

								   return true;
							   }}; // can't know until we parse the argument
	OptParser switchRecentParser{L"Recent", cmdmodeContents, 1, 1, OPT_NOOPT, [](auto _options) {
									 auto options = GetMyOptions(_options);
									 options->ulCount = strings::wstringToUlong(switchRecentParser.args.front(), 10);
									 return true;
								 }};
	OptParser switchStoreParser{L"Store",
								cmdmodeStoreProperties,
								0,
								1,
								OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC,
								[](auto _options) {
									auto options = GetMyOptions(_options);
									if (!switchStoreParser.args.empty())
									{
										LPWSTR szEndPtr = nullptr;
										options->ulStore =
											wcstoul(switchStoreParser.args.front().c_str(), &szEndPtr, 10);

										// If we parsed completely, this was a store number
										if (NULL == szEndPtr[0])
										{
											// Increment ulStore so we can use to distinguish an unset value
											options->ulStore++;
										}
										// Else it was a naked option - leave it on the stack
									}

									return true;
								}};
	OptParser switchVersionParser{L"Version", cmdmodeUnknown, 1, 1, OPT_NOOPT, [](auto _options) {
									  auto options = GetMyOptions(_options);
									  options->lpszVersion = switchVersionParser.args.front();
									  return true;
								  }};
	OptParser switchSizeParser{L"Size",
							   cmdmodeFolderSize,
							   0,
							   0,
							   OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDFOLDER};
	OptParser switchPSTParser{L"PST", cmdmodePST, 0, 0, OPT_NEEDINPUTFILE};
	OptParser switchProfileSectionParser{L"ProfileSection",
										 cmdmodeProfile,
										 1,
										 1,
										 OPT_PROFILE | OPT_NEEDMAPIINIT | OPT_INITMFC,
										 [](auto _options) {
											 auto options = GetMyOptions(_options);
											 options->lpszProfileSection = switchProfileSectionParser.args.front();
											 return true;
										 }};
	OptParser switchByteSwappedParser{L"ByteSwapped",
									  cmdmodeProfile,
									  0,
									  0,
									  OPT_PROFILE | OPT_NEEDMAPIINIT | OPT_INITMFC,
									  [](auto _options) {
										  auto options = GetMyOptions(_options);
										  options->bByteSwapped = true;
										  return true;
									  }};
	OptParser switchReceiveFolderParser{L"ReceiveFolder",
										cmdmodeReceiveFolder,
										0,
										1,
										OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_NEEDSTORE | OPT_INITMFC,
										[](auto _options) {
											auto options = GetMyOptions(_options);
											if (!switchReceiveFolderParser.args.empty())
											{
												LPWSTR szEndPtr = nullptr;
												options->ulStore = wcstoul(
													switchReceiveFolderParser.args.front().c_str(), &szEndPtr, 10);

												// If we parsed completely, this was a store number
												if (NULL == szEndPtr[0])
												{
													// Increment ulStore so we can use to distinguish an unset value
													options->ulStore++;
												}
												// Else it was a naked option - leave it on the stack
											}

											return true;
										}};
	OptParser switchSkipParser{L"Skip", cmdmodeUnknown, 0, 0, OPT_SKIPATTACHMENTS};
	OptParser switchSearchStateParser{L"SearchState",
									  cmdmodeSearchState,
									  0,
									  1,
									  OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDFOLDER};

	// If we want to add aliases for any switches, add them here
	OptParser switchHelpParserAlias{L"Help", cmdmodeHelpFull, 0, 0, OPT_INITMFC};

	std::vector<OptParser*> g_Parsers = {
		&switchHelpParser,
		&switchVerboseParser,
		&switchSearchParser,
		&switchDecimalParser,
		&switchFolderParser,
		&switchOutputParser,
		&switchDispidParser,
		&switchTypeParser,
		&switchGuidParser,
		&switchErrorParser,
		&switchParserParser,
		&switchInputParser,
		&switchBinaryParser,
		&switchAclParser,
		&switchRuleParser,
		&switchContentsParser,
		&switchAssociatedContentsParser,
		&switchMorePropertiesParser,
		&switchNoAddinsParser,
		&switchOnlineParser,
		&switchMAPIParser,
		&switchMIMEParser,
		&switchCCSFFlagsParser,
		&switchRFC822Parser,
		&switchWrapParser,
		&switchEncodingParser,
		&switchCharsetParser,
		&switchAddressBookParser,
		&switchUnicodeParser,
		&switchProfileParser,
		&switchXMLParser,
		&switchSubjectParser,
		&switchMessageClassParser,
		&switchMSGParser,
		&switchListParser,
		&switchChildFoldersParser,
		&switchFidParser,
		&switchMidParser,
		&switchFlagParser,
		&switchRecentParser,
		&switchStoreParser,
		&switchVersionParser,
		&switchSizeParser,
		&switchPSTParser,
		&switchProfileSectionParser,
		&switchByteSwappedParser,
		&switchReceiveFolderParser,
		&switchSkipParser,
		&switchSearchStateParser,
		// If we want to add aliases for any switches, add them here
		&switchHelpParserAlias,
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
		printf("   MrMAPI -%ws\n", switchHelpParser.szSwitch);
		printf(
			"   MrMAPI [-%ws] [-%ws] [-%ws] [-%ws <type>] <property number>|<property name>\n",
			switchSearchParser.szSwitch,
			switchDispidParser.szSwitch,
			switchDecimalParser.szSwitch,
			switchTypeParser.szSwitch);
		printf("   MrMAPI -%ws\n", switchGuidParser.szSwitch);
		printf("   MrMAPI -%ws <error>\n", switchErrorParser.szSwitch);
		printf(
			"   MrMAPI -%ws <type> -%ws <input file> [-%ws] [-%ws <output file>]\n",
			switchParserParser.szSwitch,
			switchInputParser.szSwitch,
			switchBinaryParser.szSwitch,
			switchOutputParser.szSwitch);
		printf(
			"   MrMAPI -%ws <flag value> [-%ws] [-%ws] <property number>|<property name>\n",
			switchFlagParser.szSwitch,
			switchDispidParser.szSwitch,
			switchDecimalParser.szSwitch);
		printf("   MrMAPI -%ws <flag name>\n", switchFlagParser.szSwitch);
		printf(
			"   MrMAPI -%ws [-%ws <profile>] [-%ws <folder>]\n",
			switchRuleParser.szSwitch,
			switchProfileParser.szSwitch,
			switchFolderParser.szSwitch);
		printf(
			"   MrMAPI -%ws [-%ws <profile>] [-%ws <folder>]\n",
			switchAclParser.szSwitch,
			switchProfileParser.szSwitch,
			switchFolderParser.szSwitch);
		printf(
			"   MrMAPI -%ws | -%ws [-%ws <profile>] [-%ws <folder>] [-%ws <output directory>]\n",
			switchContentsParser.szSwitch,
			switchAssociatedContentsParser.szSwitch,
			switchProfileParser.szSwitch,
			switchFolderParser.szSwitch,
			switchOutputParser.szSwitch);
		printf(
			"          [-%ws <subject>] [-%ws <message class>] [-%ws] [-%ws] [-%ws <count>] [-%ws]\n",
			switchSubjectParser.szSwitch,
			switchMessageClassParser.szSwitch,
			switchMSGParser.szSwitch,
			switchListParser.szSwitch,
			switchRecentParser.szSwitch,
			switchSkipParser.szSwitch);
		printf(
			"   MrMAPI -%ws [-%ws <profile>] [-%ws <folder>]\n",
			switchChildFoldersParser.szSwitch,
			switchProfileParser.szSwitch,
			switchFolderParser.szSwitch);
		printf(
			"   MrMAPI -%ws -%ws <path to input file> -%ws <path to output file> [-%ws]\n",
			switchXMLParser.szSwitch,
			switchInputParser.szSwitch,
			switchOutputParser.szSwitch,
			switchSkipParser.szSwitch);
		printf(
			"   MrMAPI -%ws [fid] [-%ws [mid]] [-%ws <profile>]\n",
			switchFidParser.szSwitch,
			switchMidParser.szSwitch,
			switchProfileParser.szSwitch);
		printf(
			"   MrMAPI [<property number>|<property name>] -%ws [<store num>] [-%ws <profile>]\n",
			switchStoreParser.szSwitch,
			switchProfileParser.szSwitch);
		printf(
			"   MrMAPI [<property number>|<property name>] -%ws <folder> [-%ws <profile>]\n",
			switchFolderParser.szSwitch,
			switchProfileParser.szSwitch);
		printf(
			"   MrMAPI -%ws -%ws <folder> [-%ws <profile>]\n",
			switchSizeParser.szSwitch,
			switchFolderParser.szSwitch,
			switchProfileParser.szSwitch);
		printf(
			"   MrMAPI -%ws | -%ws -%ws <path to input file> -%ws <path to output file> [-%ws <conversion flags>]\n",
			switchMAPIParser.szSwitch,
			switchMIMEParser.szSwitch,
			switchInputParser.szSwitch,
			switchOutputParser.szSwitch,
			switchCCSFFlagsParser.szSwitch);
		printf(
			"          [-%ws] [-%ws <Decimal number of characters>] [-%ws <Decimal number indicating encoding>]\n",
			switchRFC822Parser.szSwitch,
			switchWrapParser.szSwitch,
			switchEncodingParser.szSwitch);
		printf(
			"          [-%ws] [-%ws] [-%ws CodePage CharSetType CharSetApplyType]\n",
			switchAddressBookParser.szSwitch,
			switchUnicodeParser.szSwitch,
			switchCharsetParser.szSwitch);
		printf("   MrMAPI -%ws -%ws <path to input file>\n", switchPSTParser.szSwitch, switchInputParser.szSwitch);
		printf(
			"   MrMAPI -%ws [<profile> [-%ws <profilesection> [-%ws]] -%ws <output file>]\n",
			switchProfileParser.szSwitch,
			switchProfileSectionParser.szSwitch,
			switchByteSwappedParser.szSwitch,
			switchOutputParser.szSwitch);
		printf(
			"   MrMAPI -%ws [<store num>] [-%ws <profile>]\n",
			switchReceiveFolderParser.szSwitch,
			switchProfileParser.szSwitch);
		printf(
			"   MrMAPI -%ws -%ws <folder> [-%ws <profile>]\n",
			switchSearchStateParser.szSwitch,
			switchFolderParser.szSwitch,
			switchProfileParser.szSwitch);

		if (bFull)
		{
			printf("\n");
			printf("All switches may be shortened if the intended switch is unambiguous.\n");
			printf("For example, -T may be used instead of -%ws.\n", switchTypeParser.szSwitch);
		}
		printf("\n");
		printf("   Help:\n");
		printf("   -%ws   Display expanded help.\n", switchHelpParser.szSwitch);
		if (bFull)
		{
			printf("\n");
			printf("   Property Tag Lookup:\n");
			printf("   -S   (or -%ws) Perform substring search.\n", switchSearchParser.szSwitch);
			printf("           With no parameters prints all known properties.\n");
			printf("   -D   (or -%ws) Search dispids.\n", switchDispidParser.szSwitch);
			printf("   -N   (or -%ws) Number is in decimal. Ignored for non-numbers.\n", switchDecimalParser.szSwitch);
			printf("   -T   (or -%ws) Print information on specified type.\n", switchTypeParser.szSwitch);
			printf("           With no parameters prints list of known types.\n");
			printf("           When combined with -S, restrict output to given type.\n");
			printf("   -G   (or -%ws) Display list of known guids.\n", switchGuidParser.szSwitch);
			printf("\n");
			printf("   Flag Lookup:\n");
			printf("   -Fl  (or -%ws) Look up flags for specified property.\n", switchFlagParser.szSwitch);
			printf("           May be combined with -D and -N switches, but all flag values must be in hex.\n");
			printf("   -Fl  (or -%ws) Look up flag name and output its value.\n", switchFlagParser.szSwitch);
			printf("\n");
			printf("   Error Parsing:\n");
			printf("   -E   (or -%ws) Map an error code to its name and vice versa.\n", switchErrorParser.szSwitch);
			printf("           May be combined with -S and -N switches.\n");
			printf("\n");
			printf("   Smart View Parsing:\n");
			printf(
				"   -P   (or -%ws) Parser type (number). See list below for supported parsers.\n",
				switchParserParser.szSwitch);
			printf(
				"   -B   (or -%ws) Input file is binary. Default is hex encoded text.\n", switchBinaryParser.szSwitch);
			printf("\n");
			printf("   Rules Table:\n");
			printf("   -R   (or -%ws) Output rules table. Profile optional.\n", switchRuleParser.szSwitch);
			printf("\n");
			printf("   ACL Table:\n");
			printf("   -A   (or -%ws) Output ACL table. Profile optional.\n", switchAclParser.szSwitch);
			printf("\n");
			printf("   Contents Table:\n");
			printf(
				"   -C   (or -%ws) Output contents table. May be combined with -H. Profile optional.\n",
				switchContentsParser.szSwitch);
			printf(
				"   -H   (or -%ws) Output associated contents table. May be combined with -C. Profile optional\n",
				switchAssociatedContentsParser.szSwitch);
			printf("   -Su  (or -%ws) Subject of messages to output.\n", switchSubjectParser.szSwitch);
			printf("   -Me  (or -%ws) Message class of messages to output.\n", switchMessageClassParser.szSwitch);
			printf("   -Ms  (or -%ws) Output as .MSG instead of XML.\n", switchMSGParser.szSwitch);
			printf("   -L   (or -%ws) List details to screen and do not output files.\n", switchListParser.szSwitch);
			printf(
				"   -Re  (or -%ws) Restrict output to the 'count' most recent messages.\n",
				switchRecentParser.szSwitch);
			printf("\n");
			printf("   Child Folders:\n");
			printf("   -Chi (or -%ws) Display child folders of selected folder.\n", switchChildFoldersParser.szSwitch);
			printf("\n");
			printf("   MSG File Properties\n");
			printf("   -X   (or -%ws) Output properties of an MSG file as XML.\n", switchXMLParser.szSwitch);
			printf("\n");
			printf("   MID/FID Lookup\n");
			printf("   -Fi  (or -%ws) Folder ID (FID) to search for.\n", switchFidParser.szSwitch);
			printf(
				"           If -%ws is specified without a FID, search/display all folders\n",
				switchFidParser.szSwitch);
			printf("   -Mid (or -%ws) Message ID (MID) to search for.\n", switchMidParser.szSwitch);
			printf(
				"           If -%ws is specified without a MID, display all messages in folders specified by the FID "
				"parameter.\n",
				switchMidParser.szSwitch);
			printf("\n");
			printf("   Store Properties\n");
			printf("   -St  (or -%ws) Output properties of stores as XML.\n", switchStoreParser.szSwitch);
			printf("           If store number is specified, outputs properties of a single store.\n");
			printf("           If a property is specified, outputs only that property.\n");
			printf("\n");
			printf("   Folder Properties\n");
			printf("   -F   (or -%ws) Output properties of a folder as XML.\n", switchFolderParser.szSwitch);
			printf("           If a property is specified, outputs only that property.\n");
			printf("   -Size         Output size of a folder and all subfolders.\n");
			printf("           Use -%ws to specify which folder to scan.\n", switchFolderParser.szSwitch);
			printf("   -SearchState  Output search folder state.\n");
			printf("           Use -%ws to specify which folder to scan.\n", switchFolderParser.szSwitch);
			printf("\n");
			printf("   MAPI <-> MIME Conversion:\n");
			printf("   -Ma  (or -%ws) Convert an EML file to MAPI format (MSG file).\n", switchMAPIParser.szSwitch);
			printf("   -Mi  (or -%ws) Convert an MSG file to MIME format (EML file).\n", switchMIMEParser.szSwitch);
			printf(
				"   -I   (or -%ws) Indicates the input file for conversion, either a MIME-formatted EML file or an MSG "
				"file.\n",
				switchInputParser.szSwitch);
			printf("   -O   (or -%ws) Indicates the output file for the conversion.\n", switchOutputParser.szSwitch);
			printf(
				"   -Cc  (or -%ws) Indicates specific flags to pass to the converter.\n",
				switchCCSFFlagsParser.szSwitch);
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
				switchRFC822Parser.szSwitch);
			printf("           If not present, RFC1521 is used instead.\n");
			printf(
				"   -W   (or -%ws) (MAPI->MIME only) Indicates the maximum number of characters in each line in the\n",
				switchWrapParser.szSwitch);
			printf("           generated EML. Default value is 74. A value of 0 indicates no wrapping.\n");
			printf(
				"   -En  (or -%ws) (MAPI->MIME only) Indicates the encoding type to use. Supported values are:\n",
				switchEncodingParser.szSwitch);
			printf("              1 - Base64\n");
			printf("              2 - UUENCODE\n");
			printf("              3 - Quoted-Printable\n");
			printf("              4 - 7bit (DEFAULT)\n");
			printf("              5 - 8bit\n");
			printf(
				"   -Ad  (or -%ws) Pass MAPI Address Book into converter. Profile optional.\n",
				switchAddressBookParser.szSwitch);
			printf(
				"   -U   (or -%ws) (MIME->MAPI only) The resulting MSG file should be unicode.\n",
				switchUnicodeParser.szSwitch);
			printf(
				"   -Ch  (or -%ws) (MIME->MAPI only) Character set - three required parameters:\n",
				switchCharsetParser.szSwitch);
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
			printf("   -I   (or -%ws) PST file to be analyzed.\n", switchInputParser.szSwitch);
			printf("\n");
			printf("   Profiles\n");
			printf("   -Pr  (or -%ws) Output list of profiles\n", switchProfileParser.szSwitch);
			printf("           If a profile is specified, exports that profile.\n");
			printf("   -ProfileSection If specified, output specific profile section.\n");
			printf(
				"   -B   (or -%ws) If specified, profile section guid is byte swapped.\n",
				switchByteSwappedParser.szSwitch);
			printf("   -O   (or -%ws) Indicates the output file for profile export.\n", switchOutputParser.szSwitch);
			printf("           Required if a profile is specified.\n");
			printf("\n");
			printf("   Receive Folder Table\n");
			printf(
				"   -%ws Displays Receive Folder Table for the specified store\n", switchReceiveFolderParser.szSwitch);
			printf("\n");
			printf("   Universal Options:\n");
			printf("   -I   (or -%ws) Input file.\n", switchInputParser.szSwitch);
			printf("   -O   (or -%ws) Output file or directory.\n", switchOutputParser.szSwitch);
			printf(
				"   -F   (or -%ws) Folder to scan. Default is Inbox. See list below for supported folders.\n",
				switchFolderParser.szSwitch);
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
			printf("   -Pr  (or -%ws) Profile for MAPILogonEx.\n", switchProfileParser.szSwitch);
			printf(
				"   -M   (or -%ws) More properties. Tries harder to get stream properties. May take longer.\n",
				switchMorePropertiesParser.szSwitch);
			printf("   -Sk  (or -%ws) Skip embedded message attachments on export.\n", switchSkipParser.szSwitch);
			printf("   -No  (or -%ws) No Addins. Don't load any add-ins.\n", switchNoAddinsParser.szSwitch);
			printf("   -On  (or -%ws) Online mode. Bypass cached mode.\n", switchOnlineParser.szSwitch);
			printf("   -V   (or -%ws) Verbose. Turn on all debug output.\n", switchVerboseParser.szSwitch);
			printf("\n");
			printf("   MAPI Implementation Options:\n");
			printf("   -%ws MAPI Version to load - supported values\n", switchVersionParser.szSwitch);
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

	void PostParseCheck(OPTIONS* _options)
	{
		auto options = dynamic_cast<MYOPTIONS*>(_options);
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
			WCHAR strPath[_MAX_PATH];
			GetCurrentDirectoryW(_MAX_PATH, strPath);

			options->lpszOutput = strPath;
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