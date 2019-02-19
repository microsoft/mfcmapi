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

	OptParser switchSearch{L"Search", cmdmodeUnknown, 0, 0, OPT_DOPARTIALSEARCH};
	OptParser switchDecimal{L"Number", cmdmodeUnknown, 0, 0, OPT_DODECIMAL};
	OptParser switchFolder{L"Folder",
						   cmdmodeUnknown,
						   1,
						   1,
						   OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_NEEDFOLDER | OPT_INITMFC,
						   [](auto _options) {
							   auto options = GetMyOptions(_options);
							   options->ulFolder = strings::wstringToUlong(switchFolder.args.front(), 10);
							   if (options->ulFolder)
							   {
								   options->lpszFolderPath = switchFolder.args.front();
								   options->ulFolder = mapi::DEFAULT_INBOX;
							   }

							   return true;
						   }};
	OptParser switchOutput{L"Output", cmdmodeUnknown, 1, 1, OPT_NOOPT, [](auto _options) {
							   auto options = GetMyOptions(_options);
							   options->lpszOutput = switchOutput.args.front();
							   return true;
						   }};
	OptParser switchDispid{L"Dispids", cmdmodePropTag, 0, 0, OPT_DODISPID};
	OptParser switchType{L"Type", cmdmodePropTag, 0, 1, OPT_DOTYPE};
	OptParser switchGuid{L"Guids", cmdmodeGuid, 0, 0, OPT_NOOPT};
	OptParser switchError{L"Error", cmdmodeErr, 0, 0, OPT_NOOPT};
	OptParser switchParser{L"Type", cmdmodeSmartView, 1, 1, OPT_INITMFC | OPT_NEEDINPUTFILE, [](auto _options) {
							   auto options = GetMyOptions(_options);
							   options->ulSVParser = strings::wstringToUlong(switchParser.args.front(), 10);
							   return true;
						   }};
	OptParser switchInput{L"Input", cmdmodeUnknown, 1, 1, OPT_NOOPT, [](auto _options) {
							  auto options = GetMyOptions(_options);
							  options->lpszInput = switchInput.args.front();
							  return true;
						  }};
	OptParser switchBinary{L"Binary", cmdmodeSmartView, 0, 0, OPT_BINARYFILE};
	OptParser switchAcl{L"Acl", cmdmodeAcls, 0, 0, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDFOLDER};
	OptParser switchRule{L"Rules",
						 cmdmodeRules,
						 0,
						 0,
						 OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDFOLDER};
	OptParser switchContents{L"Contents",
							 cmdmodeContents,
							 0,
							 0,
							 OPT_DOCONTENTS | OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC};
	OptParser switchAssociatedContents{L"HiddenContents",
									   cmdmodeContents,
									   0,
									   0,
									   OPT_DOASSOCIATEDCONTENTS | OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC};
	OptParser switchMoreProperties{L"MoreProperties", cmdmodeUnknown, 0, 0, OPT_RETRYSTREAMPROPS};
	OptParser switchNoAddins{L"NoAddins", cmdmodeUnknown, 0, 0, OPT_NOADDINS};
	OptParser switchOnline{L"Online", cmdmodeUnknown, 0, 0, OPT_ONLINE};
	OptParser switchMAPI{L"MAPI",
						 cmdmodeMAPIMIME,
						 0,
						 0,
						 OPT_NEEDMAPIINIT | OPT_INITMFC | OPT_NEEDINPUTFILE | OPT_NEEDOUTPUTFILE,
						 [](auto _options) {
							 auto options = GetMyOptions(_options);
							 options->MAPIMIMEFlags |= MAPIMIME_TOMAPI;
							 return true;
						 }};
	OptParser switchMIME{L"MIME",
						 cmdmodeMAPIMIME,
						 0,
						 0,
						 OPT_NEEDMAPIINIT | OPT_INITMFC | OPT_NEEDINPUTFILE | OPT_NEEDOUTPUTFILE,
						 [](auto _options) {
							 auto options = GetMyOptions(_options);
							 options->MAPIMIMEFlags |= MAPIMIME_TOMIME;
							 return true;
						 }};
	OptParser switchCCSFFlags{L"CCSFFlags", cmdmodeMAPIMIME, 1, 1, OPT_NOOPT, [](auto _options) {
								  auto options = GetMyOptions(_options);
								  options->convertFlags =
									  static_cast<CCSFLAGS>(strings::wstringToUlong(switchCCSFFlags.args.front(), 10));
								  return true;
							  }};
	OptParser switchRFC822{L"RFC822", cmdmodeMAPIMIME, 0, 0, OPT_NOOPT, [](auto _options) {
							   auto options = GetMyOptions(_options);
							   options->MAPIMIMEFlags |= MAPIMIME_RFC822;
							   return true;
						   }};
	OptParser switchWrap{L"Wrap", cmdmodeMAPIMIME, 1, 1, OPT_NOOPT, [](auto _options) {
							 auto options = GetMyOptions(_options);
							 options->ulWrapLines = strings::wstringToUlong(switchWrap.args.front(), 10);
							 options->MAPIMIMEFlags |= MAPIMIME_WRAP;
							 return true;
						 }};
	OptParser switchEncoding{L"Encoding", cmdmodeMAPIMIME, 1, 1, OPT_NOOPT, [](auto _options) {
								 auto options = GetMyOptions(_options);
								 options->ulEncodingType = strings::wstringToUlong(switchEncoding.args.front(), 10);
								 options->MAPIMIMEFlags |= MAPIMIME_ENCODING;
								 return true;
							 }};
	OptParser switchCharset{L"Charset", cmdmodeMAPIMIME, 3, 3, OPT_NOOPT, [](auto _options) {
								auto options = GetMyOptions(_options);
								options->ulCodePage = strings::wstringToUlong(switchCharset.args[0], 10);
								options->cSetType =
									static_cast<CHARSETTYPE>(strings::wstringToUlong(switchCharset.args[1], 10));
								if (options->cSetType > CHARSET_WEB)
								{
									return false;
								}

								options->cSetApplyType =
									static_cast<CSETAPPLYTYPE>(strings::wstringToUlong(switchCharset.args[2], 10));
								if (options->cSetApplyType > CSET_APPLY_TAG_ALL)
								{
									return false;
								}

								options->MAPIMIMEFlags |= MAPIMIME_CHARSET;
								return true;
							}};
	OptParser switchAddressBook{L"AddressBook", cmdmodeMAPIMIME, 0, 0, OPT_NEEDMAPILOGON, [](auto _options) {
									auto options = GetMyOptions(_options);
									options->MAPIMIMEFlags |= MAPIMIME_ADDRESSBOOK;
									return true;
								}}; // special case which needs a logon
	OptParser switchUnicode{L"Unicode", cmdmodeMAPIMIME, 0, 0, OPT_NOOPT, [](auto _options) {
								auto options = GetMyOptions(_options);
								options->MAPIMIMEFlags |= MAPIMIME_UNICODE;
								return true;
							}};
	OptParser switchProfile{L"Profile", cmdmodeUnknown, 0, 1, OPT_PROFILE, [](auto _options) {
								auto options = GetMyOptions(_options);
								if (!switchProfile.args.empty())
								{
									options->lpszProfile = switchProfile.args.front();
								}

								return true;
							}};
	OptParser switchXML{L"XML", cmdmodeXML, 0, 0, OPT_NEEDMAPIINIT | OPT_INITMFC | OPT_NEEDINPUTFILE};
	OptParser switchSubject{L"Subject", cmdmodeContents, 1, 1, OPT_NOOPT};
	OptParser switchMessageClass{L"MessageClass", cmdmodeContents, 1, 1, OPT_NOOPT};
	OptParser switchMSG{L"MSG", cmdmodeContents, 0, 0, OPT_MSG};
	OptParser switchList{L"List", cmdmodeContents, 0, 0, OPT_LIST};
	OptParser switchChildFolders{L"ChildFolders",
								 cmdmodeChildFolders,
								 0,
								 0,
								 OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDFOLDER};
	OptParser switchFid{L"FID",
						cmdmodeFidMid,
						0,
						1,
						OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDSTORE};
	OptParser switchMid{L"MID", cmdmodeFidMid, 0, 1, OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC};
	OptParser switchFlag{L"Flag", cmdmodeUnknown, 1, 1, OPT_NOOPT, [](auto _options) {
							 auto options = GetMyOptions(_options);
							 LPWSTR szEndPtr = nullptr;
							 // We must have a next argument, but it could be a string or a number
							 options->lpszFlagName = switchFlag.args.front();
							 options->ulFlagValue = wcstoul(switchFlag.args.front().c_str(), &szEndPtr, 16);

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
	OptParser switchRecent{L"Recent", cmdmodeContents, 1, 1, OPT_NOOPT, [](auto _options) {
							   auto options = GetMyOptions(_options);
							   options->ulCount = strings::wstringToUlong(switchRecent.args.front(), 10);
							   return true;
						   }};
	OptParser switchStore{L"Store",
						  cmdmodeStoreProperties,
						  0,
						  1,
						  OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC,
						  [](auto _options) {
							  auto options = GetMyOptions(_options);
							  if (!switchStore.args.empty())
							  {
								  LPWSTR szEndPtr = nullptr;
								  options->ulStore = wcstoul(switchStore.args.front().c_str(), &szEndPtr, 10);

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
	OptParser switchVersion{L"Version", cmdmodeUnknown, 1, 1, OPT_NOOPT};
	OptParser switchSize{L"Size",
						 cmdmodeFolderSize,
						 0,
						 0,
						 OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDFOLDER};
	OptParser switchPST{L"PST", cmdmodePST, 0, 0, OPT_NEEDINPUTFILE};
	OptParser switchProfileSection{L"ProfileSection",
								   cmdmodeProfile,
								   1,
								   1,
								   OPT_PROFILE | OPT_NEEDMAPIINIT | OPT_INITMFC};
	OptParser switchByteSwapped{L"ByteSwapped",
								cmdmodeProfile,
								0,
								0,
								OPT_PROFILE | OPT_NEEDMAPIINIT | OPT_INITMFC,
								[](auto _options) {
									auto options = GetMyOptions(_options);
									options->bByteSwapped = true;
									return true;
								}};
	OptParser switchReceiveFolder{L"ReceiveFolder",
								  cmdmodeReceiveFolder,
								  0,
								  1,
								  OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_NEEDSTORE | OPT_INITMFC,
								  [](auto _options) {
									  auto options = GetMyOptions(_options);
									  if (!switchReceiveFolder.args.empty())
									  {
										  LPWSTR szEndPtr = nullptr;
										  options->ulStore =
											  wcstoul(switchReceiveFolder.args.front().c_str(), &szEndPtr, 10);

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
	OptParser switchSkip{L"Skip", cmdmodeUnknown, 0, 0, OPT_SKIPATTACHMENTS};
	OptParser switchSearchState{L"SearchState",
								cmdmodeSearchState,
								0,
								1,
								OPT_NEEDMAPIINIT | OPT_NEEDMAPILOGON | OPT_INITMFC | OPT_NEEDFOLDER};

	// If we want to add aliases for any switches, add them here
	OptParser switchHelpAlias{L"Help", cmdmodeHelpFull, 0, 0, OPT_INITMFC};

	std::vector<OptParser*> g_Parsers = {
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
		printf("   MrMAPI -%ws\n", switchHelp.szSwitch);
		printf(
			"   MrMAPI [-%ws] [-%ws] [-%ws] [-%ws <type>] <property number>|<property name>\n",
			switchSearch.szSwitch,
			switchDispid.szSwitch,
			switchDecimal.szSwitch,
			switchType.szSwitch);
		printf("   MrMAPI -%ws\n", switchGuid.szSwitch);
		printf("   MrMAPI -%ws <error>\n", switchError.szSwitch);
		printf(
			"   MrMAPI -%ws <type> -%ws <input file> [-%ws] [-%ws <output file>]\n",
			switchParser.szSwitch,
			switchInput.szSwitch,
			switchBinary.szSwitch,
			switchOutput.szSwitch);
		printf(
			"   MrMAPI -%ws <flag value> [-%ws] [-%ws] <property number>|<property name>\n",
			switchFlag.szSwitch,
			switchDispid.szSwitch,
			switchDecimal.szSwitch);
		printf("   MrMAPI -%ws <flag name>\n", switchFlag.szSwitch);
		printf(
			"   MrMAPI -%ws [-%ws <profile>] [-%ws <folder>]\n",
			switchRule.szSwitch,
			switchProfile.szSwitch,
			switchFolder.szSwitch);
		printf(
			"   MrMAPI -%ws [-%ws <profile>] [-%ws <folder>]\n",
			switchAcl.szSwitch,
			switchProfile.szSwitch,
			switchFolder.szSwitch);
		printf(
			"   MrMAPI -%ws | -%ws [-%ws <profile>] [-%ws <folder>] [-%ws <output directory>]\n",
			switchContents.szSwitch,
			switchAssociatedContents.szSwitch,
			switchProfile.szSwitch,
			switchFolder.szSwitch,
			switchOutput.szSwitch);
		printf(
			"          [-%ws <subject>] [-%ws <message class>] [-%ws] [-%ws] [-%ws <count>] [-%ws]\n",
			switchSubject.szSwitch,
			switchMessageClass.szSwitch,
			switchMSG.szSwitch,
			switchList.szSwitch,
			switchRecent.szSwitch,
			switchSkip.szSwitch);
		printf(
			"   MrMAPI -%ws [-%ws <profile>] [-%ws <folder>]\n",
			switchChildFolders.szSwitch,
			switchProfile.szSwitch,
			switchFolder.szSwitch);
		printf(
			"   MrMAPI -%ws -%ws <path to input file> -%ws <path to output file> [-%ws]\n",
			switchXML.szSwitch,
			switchInput.szSwitch,
			switchOutput.szSwitch,
			switchSkip.szSwitch);
		printf(
			"   MrMAPI -%ws [fid] [-%ws [mid]] [-%ws <profile>]\n",
			switchFid.szSwitch,
			switchMid.szSwitch,
			switchProfile.szSwitch);
		printf(
			"   MrMAPI [<property number>|<property name>] -%ws [<store num>] [-%ws <profile>]\n",
			switchStore.szSwitch,
			switchProfile.szSwitch);
		printf(
			"   MrMAPI [<property number>|<property name>] -%ws <folder> [-%ws <profile>]\n",
			switchFolder.szSwitch,
			switchProfile.szSwitch);
		printf(
			"   MrMAPI -%ws -%ws <folder> [-%ws <profile>]\n",
			switchSize.szSwitch,
			switchFolder.szSwitch,
			switchProfile.szSwitch);
		printf(
			"   MrMAPI -%ws | -%ws -%ws <path to input file> -%ws <path to output file> [-%ws <conversion flags>]\n",
			switchMAPI.szSwitch,
			switchMIME.szSwitch,
			switchInput.szSwitch,
			switchOutput.szSwitch,
			switchCCSFFlags.szSwitch);
		printf(
			"          [-%ws] [-%ws <Decimal number of characters>] [-%ws <Decimal number indicating encoding>]\n",
			switchRFC822.szSwitch,
			switchWrap.szSwitch,
			switchEncoding.szSwitch);
		printf(
			"          [-%ws] [-%ws] [-%ws CodePage CharSetType CharSetApplyType]\n",
			switchAddressBook.szSwitch,
			switchUnicode.szSwitch,
			switchCharset.szSwitch);
		printf("   MrMAPI -%ws -%ws <path to input file>\n", switchPST.szSwitch, switchInput.szSwitch);
		printf(
			"   MrMAPI -%ws [<profile> [-%ws <profilesection> [-%ws]] -%ws <output file>]\n",
			switchProfile.szSwitch,
			switchProfileSection.szSwitch,
			switchByteSwapped.szSwitch,
			switchOutput.szSwitch);
		printf("   MrMAPI -%ws [<store num>] [-%ws <profile>]\n", switchReceiveFolder.szSwitch, switchProfile.szSwitch);
		printf(
			"   MrMAPI -%ws -%ws <folder> [-%ws <profile>]\n",
			switchSearchState.szSwitch,
			switchFolder.szSwitch,
			switchProfile.szSwitch);

		if (bFull)
		{
			printf("\n");
			printf("All switches may be shortened if the intended switch is unambiguous.\n");
			printf("For example, -T may be used instead of -%ws.\n", switchType.szSwitch);
		}
		printf("\n");
		printf("   Help:\n");
		printf("   -%ws   Display expanded help.\n", switchHelp.szSwitch);
		if (bFull)
		{
			printf("\n");
			printf("   Property Tag Lookup:\n");
			printf("   -S   (or -%ws) Perform substring search.\n", switchSearch.szSwitch);
			printf("           With no parameters prints all known properties.\n");
			printf("   -D   (or -%ws) Search dispids.\n", switchDispid.szSwitch);
			printf("   -N   (or -%ws) Number is in decimal. Ignored for non-numbers.\n", switchDecimal.szSwitch);
			printf("   -T   (or -%ws) Print information on specified type.\n", switchType.szSwitch);
			printf("           With no parameters prints list of known types.\n");
			printf("           When combined with -S, restrict output to given type.\n");
			printf("   -G   (or -%ws) Display list of known guids.\n", switchGuid.szSwitch);
			printf("\n");
			printf("   Flag Lookup:\n");
			printf("   -Fl  (or -%ws) Look up flags for specified property.\n", switchFlag.szSwitch);
			printf("           May be combined with -D and -N switches, but all flag values must be in hex.\n");
			printf("   -Fl  (or -%ws) Look up flag name and output its value.\n", switchFlag.szSwitch);
			printf("\n");
			printf("   Error Parsing:\n");
			printf("   -E   (or -%ws) Map an error code to its name and vice versa.\n", switchError.szSwitch);
			printf("           May be combined with -S and -N switches.\n");
			printf("\n");
			printf("   Smart View Parsing:\n");
			printf(
				"   -P   (or -%ws) Parser type (number). See list below for supported parsers.\n",
				switchParser.szSwitch);
			printf("   -B   (or -%ws) Input file is binary. Default is hex encoded text.\n", switchBinary.szSwitch);
			printf("\n");
			printf("   Rules Table:\n");
			printf("   -R   (or -%ws) Output rules table. Profile optional.\n", switchRule.szSwitch);
			printf("\n");
			printf("   ACL Table:\n");
			printf("   -A   (or -%ws) Output ACL table. Profile optional.\n", switchAcl.szSwitch);
			printf("\n");
			printf("   Contents Table:\n");
			printf(
				"   -C   (or -%ws) Output contents table. May be combined with -H. Profile optional.\n",
				switchContents.szSwitch);
			printf(
				"   -H   (or -%ws) Output associated contents table. May be combined with -C. Profile optional\n",
				switchAssociatedContents.szSwitch);
			printf("   -Su  (or -%ws) Subject of messages to output.\n", switchSubject.szSwitch);
			printf("   -Me  (or -%ws) Message class of messages to output.\n", switchMessageClass.szSwitch);
			printf("   -Ms  (or -%ws) Output as .MSG instead of XML.\n", switchMSG.szSwitch);
			printf("   -L   (or -%ws) List details to screen and do not output files.\n", switchList.szSwitch);
			printf("   -Re  (or -%ws) Restrict output to the 'count' most recent messages.\n", switchRecent.szSwitch);
			printf("\n");
			printf("   Child Folders:\n");
			printf("   -Chi (or -%ws) Display child folders of selected folder.\n", switchChildFolders.szSwitch);
			printf("\n");
			printf("   MSG File Properties\n");
			printf("   -X   (or -%ws) Output properties of an MSG file as XML.\n", switchXML.szSwitch);
			printf("\n");
			printf("   MID/FID Lookup\n");
			printf("   -Fi  (or -%ws) Folder ID (FID) to search for.\n", switchFid.szSwitch);
			printf("           If -%ws is specified without a FID, search/display all folders\n", switchFid.szSwitch);
			printf("   -Mid (or -%ws) Message ID (MID) to search for.\n", switchMid.szSwitch);
			printf(
				"           If -%ws is specified without a MID, display all messages in folders specified by the FID "
				"parameter.\n",
				switchMid.szSwitch);
			printf("\n");
			printf("   Store Properties\n");
			printf("   -St  (or -%ws) Output properties of stores as XML.\n", switchStore.szSwitch);
			printf("           If store number is specified, outputs properties of a single store.\n");
			printf("           If a property is specified, outputs only that property.\n");
			printf("\n");
			printf("   Folder Properties\n");
			printf("   -F   (or -%ws) Output properties of a folder as XML.\n", switchFolder.szSwitch);
			printf("           If a property is specified, outputs only that property.\n");
			printf("   -Size         Output size of a folder and all subfolders.\n");
			printf("           Use -%ws to specify which folder to scan.\n", switchFolder.szSwitch);
			printf("   -SearchState  Output search folder state.\n");
			printf("           Use -%ws to specify which folder to scan.\n", switchFolder.szSwitch);
			printf("\n");
			printf("   MAPI <-> MIME Conversion:\n");
			printf("   -Ma  (or -%ws) Convert an EML file to MAPI format (MSG file).\n", switchMAPI.szSwitch);
			printf("   -Mi  (or -%ws) Convert an MSG file to MIME format (EML file).\n", switchMIME.szSwitch);
			printf(
				"   -I   (or -%ws) Indicates the input file for conversion, either a MIME-formatted EML file or an MSG "
				"file.\n",
				switchInput.szSwitch);
			printf("   -O   (or -%ws) Indicates the output file for the conversion.\n", switchOutput.szSwitch);
			printf("   -Cc  (or -%ws) Indicates specific flags to pass to the converter.\n", switchCCSFFlags.szSwitch);
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
				switchRFC822.szSwitch);
			printf("           If not present, RFC1521 is used instead.\n");
			printf(
				"   -W   (or -%ws) (MAPI->MIME only) Indicates the maximum number of characters in each line in the\n",
				switchWrap.szSwitch);
			printf("           generated EML. Default value is 74. A value of 0 indicates no wrapping.\n");
			printf(
				"   -En  (or -%ws) (MAPI->MIME only) Indicates the encoding type to use. Supported values are:\n",
				switchEncoding.szSwitch);
			printf("              1 - Base64\n");
			printf("              2 - UUENCODE\n");
			printf("              3 - Quoted-Printable\n");
			printf("              4 - 7bit (DEFAULT)\n");
			printf("              5 - 8bit\n");
			printf(
				"   -Ad  (or -%ws) Pass MAPI Address Book into converter. Profile optional.\n",
				switchAddressBook.szSwitch);
			printf(
				"   -U   (or -%ws) (MIME->MAPI only) The resulting MSG file should be unicode.\n",
				switchUnicode.szSwitch);
			printf(
				"   -Ch  (or -%ws) (MIME->MAPI only) Character set - three required parameters:\n",
				switchCharset.szSwitch);
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
			printf("   -I   (or -%ws) PST file to be analyzed.\n", switchInput.szSwitch);
			printf("\n");
			printf("   Profiles\n");
			printf("   -Pr  (or -%ws) Output list of profiles\n", switchProfile.szSwitch);
			printf("           If a profile is specified, exports that profile.\n");
			printf("   -ProfileSection If specified, output specific profile section.\n");
			printf(
				"   -B   (or -%ws) If specified, profile section guid is byte swapped.\n", switchByteSwapped.szSwitch);
			printf("   -O   (or -%ws) Indicates the output file for profile export.\n", switchOutput.szSwitch);
			printf("           Required if a profile is specified.\n");
			printf("\n");
			printf("   Receive Folder Table\n");
			printf("   -%ws Displays Receive Folder Table for the specified store\n", switchReceiveFolder.szSwitch);
			printf("\n");
			printf("   Universal Options:\n");
			printf("   -I   (or -%ws) Input file.\n", switchInput.szSwitch);
			printf("   -O   (or -%ws) Output file or directory.\n", switchOutput.szSwitch);
			printf(
				"   -F   (or -%ws) Folder to scan. Default is Inbox. See list below for supported folders.\n",
				switchFolder.szSwitch);
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
			printf("   -Pr  (or -%ws) Profile for MAPILogonEx.\n", switchProfile.szSwitch);
			printf(
				"   -M   (or -%ws) More properties. Tries harder to get stream properties. May take longer.\n",
				switchMoreProperties.szSwitch);
			printf("   -Sk  (or -%ws) Skip embedded message attachments on export.\n", switchSkip.szSwitch);
			printf("   -No  (or -%ws) No Addins. Don't load any add-ins.\n", switchNoAddins.szSwitch);
			printf("   -On  (or -%ws) Online mode. Bypass cached mode.\n", switchOnline.szSwitch);
			printf("   -V   (or -%ws) Verbose. Turn on all debug output.\n", switchVerbose.szSwitch);
			printf("\n");
			printf("   MAPI Implementation Options:\n");
			printf("   -%ws MAPI Version to load - supported values\n", switchVersion.szSwitch);
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
				options->options & OPT_DOPARTIALSEARCH && options->options & OPT_DOTYPE && !switchType.args.empty())
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
			else if (!switchProfileSection.hasArgs() && options->lpszProfile.empty())
				options->mode = cmdmodeHelp;

			break;
		default:
			break;
		}
	}
} // namespace cli