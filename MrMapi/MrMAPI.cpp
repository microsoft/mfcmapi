#include <StdAfx.h>
#include <MrMapi/MrMAPI.h>
#include <MrMapi/mmcli.h>
#include <MrMapi/MMAcls.h>
#include <MrMapi/MMContents.h>
#include <MrMapi/MMErr.h>
#include <MrMapi/MMFidMid.h>
#include <MrMapi/MMFolder.h>
#include <MrMapi/MMProfile.h>
#include <MrMapi/MMPropTag.h>
#include <MrMapi/MMRules.h>
#include <MrMapi/MMSmartView.h>
#include <MrMapi/MMStore.h>
#include <MrMapi/MMMapiMime.h>
#include <MrMapi/MMPst.h>
#include <MrMapi/MMReceiveFolder.h>
#include <core/utility/strings.h>
#include <core/utility/import.h>
#include <core/mapi/mapiStoreFunctions.h>
#include <core/mapi/cache/namedPropCache.h>
#include <core/mapi/stubutils.h>
#include <core/addin/addin.h>
#include <core/utility/registry.h>
#include <core/utility/output.h>
#include <core/utility/cli.h>

// Initialize MFC for LoadString support later on
void InitMFC() { AfxWinInit(::GetModuleHandle(nullptr), nullptr, ::GetCommandLine(), 0); }

_Check_return_ LPMAPISESSION MrMAPILogonEx(const std::wstring& lpszProfile)
{
	auto ulFlags = MAPI_EXTENDED | MAPI_NO_MAIL | MAPI_UNICODE | MAPI_NEW_SESSION;
	if (lpszProfile.empty()) ulFlags |= MAPI_USE_DEFAULT;

	// TODO: profile parameter should be ansi in ansi builds
	LPMAPISESSION lpSession = nullptr;
	const auto hRes = WC_MAPI(
		MAPILogonEx(NULL, LPTSTR((lpszProfile.empty() ? NULL : lpszProfile.c_str())), NULL, ulFlags, &lpSession));
	if (FAILED(hRes)) printf("MAPILogonEx returned an error: 0x%08lx\n", hRes);
	return lpSession;
}

_Check_return_ LPMDB OpenExchangeOrDefaultMessageStore(_In_ LPMAPISESSION lpMAPISession)
{
	if (!lpMAPISession) return nullptr;

	auto lpMDB = mapi::store::OpenMessageStoreGUID(lpMAPISession, pbExchangeProviderPrimaryUserGuid);
	if (!lpMDB)
	{
		lpMDB = mapi::store::OpenDefaultMessageStore(lpMAPISession);
	}

	return lpMDB;
}

// Returns true if we've done everything we need to do and can exit the program.
// Returns false to continue work.
bool LoadMAPIVersion(const std::wstring& lpszVersion)
{
	// Load DLLS and get functions from them
	import::ImportProcs();
	output::DebugPrint(DBGGeneric, L"LoadMAPIVersion(%ws)\n", lpszVersion.c_str());

	std::wstring szPath;
	auto paths = mapistub::GetMAPIPaths();
	if (lpszVersion == L"0")
	{
		output::DebugPrint(DBGGeneric, L"Listing MAPI\n");
		for (const auto& path : paths)
		{

			printf("MAPI path: %ws\n", strings::wstringToLower(path).c_str());
		}
		return true;
	}

	const auto ulVersion = strings::wstringToUlong(lpszVersion, 10);
	if (ulVersion == 0)
	{
		output::DebugPrint(DBGGeneric, L"Got a string\n");

		for (const auto& path : paths)
		{
			strings::wstringToLower(path);

			if (strings::wstringToLower(path).find(strings::wstringToLower(lpszVersion)) != std::wstring::npos)
			{
				szPath = path;
				break;
			}
		}
	}
	else
	{
		output::DebugPrint(DBGGeneric, L"Got a number %u\n", ulVersion);
		switch (ulVersion)
		{
		case 1: // system
			szPath = mapistub::GetMAPISystemDir();
			break;
		case 11: // Outlook 2003 (11)
			szPath = mapistub::GetInstalledOutlookMAPI(oqcOffice11);
			break;
		case 12: // Outlook 2007 (12)
			szPath = mapistub::GetInstalledOutlookMAPI(oqcOffice12);
			break;
		case 14: // Outlook 2010 (14)
			szPath = mapistub::GetInstalledOutlookMAPI(oqcOffice14);
			break;
		case 15: // Outlook 2013 (15)
			szPath = mapistub::GetInstalledOutlookMAPI(oqcOffice15);
			break;
		case 16: // Outlook 2016 (16)
			szPath = mapistub::GetInstalledOutlookMAPI(oqcOffice16);
			break;
		}
	}

	if (!szPath.empty())
	{
		output::DebugPrint(DBGGeneric, L"Found MAPI path %ws\n", szPath.c_str());
		const auto hMAPI = WC_D(HMODULE, import::MyLoadLibraryW(szPath));
		mapistub::SetMAPIHandle(hMAPI);
	}

	return false;
}

void main(_In_ int argc, _In_count_(argc) char* argv[])
{
	auto hRes = S_OK;
	auto bMAPIInit = false;
	LPMAPISESSION lpMAPISession{};
	LPMDB lpMDB{};
	LPMAPIFOLDER lpFolder{};

	registry::doSmartView = true;
	registry::useGetPropList = true;
	registry::parseNamedProps = true;
	registry::cacheNamedProps = true;
	registry::debugTag = 0;

	SetDllDirectory(_T(""));
	import::MyHeapSetInformation(nullptr, HeapEnableTerminationOnCorruption, nullptr, 0);

	// Set up our property arrays or nothing works
	addin::MergeAddInArrays();

	auto ProgOpts = cli::OPTIONS{};
	auto cl = cli::GetCommandLine(argc, argv);
	cli::ParseArgs(ProgOpts, cl, cli::g_options, cli::PostParseCheck);

	// Must be first after ParseArgs
	if (ProgOpts.options & cli::OPT_INITMFC)
	{
		InitMFC();
	}

	if (ProgOpts.options & cli::OPT_VERBOSE)
	{
		registry::debugTag = 0xFFFFFFFF;
		cli::PrintArgs(ProgOpts, cli::g_options);
	}

	if (!(ProgOpts.options & cli::OPT_NOADDINS))
	{
		registry::loadAddIns = true;
		addin::LoadAddIns();
	}

	if (cli::switchVersion.isSet())
	{
		if (LoadMAPIVersion(cli::switchVersion.getArg(0))) return;
	}

	if (ProgOpts.mode == cli::cmdmodeHelp)
	{
		cli::DisplayUsage(false);
	}
	else if (ProgOpts.mode == cli::cmdmodeHelpFull)
	{
		cli::DisplayUsage(true);
	}
	else
	{
		if (ProgOpts.options & cli::OPT_ONLINE)
		{
			registry::forceMapiNoCache = true;
			registry::forceMDBOnline = true;
		}

		// Log on to MAPI if needed
		if (ProgOpts.options & cli::OPT_NEEDMAPIINIT)
		{
			hRes = WC_MAPI(MAPIInitialize(NULL));
			if (FAILED(hRes))
			{
				printf("Error initializing MAPI: 0x%08lx\n", hRes);
			}
			else
			{
				bMAPIInit = true;
			}
		}

		if (bMAPIInit && ProgOpts.options & cli::OPT_NEEDMAPILOGON)
		{
			lpMAPISession = MrMAPILogonEx(cli::switchProfile.getArg(0));
		}

		// If they need a folder get it and store at the same time from the folder id
		if (lpMAPISession && ProgOpts.options & cli::OPT_NEEDFOLDER)
		{
			hRes = WC_H(HrMAPIOpenStoreAndFolder(lpMAPISession, cli::switchFolder.getArg(0), &lpMDB, &lpFolder));
			if (FAILED(hRes)) printf("HrMAPIOpenStoreAndFolder returned an error: 0x%08lx\n", hRes);
		}

		// If they passed a store index then open it
		ULONG ulStoreIndex{};
		auto gotStoreIndex = strings::tryWstringToUlong(ulStoreIndex, cli::switchStore.getArg(0), 10);
		if (!gotStoreIndex)
			gotStoreIndex = strings::tryWstringToUlong(ulStoreIndex, cli::switchReceiveFolder.getArg(0), 10);

		if (gotStoreIndex)
		{
			lpMDB = OpenStore(lpMAPISession, ulStoreIndex);
			if (!lpMDB) printf("OpenStore failed\n");
		}

		if (lpMAPISession && ProgOpts.options & cli::OPT_NEEDSTORE && !lpMDB)
		{
			// If they needed a store but didn't specify, get the default one
			lpMDB = OpenExchangeOrDefaultMessageStore(lpMAPISession);
			if (!lpMDB) printf("OpenExchangeOrDefaultMessageStore failed.\n");
		}

		switch (ProgOpts.mode)
		{
		case cli::cmdmodePropTag:
			DoPropTags(ProgOpts);
			break;
		case cli::cmdmodeGuid:
			DoGUIDs();
			break;
		case cli::cmdmodeSmartView:
			DoSmartView(ProgOpts);
			break;
		case cli::cmdmodeAcls:
			DoAcls(lpFolder);
			break;
		case cli::cmdmodeRules:
			DoRules(lpFolder);
			break;
		case cli::cmdmodeErr:
			DoErrorParse(ProgOpts);
			break;
		case cli::cmdmodeContents:
			DoContents(ProgOpts, lpMDB, lpFolder);
			break;
		case cli::cmdmodeXML:
			DoMSG(ProgOpts);
			break;
		case cli::cmdmodeFidMid:
			DoFidMid(lpMDB);
			break;
		case cli::cmdmodeStoreProperties:
			DoStore(ProgOpts, lpMAPISession, lpMDB);
			break;
		case cli::cmdmodeMAPIMIME:
			DoMAPIMIME(lpMAPISession);
			break;
		case cli::cmdmodeChildFolders:
			DoChildFolders(lpFolder);
			break;
		case cli::cmdmodeFlagSearch:
			DoFlagSearch();
			break;
		case cli::cmdmodeFolderProps:
			DoFolderProps(ProgOpts, lpFolder);
			break;
		case cli::cmdmodeFolderSize:
			DoFolderSize(lpFolder);
			break;
		case cli::cmdmodePST:
			DoPST(ProgOpts);
			break;
		case cli::cmdmodeProfile:
			output::DoProfile(ProgOpts);
			break;
		case cli::cmdmodeReceiveFolder:
			DoReceiveFolder(lpMDB);
			break;
		case cli::cmdmodeSearchState:
			DoSearchState(lpFolder);
			break;
		case cli::cmdmodeUnknown:
			break;
		case cli::cmdmodeHelp:
			break;
		default:
			break;
		}
	}

	cache::UninitializeNamedPropCache();

	if (bMAPIInit)
	{
		if (lpFolder) lpFolder->Release();
		if (lpMDB) lpMDB->Release();
		if (lpMAPISession) lpMAPISession->Release();
		MAPIUninitialize();
	}

	if (!(ProgOpts.options & cli::OPT_NOADDINS))
	{
		addin::UnloadAddIns();
	}
}