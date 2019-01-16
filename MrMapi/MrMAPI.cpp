#include <StdAfx.h>

#include <MrMapi/cli.h>
#include <MrMapi/MrMAPI.h>
#include <core/utility/strings.h>
#include <Interpret/InterpretProp.h>
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
#include <ImportProcs.h>
#include <MAPI/MAPIStoreFunctions.h>
#include <MrMapi/MMPst.h>
#include <MrMapi/MMReceiveFolder.h>
#include <MAPI/Cache/NamedPropCache.h>
#include <MAPI/StubUtils.h>

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

	SetDllDirectory(_T(""));
	import::MyHeapSetInformation(nullptr, HeapEnableTerminationOnCorruption, nullptr, 0);

	// Set up our property arrays or nothing works
	addin::MergeAddInArrays();

	registry::doSmartView = true;
	registry::useGetPropList = true;
	registry::parseNamedProps = true;
	registry::cacheNamedProps = true;

	auto cl = cli::GetCommandLine(argc, argv);
	auto ProgOpts = cli::ParseArgs(cl);

	// Must be first after ParseArgs
	if (ProgOpts.ulOptions & cli::OPT_INITMFC)
	{
		InitMFC();
	}

	if (ProgOpts.ulOptions & cli::OPT_VERBOSE)
	{
		registry::debugTag = 0xFFFFFFFF;
		PrintArgs(ProgOpts);
	}

	if (!(ProgOpts.ulOptions & cli::OPT_NOADDINS))
	{
		registry::loadAddIns = true;
		addin::LoadAddIns();
	}

	if (!ProgOpts.lpszVersion.empty())
	{
		if (LoadMAPIVersion(ProgOpts.lpszVersion)) return;
	}

	if (ProgOpts.Mode == cli::cmdmodeHelp)
	{
		cli::DisplayUsage(false);
	}
	else if (ProgOpts.Mode == cli::cmdmodeHelpFull)
	{
		cli::DisplayUsage(true);
	}
	else
	{
		if (ProgOpts.ulOptions & cli::OPT_ONLINE)
		{
			registry::forceMapiNoCache = true;
			registry::forceMDBOnline = true;
		}

		// Log on to MAPI if needed
		if (ProgOpts.ulOptions & cli::OPT_NEEDMAPIINIT)
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

		if (bMAPIInit && ProgOpts.ulOptions & cli::OPT_NEEDMAPILOGON)
		{
			ProgOpts.lpMAPISession = MrMAPILogonEx(ProgOpts.lpszProfile);
		}

		// If they need a folder get it and store at the same time from the folder id
		if (ProgOpts.lpMAPISession && ProgOpts.ulOptions & cli::OPT_NEEDFOLDER)
		{
			hRes = WC_H(HrMAPIOpenStoreAndFolder(
				ProgOpts.lpMAPISession,
				ProgOpts.ulFolder,
				ProgOpts.lpszFolderPath,
				&ProgOpts.lpMDB,
				&ProgOpts.lpFolder));
			if (FAILED(hRes)) printf("HrMAPIOpenStoreAndFolder returned an error: 0x%08lx\n", hRes);
		}
		else if (ProgOpts.lpMAPISession && ProgOpts.ulOptions & cli::OPT_NEEDSTORE)
		{
			// They asked us for a store, if they passed a store index give them that one
			if (ProgOpts.ulStore != 0)
			{
				// Decrement by one here on the index since we incremented during parameter parsing
				// This is so zero indicate they did not specify a store
				ProgOpts.lpMDB = OpenStore(ProgOpts.lpMAPISession, ProgOpts.ulStore - 1);
				if (!ProgOpts.lpMDB) printf("OpenStore failed\n");
			}
			else
			{
				// If they needed a store but didn't specify, get the default one
				ProgOpts.lpMDB = OpenExchangeOrDefaultMessageStore(ProgOpts.lpMAPISession);
				if (!ProgOpts.lpMDB) printf("OpenExchangeOrDefaultMessageStore failed.\n");
			}
		}

		switch (ProgOpts.Mode)
		{
		case cli::cmdmodePropTag:
			DoPropTags(ProgOpts);
			break;
		case cli::cmdmodeGuid:
			DoGUIDs(ProgOpts);
			break;
		case cli::cmdmodeSmartView:
			DoSmartView(ProgOpts);
			break;
		case cli::cmdmodeAcls:
			DoAcls(ProgOpts);
			break;
		case cli::cmdmodeRules:
			DoRules(ProgOpts);
			break;
		case cli::cmdmodeErr:
			DoErrorParse(ProgOpts);
			break;
		case cli::cmdmodeContents:
			DoContents(ProgOpts);
			break;
		case cli::cmdmodeXML:
			DoMSG(ProgOpts);
			break;
		case cli::cmdmodeFidMid:
			DoFidMid(ProgOpts);
			break;
		case cli::cmdmodeStoreProperties:
			DoStore(ProgOpts);
			break;
		case cli::cmdmodeMAPIMIME:
			DoMAPIMIME(ProgOpts);
			break;
		case cli::cmdmodeChildFolders:
			DoChildFolders(ProgOpts);
			break;
		case cli::cmdmodeFlagSearch:
			DoFlagSearch(ProgOpts);
			break;
		case cli::cmdmodeFolderProps:
			DoFolderProps(ProgOpts);
			break;
		case cli::cmdmodeFolderSize:
			DoFolderSize(ProgOpts);
			break;
		case cli::cmdmodePST:
			DoPST(ProgOpts);
			break;
		case cli::cmdmodeProfile:
			output::DoProfile(ProgOpts);
			break;
		case cli::cmdmodeReceiveFolder:
			DoReceiveFolder(ProgOpts);
			break;
		case cli::cmdmodeSearchState:
			DoSearchState(ProgOpts);
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
		if (ProgOpts.lpFolder) ProgOpts.lpFolder->Release();
		if (ProgOpts.lpMDB) ProgOpts.lpMDB->Release();
		if (ProgOpts.lpMAPISession) ProgOpts.lpMAPISession->Release();
		MAPIUninitialize();
	}

	if (!(ProgOpts.ulOptions & cli::OPT_NOADDINS))
	{
		addin::UnloadAddIns();
	}
}