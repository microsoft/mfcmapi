#include <StdAfx.h>
#include <MrMapi/MMContents.h>
#include <core/mapi/processor/dumpStore.h>
#include <MrMapi/mmcli.h>
#include <core/mapi/mapiOutput.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFile.h>

void DumpContentsTable(
	_In_z_ LPCWSTR lpszProfile,
	_In_ LPMDB lpMDB,
	_In_ LPMAPIFOLDER lpFolder,
	_In_z_ LPWSTR lpszDir,
	_In_z_ LPCWSTR lpszFolder,
	_In_ ULONG ulCount,
	_In_opt_ LPSRestriction lpRes)
{
	output::DebugPrint(
		DBGGeneric,
		L"DumpContentsTable: Outputting folder %ws from profile %ws to %ws\n",
		lpszFolder,
		lpszProfile,
		lpszDir);
	if (cli::switchContents.isSet()) output::DebugPrint(DBGGeneric, L"DumpContentsTable: Outputting Contents\n");
	if (cli::switchAssociatedContents.isSet())
		output::DebugPrint(DBGGeneric, L"DumpContentsTable: Outputting Associated Contents\n");
	if (cli::switchMSG.isSet()) output::DebugPrint(DBGGeneric, L"DumpContentsTable: Outputting as MSG\n");
	if (cli::switchMoreProperties.isSet())
		output::DebugPrint(DBGGeneric, L"DumpContentsTable: Will retry stream properties\n");
	if (cli::switchSkip.isSet()) output::DebugPrint(DBGGeneric, L"DumpContentsTable: Will skip attachments\n");
	if (cli::switchList.isSet()) output::DebugPrint(DBGGeneric, L"DumpContentsTable: List only mode\n");
	if (ulCount) output::DebugPrint(DBGGeneric, L"DumpContentsTable: Limiting output to %u messages.\n", ulCount);

	if (lpFolder)
	{
		mapi::processor::dumpStore MyDumpStore;
		SSortOrderSet SortOrder = {0};
		MyDumpStore.InitMDB(lpMDB);
		MyDumpStore.InitFolder(lpFolder);
		MyDumpStore.InitFolderPathRoot(lpszDir);
		MyDumpStore.InitFolderContentsRestriction(lpRes);
		if (cli::switchMSG.isSet()) MyDumpStore.EnableMSG();
		if (cli::switchList.isSet()) MyDumpStore.EnableList();
		if (ulCount)
		{
			MyDumpStore.InitMaxOutput(ulCount);
			SortOrder.cSorts = 1;
			SortOrder.cCategories = 0;
			SortOrder.cExpanded = 0;
			SortOrder.aSort[0].ulPropTag = PR_MESSAGE_DELIVERY_TIME;
			SortOrder.aSort[0].ulOrder = TABLE_SORT_DESCEND;
			MyDumpStore.InitSortOrder(&SortOrder);
		}

		if (!(cli::switchMoreProperties.isSet())) MyDumpStore.DisableStreamRetry();
		if (cli::switchSkip.isSet()) MyDumpStore.DisableEmbeddedAttachments();

		MyDumpStore.ProcessFolders(cli::switchContents.isSet(), cli::switchAssociatedContents.isSet(), false);
	}
}

void DumpMSG(
	_In_z_ LPCWSTR lpszMSGFile,
	_In_z_ LPCWSTR lpszXMLFile,
	_In_ bool bRetryStreamProps,
	_In_ bool bOutputAttachments)
{
	auto lpMessage = file::LoadMSGToMessage(lpszMSGFile);
	if (lpMessage)
	{
		mapi::processor::dumpStore MyDumpStore;
		MyDumpStore.InitMessagePath(lpszXMLFile);
		if (!bRetryStreamProps) MyDumpStore.DisableStreamRetry();
		if (!bOutputAttachments) MyDumpStore.DisableEmbeddedAttachments();

		// Just assume this message might have attachments
		MyDumpStore.ProcessMessage(lpMessage, true, nullptr);
		lpMessage->Release();
	}
}

void DoContents(LPMDB lpMDB, LPMAPIFOLDER lpFolder)
{
	SRestriction sResTop = {};
	SRestriction sResMiddle[2] = {};
	SRestriction sResSubject[2] = {};
	SRestriction sResMessageClass[2] = {};
	SPropValue sPropValue[2] = {};
	LPSRestriction lpRes = nullptr;
	if (cli::switchSubject.hasArgs() || cli::switchMessageClass.hasArgs())
	{
		// RES_AND
		//   RES_AND (optional)
		//     RES_EXIST - PR_SUBJECT_W
		//     RES_CONTENT - lpszSubject
		//   RES_AND (optional)
		//     RES_EXIST - PR_MESSAGE_CLASS_W
		//     RES_CONTENT - lpszMessageClass
		auto i = 0;
		const auto szSubject = cli::switchSubject.getArg(0);
		if (!szSubject.empty())
		{
			sResMiddle[i].rt = RES_AND;
			sResMiddle[i].res.resAnd.cRes = 2;
			sResMiddle[i].res.resAnd.lpRes = &sResSubject[0];
			sResSubject[0].rt = RES_EXIST;
			sResSubject[0].res.resExist.ulPropTag = PR_SUBJECT_W;
			sResSubject[1].rt = RES_CONTENT;
			sResSubject[1].res.resContent.ulPropTag = PR_SUBJECT_W;
			sResSubject[1].res.resContent.ulFuzzyLevel = FL_FULLSTRING | FL_IGNORECASE;
			sResSubject[1].res.resContent.lpProp = &sPropValue[0];
			sPropValue[0].ulPropTag = PR_SUBJECT_W;
			sPropValue[0].Value.lpszW = const_cast<LPWSTR>(szSubject.c_str());
			i++;
		}

		const auto szMessageClass = cli::switchMessageClass.getArg(0);
		if (!szMessageClass.empty())
		{
			sResMiddle[i].rt = RES_AND;
			sResMiddle[i].res.resAnd.cRes = 2;
			sResMiddle[i].res.resAnd.lpRes = &sResMessageClass[0];
			sResMessageClass[0].rt = RES_EXIST;
			sResMessageClass[0].res.resExist.ulPropTag = PR_MESSAGE_CLASS_W;
			sResMessageClass[1].rt = RES_CONTENT;
			sResMessageClass[1].res.resContent.ulPropTag = PR_MESSAGE_CLASS_W;
			sResMessageClass[1].res.resContent.ulFuzzyLevel = FL_FULLSTRING | FL_IGNORECASE;
			sResMessageClass[1].res.resContent.lpProp = &sPropValue[1];
			sPropValue[1].ulPropTag = PR_MESSAGE_CLASS_W;
			sPropValue[1].Value.lpszW = const_cast<LPWSTR>(szMessageClass.c_str());
			i++;
		}

		sResTop.rt = RES_AND;
		sResTop.res.resAnd.cRes = i;
		sResTop.res.resAnd.lpRes = &sResMiddle[0];
		lpRes = &sResTop;
		output::outputRestriction(DBGGeneric, nullptr, lpRes, nullptr);
	}

	DumpContentsTable(
		cli::switchProfile.getArg(0).c_str(),
		lpMDB,
		lpFolder,
		cli::switchOutput.hasArgs() ? cli::switchOutput.getArg(0).c_str() : L".",
		cli::switchFolder.getArg(0).c_str(),
		cli::switchRecent.getArgAsULONG(0),
		lpRes);
}

void DoMSG(_In_ cli::OPTIONS ProgOpts)
{
	DumpMSG(
		cli::switchInput.getArg(0).c_str(),
		cli::switchOutput.hasArgs() ? cli::switchOutput.getArg(0).c_str() : L".",
		cli::switchMoreProperties.isSet(),
		!cli::switchSkip.isSet());
}