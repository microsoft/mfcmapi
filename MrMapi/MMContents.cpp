#include <StdAfx.h>
#include <MAPI/MAPIProcessor/DumpStore.h>
#include <MrMapi/cli.h>
#include <core/mapi/mapiOutput.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFile.h>

void DumpContentsTable(
	_In_z_ LPCWSTR lpszProfile,
	_In_ LPMDB lpMDB,
	_In_ LPMAPIFOLDER lpFolder,
	_In_z_ LPWSTR lpszDir,
	_In_ ULONG ulOptions,
	_In_ ULONG ulFolder,
	_In_z_ LPCWSTR lpszFolder,
	_In_ ULONG ulCount,
	_In_opt_ LPSRestriction lpRes)
{
	output::DebugPrint(
		DBGGeneric,
		L"DumpContentsTable: Outputting folder %u / %ws from profile %ws to %ws\n",
		ulFolder,
		lpszFolder ? lpszFolder : L"",
		lpszProfile,
		lpszDir);
	if (ulOptions & cli::OPT_DOCONTENTS) output::DebugPrint(DBGGeneric, L"DumpContentsTable: Outputting Contents\n");
	if (ulOptions & cli::OPT_DOASSOCIATEDCONTENTS)
		output::DebugPrint(DBGGeneric, L"DumpContentsTable: Outputting Associated Contents\n");
	if (ulOptions & cli::OPT_MSG) output::DebugPrint(DBGGeneric, L"DumpContentsTable: Outputting as MSG\n");
	if (ulOptions & cli::OPT_RETRYSTREAMPROPS)
		output::DebugPrint(DBGGeneric, L"DumpContentsTable: Will retry stream properties\n");
	if (ulOptions & cli::OPT_SKIPATTACHMENTS)
		output::DebugPrint(DBGGeneric, L"DumpContentsTable: Will skip attachments\n");
	if (ulOptions & cli::OPT_LIST) output::DebugPrint(DBGGeneric, L"DumpContentsTable: List only mode\n");
	if (ulCount) output::DebugPrint(DBGGeneric, L"DumpContentsTable: Limiting output to %u messages.\n", ulCount);

	if (lpFolder)
	{
		mapiprocessor::CDumpStore MyDumpStore;
		SSortOrderSet SortOrder = {0};
		MyDumpStore.InitMDB(lpMDB);
		MyDumpStore.InitFolder(lpFolder);
		MyDumpStore.InitFolderPathRoot(lpszDir);
		MyDumpStore.InitFolderContentsRestriction(lpRes);
		if (ulOptions & cli::OPT_MSG) MyDumpStore.EnableMSG();
		if (ulOptions & cli::OPT_LIST) MyDumpStore.EnableList();
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

		if (!(ulOptions & cli::OPT_RETRYSTREAMPROPS)) MyDumpStore.DisableStreamRetry();
		if (ulOptions & cli::OPT_SKIPATTACHMENTS) MyDumpStore.DisableEmbeddedAttachments();

		MyDumpStore.ProcessFolders(
			0 != (ulOptions & cli::OPT_DOCONTENTS), 0 != (ulOptions & cli::OPT_DOASSOCIATEDCONTENTS), false);
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
		mapiprocessor::CDumpStore MyDumpStore;
		MyDumpStore.InitMessagePath(lpszXMLFile);
		if (!bRetryStreamProps) MyDumpStore.DisableStreamRetry();
		if (!bOutputAttachments) MyDumpStore.DisableEmbeddedAttachments();

		// Just assume this message might have attachments
		MyDumpStore.ProcessMessage(lpMessage, true, nullptr);
		lpMessage->Release();
	}
}

void DoContents(_In_ cli::MYOPTIONS ProgOpts)
{
	SRestriction sResTop = {0};
	SRestriction sResMiddle[2] = {0};
	SRestriction sResSubject[2] = {0};
	SRestriction sResMessageClass[2] = {0};
	SPropValue sPropValue[2] = {0};
	LPSRestriction lpRes = nullptr;
	if (!ProgOpts.lpszSubject.empty() || !ProgOpts.lpszMessageClass.empty())
	{
		// RES_AND
		//   RES_AND (optional)
		//     RES_EXIST - PR_SUBJECT_W
		//     RES_CONTENT - lpszSubject
		//   RES_AND (optional)
		//     RES_EXIST - PR_MESSAGE_CLASS_W
		//     RES_CONTENT - lpszMessageClass
		auto i = 0;
		if (!ProgOpts.lpszSubject.empty())
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
			sPropValue[0].Value.lpszW = const_cast<LPWSTR>(ProgOpts.lpszSubject.c_str());
			i++;
		}

		if (!ProgOpts.lpszMessageClass.empty())
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
			sPropValue[1].Value.lpszW = const_cast<LPWSTR>(ProgOpts.lpszMessageClass.c_str());
			i++;
		}
		sResTop.rt = RES_AND;
		sResTop.res.resAnd.cRes = i;
		sResTop.res.resAnd.lpRes = &sResMiddle[0];
		lpRes = &sResTop;
		output::outputRestriction(DBGGeneric, nullptr, lpRes, nullptr);
	}

	DumpContentsTable(
		ProgOpts.lpszProfile.c_str(),
		ProgOpts.lpMDB,
		ProgOpts.lpFolder,
		!ProgOpts.lpszOutput.empty() ? ProgOpts.lpszOutput.c_str() : L".",
		ProgOpts.ulOptions,
		ProgOpts.ulFolder,
		ProgOpts.lpszFolderPath.c_str(),
		ProgOpts.ulCount,
		lpRes);
}

void DoMSG(_In_ cli::MYOPTIONS ProgOpts)
{
	DumpMSG(
		ProgOpts.lpszInput.c_str(),
		!ProgOpts.lpszOutput.empty() ? ProgOpts.lpszOutput.c_str() : L".",
		0 != (ProgOpts.ulOptions & cli::OPT_RETRYSTREAMPROPS),
		0 == (ProgOpts.ulOptions & cli::OPT_SKIPATTACHMENTS));
}