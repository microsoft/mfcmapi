#include "stdafx.h"
#include "MrMAPI.h"
#include "MMAcls.h"
#include "MMContents.h"
#include "MAPIStoreFunctions.h"
#include "DumpStore.h"
#include "File.h"

void DumpContentsTable(_In_z_ LPWSTR lpszProfile, _In_z_ LPWSTR lpszDir, _In_ bool bContents, _In_ bool bAssociated, _In_ bool bRetryStreamProps, _In_ ULONG ulFolder)
{
	InitMFC();
	HRESULT hRes = S_OK;
	LPMAPISESSION lpMAPISession = NULL;
	LPMDB lpMDB = NULL;
	LPMAPIFOLDER lpFolder = NULL;
	LPEXCHANGEMODIFYTABLE lpExchTbl = NULL;
	LPMAPITABLE lpTbl = NULL;

	WC_H(MAPIInitialize(NULL));

	WC_H(MrMAPILogonEx(lpszProfile,&lpMAPISession));
	if (lpMAPISession)
	{
		WC_H(OpenMessageStoreGUID(lpMAPISession, pbExchangeProviderPrimaryUserGuid, &lpMDB));
	}
	if (lpMDB)
	{
		WC_H(OpenDefaultFolder(ulFolder,lpMDB,&lpFolder));
	}
	if (lpFolder)
	{
		CDumpStore MyDumpStore;
		MyDumpStore.InitMDB(lpMDB);
		MyDumpStore.InitFolder(lpFolder);
		MyDumpStore.InitFolderPathRoot(lpszDir);
		if (bRetryStreamProps) MyDumpStore.EnableStreamRetry();
		MyDumpStore.ProcessFolders(
			bContents,
			bAssociated,
			false);
	}

	if (lpTbl) lpTbl->Release();
	if (lpExchTbl) lpExchTbl->Release();
	if (lpFolder) lpFolder->Release();
	if (lpMDB) lpMDB->Release();
	if (lpMAPISession) lpMAPISession->Release();
	MAPIUninitialize();
} // DumpContentsTable

void DumpMSG(_In_z_ LPCWSTR lpszMSGFile, _In_z_ LPCWSTR lpszXMLFile, _In_ bool bRetryStreamProps)
{
	InitMFC();
	HRESULT hRes = S_OK;
	LPMESSAGE lpMessage = NULL;

	WC_H(MAPIInitialize(NULL));

	WC_H(LoadMSGToMessage(lpszMSGFile, &lpMessage));

	if (lpMessage)
	{
		CDumpStore MyDumpStore;
		MyDumpStore.InitMessagePath(lpszXMLFile);
		if (bRetryStreamProps) MyDumpStore.EnableStreamRetry();

		// Just assume this message might have attachments
		MyDumpStore.ProcessMessage(lpMessage,true,NULL);
		lpMessage->Release();
	}

	MAPIUninitialize();
} // DumpMSG

void DoContents(_In_ MYOPTIONS ProgOpts)
{
	DumpContentsTable(
		ProgOpts.lpszProfile,
		ProgOpts.lpszOutput?ProgOpts.lpszOutput:L".",
		ProgOpts.bDoContents,
		ProgOpts.bDoAssociatedContents,
		ProgOpts.bRetryStreamProps,
		ProgOpts.ulFolder);
} // DoContents

void DoMSG(_In_ MYOPTIONS ProgOpts)
{
	DumpMSG(
		ProgOpts.lpszInput,
		ProgOpts.lpszOutput?ProgOpts.lpszOutput:L".",
		ProgOpts.bRetryStreamProps);
} // DoMAPIMIME