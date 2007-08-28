// MyWinApp.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "Error.h"

#include "MyWinApp.h"

#include "ParentWnd.h"
#include "MainDlg.h"

#include "MapiObjects.h"
#include "FolderDlg.h"

#include "MAPIFunctions.h"
#include "MAPIStoreFunctions.h"
#include "MFCUtilityFunctions.h"
#include "ColumnTags.h"

//Only include for test
#include "AbContDlg.h"
#include "AbDlg.h"
#include "AttachmentsDlg.h"
#include "FolderDlg.h"
#include "MailboxTableDlg.h"
#include "MsgServiceTableDlg.h"
#include "MsgStoreDlg.h"
#include "ProfileListDlg.h"
#include "ProviderTableDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CMyWinApp

/////////////////////////////////////////////////////////////////////////////
// The one and only CMyWinApp object

CMyWinApp theApp;

/*
void TestApp(CParentWnd* pWnd)
{
	HRESULT hRes = S_OK;
	LPMAPISESSION lpMAPISession = NULL;
	LPMDB lpMDB = NULL;

	//must be first - initializes MAPI
	CMapiObjects* MyObjects = new CMapiObjects(NULL);

	EC_H_CANCEL(MAPILogonEx(
		NULL,
		NULL,
		NULL,
		MAPI_EXTENDED | MAPI_LOGON_UI | MAPI_NEW_SESSION,
		&lpMAPISession));

	EC_H(OpenDefaultMessageStore(
								lpMAPISession,
								&lpMDB));

	LPMAPIFOLDER lpInbox = NULL;
//	ULONG ulObjectType = NULL;

	EC_H(GetInbox(lpMDB,&lpInbox));


	MyObjects->SetMDB(lpMDB);
	MyObjects->SetSession(lpMAPISession);

	new CMsgStoreDlg(
		pWnd,
		MyObjects,
		lpInbox,
		mfcmapiDO_NOT_SHOW_DELETED_ITEMS);
//	new CFolderDlg(pWnd,MyObjects,lpInbox,mfcmapiSHOW_NORMAL_CONTENTS,mfcmapiDO_NOT_SHOW_DELETED_ITEMS);

	if (lpInbox) lpInbox->Release();
	if (lpMDB) lpMDB->Release();
	if (lpMAPISession) lpMAPISession->Release();

	//must be last
	if (MyObjects) MyObjects->Release();
}*/


//save this for a later version
/*
//either match the string directly
//or remove the first character of the arg and then match it
//this will allow "i", "/i", "\i", "-i", "I", "/I", "\I", "-I" all to match "i" // STRING_OK
//DUMB BUG: This will also match "hi", "bi", "fi", etc...need to code up a 'correct' algorithm // STRING_OK
BOOL bCheckParam(LPTSTR szArg,LPTSTR szStringToMatch)
{
	HRESULT hRes = S_OK;
	size_t ulArgLen = 0;
	size_t ulMatchLen = 0;

	EC_H(StringCchLength(szArg,STRSAFE_MAX_CCH,&ulArgLen));
	EC_H(StringCchLength(szStringToMatch,STRSAFE_MAX_CCH,&ulMatchLen));

	size_t ulLen = ulArgLen;
	if (ulLen > ulMatchLen) ulLen = ulMatchLen;
	if (CSTR_EQUAL == CompareString(g_lcid, NORM_IGNORECASE, szArg, ulLen, szStringToMatch, ulLen))
	{
		return true;
	}
	if (ulArgLen >=2)
	{
		ulLen = ulArgLen-1;
		if (ulLen > ulMatchLen) ulLen = ulMatchLen;
		if (CSTR_EQUAL == CompareString(g_lcid, NORM_IGNORECASE, szArg+1, ulLen, szStringToMatch, ulLen))
		{
			return true;
		}
	}
	return false;
}*/

/////////////////////////////////////////////////////////////////////////////
// CMyWinApp initialization

BOOL CMyWinApp::InitInstance()
{
	//Create a parent window that all objects get a pointer to, ensuring we don't quit
	//this thread until all objects have freed themselves.
	//BTW, debug output will not work until this loads - perhaps I should move all that into this function?
	CParentWnd *pWnd = new CParentWnd();
	m_pMainWnd = (CWnd *) pWnd;

	CMapiObjects* MyObjects = new CMapiObjects(NULL);

//save this for a later version
/*	//parse out the command line
	LPTSTR*	argv = __targv;
	int		argc = __argc;

	for (int i = 0;i<argc;i++)
	{
		DebugPrint(DBGGeneric,_T("%s\n"),argv[i]);
	}*/

	if (pWnd) new CMainDlg(pWnd,MyObjects);

//	TestApp(pWnd);

	if (MyObjects) MyObjects->Release();

	//Test calls to CBaseDialog derived objects
//	new CAbContDlg(pWnd,NULL);
//	new CAbDlg(pWnd,NULL,NULL);
//	new CAttachmentsDlg(pWnd,NULL,NULL,NULL);
//	new CFolderDlg(pWnd,NULL,NULL,NULL,mfcmapiSHOW_NORMAL_CONTENTS,mfcmapiSHOW_DELETED_ITEMS);
//	new CMsgServiceTableDlg(pWnd,NULL,NULL,NULL);
//	new CMsgStoreDlg(pWnd,NULL,NULL,mfcmapiDO_NOT_SHOW_DELETED_ITEMS);
//	new CProfileListDlg(pWnd,NULL,NULL);
//	new CProviderTableDlg(pWnd,NULL,NULL,NULL);

//	CBaseDialog *MyComDlg = new CBaseDialog(pWnd,NULL,NULL);
//	MyComDlg->CreateDialogAndMenu(NULL);
/*
	CContentsTableDlg *MyContents = new CContentsTableDlg(
		pWnd,
		NULL,
		NULL,
		mfcmapiCALL_CREATE_DIALOG,
		NULL,
		NULL,
		NULL,
		NULL);


	CHierarchyTableDlg *MyHierarchy = new CHierarchyTableDlg(
		pWnd,
		NULL,
		NULL,
		NULL);
	MyHierarchy->CreateDialogAndMenu(NULL);
//*/

	if (pWnd) pWnd->Release();

	return TRUE;
}

