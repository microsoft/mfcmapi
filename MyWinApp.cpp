// MyWinApp.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "MyWinApp.h"
#include "ParentWnd.h"
#include "MainDlg.h"
#include "MapiObjects.h"
#include "ImportProcs.h"

/////////////////////////////////////////////////////////////////////////////
// CMyWinApp

/////////////////////////////////////////////////////////////////////////////
// The one and only CMyWinApp object
CMyWinApp theApp;

CMyWinApp::CMyWinApp()
{
	// Assume true if we don't find a reg key set to false.
	BOOL bTerminateOnCorruption = true;

	HKEY hRootKey = NULL;
	LONG lStatus = ERROR_SUCCESS;

	lStatus = RegOpenKeyEx(
		HKEY_CURRENT_USER,
		RKEY_ROOT,
		NULL,
		KEY_READ,
		&hRootKey);
	if (ERROR_SUCCESS == lStatus && hRootKey)
	{
		DWORD dwRegVal = 0;
		DWORD dwType = REG_DWORD;
		ULONG cb = sizeof(dwRegVal);
		lStatus = RegQueryValueEx(
			hRootKey,
			RegKeys[regkeyHEAPENABLETERMINATIONONCORRUPTION].szKeyName,
			NULL,
			&dwType,
			(LPBYTE) &dwRegVal,
			&cb);
		if (ERROR_SUCCESS == lStatus && !dwRegVal)
		{
			bTerminateOnCorruption = false;
		}
	}
	if (hRootKey) RegCloseKey(hRootKey);

	if (bTerminateOnCorruption)
	{
		MyHeapSetInformation(NULL, HeapEnableTerminationOnCorruption, NULL, 0);
	}
} // CMyWinApp

/////////////////////////////////////////////////////////////////////////////
// CMyWinApp initialization

_Check_return_ BOOL CMyWinApp::InitInstance()
{
	// Create a parent window that all objects get a pointer to, ensuring we don't
	// quit this thread until all objects have freed themselves.
	CParentWnd *pWnd = new CParentWnd();
	if (pWnd)
	{
		m_pMainWnd = (CWnd *) pWnd;
		CMapiObjects* MyObjects = new CMapiObjects(NULL);
		if (MyObjects)
		{
			new CMainDlg(pWnd,MyObjects);
			MyObjects->Release();
		}
		pWnd->Release();
	}

	return TRUE;
} // CMyWinApp::InitInstance