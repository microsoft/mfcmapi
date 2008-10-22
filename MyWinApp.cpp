// MyWinApp.cpp : Defines the class behaviors for the application.
//

#include "stdafx.h"
#include "MyWinApp.h"
#include "ParentWnd.h"
#include "MainDlg.h"
#include "MapiObjects.h"

/////////////////////////////////////////////////////////////////////////////
// CMyWinApp

/////////////////////////////////////////////////////////////////////////////
// The one and only CMyWinApp object
CMyWinApp theApp;

/////////////////////////////////////////////////////////////////////////////
// CMyWinApp initialization

BOOL CMyWinApp::InitInstance()
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
}