// MyWinApp.cpp : Defines the class behaviors for the application.

#include "stdafx.h"
#include "MyWinApp.h"
#include "ParentWnd.h"
#include "MapiObjects.h"
#include "ImportProcs.h"
#include "Dialogs/ContentsTable/MainDlg.h"

// The one and only CMyWinApp object
CMyWinApp theApp;

CMyWinApp::CMyWinApp()
{
	// Assume true if we don't find a reg key set to false.
	auto bTerminateOnCorruption = true;

	HKEY hRootKey = nullptr;
	auto lStatus = RegOpenKeyExW(
		HKEY_CURRENT_USER,
		RKEY_ROOT,
		NULL,
		KEY_READ,
		&hRootKey);
	if (ERROR_SUCCESS == lStatus)
	{
		DWORD dwRegVal = 0;
		DWORD dwType = REG_DWORD;
		ULONG cb = sizeof dwRegVal;
		lStatus = RegQueryValueExW(
			hRootKey,
			RegKeys[regkeyHEAPENABLETERMINATIONONCORRUPTION].szKeyName.c_str(),
			nullptr,
			&dwType,
			reinterpret_cast<LPBYTE>(&dwRegVal),
			&cb);
		if (ERROR_SUCCESS == lStatus && !dwRegVal)
		{
			bTerminateOnCorruption = false;
		}

		if (hRootKey) RegCloseKey(hRootKey);
	}

	if (bTerminateOnCorruption)
	{
		MyHeapSetInformation(nullptr, HeapEnableTerminationOnCorruption, nullptr, 0);
	}
}

// CMyWinApp initialization
BOOL CMyWinApp::InitInstance()
{
	// Create a parent window that all objects get a pointer to, ensuring we don't
	// quit this thread until all objects have freed themselves.
	auto pWnd = new CParentWnd();
	if (pWnd)
	{
		m_pMainWnd = static_cast<CWnd *>(pWnd);
		auto MyObjects = new CMapiObjects(nullptr);
		if (MyObjects)
		{
			new CMainDlg(pWnd, MyObjects);
			MyObjects->Release();
		}
		pWnd->Release();
	}

	return true;
}