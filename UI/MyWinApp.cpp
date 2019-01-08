// Defines the class behaviors for the application.

#include <StdAfx.h>
#include <UI/MyWinApp.h>
#include <UI/ParentWnd.h>
#include <MAPI/Cache/MapiObjects.h>
#include <ImportProcs.h>
#include <UI/Dialogs/ContentsTable/MainDlg.h>

// The one and only CMyWinApp object
ui::CMyWinApp theApp;

namespace ui
{
	CMyWinApp::CMyWinApp()
	{
		// Assume true if we don't find a reg key set to false.
		auto bTerminateOnCorruption = true;

		HKEY hRootKey = nullptr;
		auto lStatus = RegOpenKeyExW(HKEY_CURRENT_USER, RKEY_ROOT, NULL, KEY_READ, &hRootKey);
		if (lStatus == ERROR_SUCCESS)
		{
			DWORD dwRegVal = 0;
			DWORD dwType = REG_DWORD;
			ULONG cb = sizeof dwRegVal;
			lStatus = RegQueryValueExW(
				hRootKey,
				registry::heapEnableTerminationOnCorruption.szKeyName.c_str(),
				nullptr,
				&dwType,
				reinterpret_cast<LPBYTE>(&dwRegVal),
				&cb);
			if (lStatus == ERROR_SUCCESS && !dwRegVal)
			{
				bTerminateOnCorruption = false;
			}

			if (hRootKey) RegCloseKey(hRootKey);
		}

		if (bTerminateOnCorruption)
		{
			import::MyHeapSetInformation(nullptr, HeapEnableTerminationOnCorruption, nullptr, 0);
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
			m_pMainWnd = static_cast<CWnd*>(pWnd);
			auto MyObjects = new (std::nothrow) cache::CMapiObjects(nullptr);
			if (MyObjects)
			{
				new dialog::CMainDlg(pWnd, MyObjects);
				MyObjects->Release();
			}
			pWnd->Release();
		}

		return true;
	}
} // namespace ui