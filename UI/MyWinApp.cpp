// Defines the class behaviors for the application.

#include <StdAfx.h>
#include <UI/MyWinApp.h>
#include <UI/ParentWnd.h>
#include <core/mapi/cache/mapiObjects.h>
#include <core/utility/import.h>
#include <UI/Dialogs/ContentsTable/MainDlg.h>
#include <UI/mapiui.h>
#include <UI/UIFunctions.h>
#include <core/utility/registry.h>

// The one and only CMyWinApp object
ui::CMyWinApp theApp;

namespace ui
{
	CMyWinApp::CMyWinApp()
	{
		// Assume true if we don't find a reg key set to false.
		auto bTerminateOnCorruption = true;

		// Remove current working directory from DLL search path
		(void) ::SetDllDirectoryW(L"");

		HKEY hRootKey = nullptr;
		auto lStatus = RegOpenKeyExW(HKEY_CURRENT_USER, registry::RKEY_ROOT, NULL, KEY_READ, &hRootKey);
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

	BOOL CMyWinApp::PreTranslateMessage(MSG* pMsg)
	{
		if (pMsg && pMsg->message == WM_KEYDOWN && pMsg->wParam == 'D' && GetKeyState(VK_CONTROL) < 0 &&
			GetKeyState(VK_SHIFT) < 0 && GetKeyState(VK_MENU) >= 0)
		{
			if (!(pMsg->lParam & (1L << 30)))
			{
				const auto theme = GetEffectiveTheme();
				const auto dark = theme == uiTheme::Dark ||
								  (theme == uiTheme::HighContrast &&
								   registry::uiTheme.ulCurDWORD == static_cast<DWORD>(uiThemePreference::Dark));
				registry::uiTheme = static_cast<DWORD>(dark ? uiThemePreference::Light : uiThemePreference::Dark);
				registry::WriteToRegistry();
				ApplyThemeToAllWindows();
			}

			return true;
		}

		return CWinApp::PreTranslateMessage(pMsg);
	}

	// CMyWinApp initialization
	BOOL CMyWinApp::InitInstance()
	{
		mapiui::initCallbacks();

		// Create a parent window that all objects grab a pointer to, ensuring we don't
		// quit this thread until all objects have freed themselves.
		auto pWnd = GetParentWnd();
		if (pWnd)
		{
			m_pMainWnd = pWnd;
			new dialog::CMainDlg(std::make_shared<cache::CMapiObjects>(nullptr));
			pWnd->Release();
		}

		return true;
	}
} // namespace ui