#include <StdAfx.h>
#include <UI/Dialogs/Dialog.h>
#include <UI/UIFunctions.h>
#include <ImportProcs.h>
#include <propkey.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>

namespace dialog
{
	static std::wstring CLASS = L"CMyDialog";

	CMyDialog::CMyDialog() { Constructor(); }

	CMyDialog::CMyDialog(const UINT nIDTemplate, CWnd* pParentWnd) : CDialog(nIDTemplate, pParentWnd) { Constructor(); }

	void CMyDialog::Constructor()
	{
		TRACE_CONSTRUCTOR(CLASS);
		m_lpNonModalParent = nullptr;
		m_hwndCenteringWindow = nullptr;
		m_iAutoCenterWidth = NULL;
		m_iStatusHeight = 0;

		// If the previous foreground window is ours, remember its handle for computing cascades
		m_hWndPrevious = ::GetForegroundWindow();
		DWORD pid = NULL;
		(void) GetWindowThreadProcessId(m_hWndPrevious, &pid);
		if (GetCurrentProcessId() != pid)
		{
			m_hWndPrevious = nullptr;
		}
	}

	CMyDialog::~CMyDialog()
	{
		TRACE_DESTRUCTOR(CLASS);
		if (m_lpNonModalParent) m_lpNonModalParent->Release();
	}

	void CMyDialog::SetStatusHeight(const int iHeight) { m_iStatusHeight = iHeight; }

	int CMyDialog::GetStatusHeight() const { return m_iStatusHeight; }

	std::wstring FormatHT(const LRESULT ht)
	{
		std::wstring szRet;
		switch (ht)
		{
		case HTNOWHERE:
			szRet = L"HTNOWHERE";
			break;
		case HTCLIENT:
			szRet = L"HTCLIENT";
			break;
		case HTCAPTION:
			szRet = L"HTCAPTION";
			break;
		case HTCLOSE:
			szRet = L"HTCLOSE";
			break;
		case HTMAXBUTTON:
			szRet = L"HTMAXBUTTON";
			break;
		case HTMINBUTTON:
			szRet = L"HTMINBUTTON";
			break;
		}

		return strings::format(L"ht = 0x%X = %ws", ht, szRet.c_str());
	}

	// Point should be screen relative coordinates
	// Upper left of screen is (0,0)
	// If this point is inside our caption rects, return the HT matching the hit
	// Else return HTNOWHERE
	LRESULT CheckButtons(HWND hWnd, POINT pt)
	{
		auto ret = HTNOWHERE;
		output::DebugPrint(DBGUI, L"CheckButtons: pt = %d %d", pt.x, pt.y);

		// Get the screen coordinates of our window
		auto rcWindow = RECT{};
		GetWindowRect(hWnd, &rcWindow);

		// We subtract to get coordinates relative to our window
		// GetCaptionRects coordinates are now compatible
		pt.x -= rcWindow.left;
		pt.y -= rcWindow.top;
		output::Outputf(DBGUI, nullptr, false, L" mapped = %d %d\r\n", pt.x, pt.y);

		auto rcCloseIcon = RECT{};
		auto rcMaxIcon = RECT{};
		auto rcMinIcon = RECT{};
		ui::GetCaptionRects(hWnd, nullptr, nullptr, &rcCloseIcon, &rcMaxIcon, &rcMinIcon, nullptr);
		output::DebugPrint(
			DBGUI, L"rcMinIcon: %d %d %d %d\n", rcMinIcon.left, rcMinIcon.top, rcMinIcon.right, rcMinIcon.bottom);
		output::DebugPrint(
			DBGUI, L"rcMaxIcon: %d %d %d %d\n", rcMaxIcon.left, rcMaxIcon.top, rcMaxIcon.right, rcMaxIcon.bottom);
		output::DebugPrint(
			DBGUI,
			L"rcCloseIcon: %d %d %d %d\n",
			rcCloseIcon.left,
			rcCloseIcon.top,
			rcCloseIcon.right,
			rcCloseIcon.bottom);
		if (PtInRect(&rcCloseIcon, pt)) ret = HTCLOSE;
		if (PtInRect(&rcMaxIcon, pt)) ret = HTMAXBUTTON;
		if (PtInRect(&rcMinIcon, pt)) ret = HTMINBUTTON;

		output::DebugPrint(DBGUI, L"CheckButtons result: %ws\r\n", FormatHT(ret).c_str());

		return ret;
	}

	// Handles WM_NCHITTEST, substituting our custom button locations for the default ones
	// Everything else stays the same
	LRESULT CMyDialog::NCHitTest(const WPARAM wParam, LPARAM const lParam)
	{
		// These are screen coordinates of the mouse pointer
		const POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		output::DebugPrint(DBGUI, L"WM_NCHITTEST: pt = %d %d\r\n", pt.x, pt.y);

		auto ht = CDialog::WindowProc(WM_NCHITTEST, wParam, lParam);
		if (ht == HTCAPTION || ht == HTCLOSE || ht == HTMAXBUTTON || ht == HTMINBUTTON)
		{
			ht = CheckButtons(m_hWnd, pt);

			// If we weren't on a button, but CDialog::WindowProc thought we were, we must be on the caption
			if (ht == HTNOWHERE) ht = HTCAPTION;
		}

		ui::DrawSystemButtons(m_hWnd, nullptr, ht, true);
		output::DebugPrint(DBGUI, L"%ws\r\n", FormatHT(ht).c_str());
		return ht;
	}

	// Performs an non-client hittest using coordinates from WM_MOUSE* messages
	LRESULT NCHitTestMouse(HWND hWnd, const LPARAM lParam)
	{
		// These are client coordinates - we need to translate them to screen coordinates
		POINT pt = {GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam)};
		output::DebugPrint(DBGUI, L"NCHitTestMouse: pt = %d %d", pt.x, pt.y);
		(void) MapWindowPoints(hWnd, nullptr, &pt, 1); // Map our client point to the screen
		output::Outputf(DBGUI, nullptr, false, L" mapped = %d %d\r\n", pt.x, pt.y);

		const auto ht = CheckButtons(hWnd, pt);
		output::DebugPrint(DBGUI, L"%ws\r\n", FormatHT(ht).c_str());
		return ht;
	}

	bool DepressSystemButton(HWND hWnd, const int iHitTest)
	{
		ui::DrawSystemButtons(hWnd, nullptr, iHitTest, false);
		SetCapture(hWnd);
		for (;;)
		{
			auto msg = MSG{};
			if (::PeekMessage(&msg, hWnd, WM_MOUSEFIRST, WM_MOUSELAST, PM_REMOVE))
			{
				switch (msg.message)
				{
				case WM_LBUTTONUP:
					output::DebugPrint(DBGUI, L"WM_LBUTTONUP\n");
					ui::DrawSystemButtons(hWnd, nullptr, HTNOWHERE, false);
					ReleaseCapture();
					return NCHitTestMouse(hWnd, msg.lParam) == iHitTest;

				case WM_MOUSEMOVE:
					output::DebugPrint(DBGUI, L"WM_MOUSEMOVE\n");
					ui::DrawSystemButtons(hWnd, nullptr, iHitTest, NCHitTestMouse(hWnd, msg.lParam) != iHitTest);

					break;
				}
			}
		}
	}

#define WM_NCUAHDRAWCAPTION 0x00AE
#define WM_NCUAHDRAWFRAME 0x00AF

	LRESULT CMyDialog::WindowProc(const UINT message, const WPARAM wParam, LPARAM lParam)
	{
		LRESULT lRes = 0;
		if (ui::HandleControlUI(message, wParam, lParam, &lRes)) return lRes;

		switch (message)
		{
		case WM_NCUAHDRAWCAPTION:
		case WM_NCUAHDRAWFRAME:
			return 0;
		case WM_NCHITTEST:
			return NCHitTest(wParam, lParam);
		case WM_NCLBUTTONDOWN:
			switch (wParam)
			{
			case HTCLOSE:
			case HTMAXBUTTON:
			case HTMINBUTTON:
				if (DepressSystemButton(m_hWnd, static_cast<int>(wParam)))
				{
					switch (wParam)
					{
					case HTCLOSE:
						::SendMessageA(m_hWnd, WM_SYSCOMMAND, SC_CLOSE, NULL);
						break;
					case HTMAXBUTTON:
						::SendMessageA(m_hWnd, WM_SYSCOMMAND, ::IsZoomed(m_hWnd) ? SC_RESTORE : SC_MAXIMIZE, NULL);
						break;
					case HTMINBUTTON:
						::SendMessageA(m_hWnd, WM_SYSCOMMAND, SC_MINIMIZE, NULL);
						break;
					}
				}

				return 0;
			}

			break;
		case WM_NCACTIVATE:
			// Pass -1 to DefWindowProc to signal we do not want our client repainted.
			lParam = -1;
			ui::DrawWindowFrame(m_hWnd, !!wParam, m_iStatusHeight);
			break;
		case WM_SETTEXT:
			CDialog::WindowProc(message, wParam, lParam);
			ui::DrawWindowFrame(m_hWnd, true, m_iStatusHeight);
			return true;
		case WM_NCPAINT:
			ui::DrawWindowFrame(m_hWnd, true, m_iStatusHeight);
			return 0;
		case WM_CREATE:
			// Ensure all windows group together by enforcing a consistent App User Model ID.
			// We don't use SetCurrentProcessExplicitAppUserModelID because logging on to MAPI somehow breaks this.
			if (import::pfnSHGetPropertyStoreForWindow)
			{
				IPropertyStore* pps = nullptr;
				const auto hRes = import::pfnSHGetPropertyStoreForWindow(m_hWnd, IID_PPV_ARGS(&pps));
				if (SUCCEEDED(hRes) && pps)
				{
					auto var = PROPVARIANT{};
					var.vt = VT_LPWSTR;
					var.pwszVal = const_cast<LPWSTR>(L"Microsoft.MFCMAPI");

					(void) pps->SetValue(PKEY_AppUserModel_ID, var);
				}

				if (pps) pps->Release();
			}

			if (import::pfnSetWindowTheme) (void) import::pfnSetWindowTheme(m_hWnd, L"", L"");
			{
				// These calls force Windows to initialize the system menu for this window.
				// This avoids repaints whenever the system menu is later accessed.
				// We eliminate classic mode visual artifacts with this call.
				(void) ::GetSystemMenu(m_hWnd, false);
				auto mbi = MENUBARINFO{};
				mbi.cbSize = sizeof mbi;
				(void) GetMenuBarInfo(m_hWnd, OBJID_SYSMENU, 0, &mbi);
			}
			break;
		}

		return CDialog::WindowProc(message, wParam, lParam);
	}

	void CMyDialog::SetTitle(_In_ const std::wstring& szTitle) const
	{
		// Set the title bar directly using DefWindowProcW to avoid converting Unicode
		DefWindowProcW(m_hWnd, WM_SETTEXT, 0, reinterpret_cast<LPARAM>(szTitle.c_str()));
		ui::DrawWindowFrame(m_hWnd, true, GetStatusHeight());
	}

	void CMyDialog::DisplayParentedDialog(ui::CParentWnd* lpNonModalParent, const UINT iAutoCenterWidth)
	{
		m_iAutoCenterWidth = iAutoCenterWidth;
		m_lpszTemplateName = MAKEINTRESOURCE(IDD_BLANK_DIALOG);

		m_lpNonModalParent = lpNonModalParent;
		if (m_lpNonModalParent) m_lpNonModalParent->AddRef();

		m_hwndCenteringWindow = GetActiveWindow();

		const auto hInst = AfxFindResourceHandle(m_lpszTemplateName, RT_DIALOG);
		const auto hResource = EC_D(HRSRC, ::FindResource(hInst, m_lpszTemplateName, RT_DIALOG));
		if (hResource)
		{
			const auto hTemplate = EC_D(HGLOBAL, LoadResource(hInst, hResource));
			if (hTemplate)
			{
				const auto lpDialogTemplate = static_cast<LPCDLGTEMPLATE>(LockResource(hTemplate));
				EC_B_S(CreateDlgIndirect(lpDialogTemplate, m_lpNonModalParent, hInst));
			}
		}
	}

	// MFC will call this function to check if it ought to center the dialog
	// We'll tell it no, but also place the dialog where we want it.
	BOOL CMyDialog::CheckAutoCenter()
	{
		// Make the editor wider - OnSize will fix the height for us
		if (m_iAutoCenterWidth)
		{
			SetWindowPos(nullptr, 0, 0, m_iAutoCenterWidth, 0, NULL);
		}

		// This effect only applies when opening non-CEditor IDD_BLANK_DIALOG windows
		if (m_hWndPrevious && !m_lpNonModalParent && m_hwndCenteringWindow &&
			IDD_BLANK_DIALOG == reinterpret_cast<intptr_t>(m_lpszTemplateName))
		{
			// Cheap cascade effect
			auto rc = RECT{};
			(void) ::GetWindowRect(m_hWndPrevious, &rc);
			const LONG lOffset = GetSystemMetrics(SM_CXSMSIZE);
			(void) ::SetWindowPos(
				m_hWnd, nullptr, rc.left + lOffset, rc.top + lOffset, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
		}
		else
		{
			CenterWindow(m_hwndCenteringWindow);
		}

		return false;
	}
} // namespace dialog