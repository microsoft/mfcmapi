// Dialog.cpp : implementation file
//

#include "stdafx.h"
#include "Dialog.h"
#include "UIFunctions.h"
#include "ImportProcs.h"
#include <propkey.h>

static wstring CLASS = L"CMyDialog";

CMyDialog::CMyDialog() :CDialog()
{
	Constructor();
}

CMyDialog::CMyDialog(UINT nIDTemplate, CWnd* pParentWnd) : CDialog(nIDTemplate, pParentWnd)
{
	Constructor();
}

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
	(void) ::GetWindowThreadProcessId(m_hWndPrevious, &pid);
	if (::GetCurrentProcessId() != pid)
	{
		m_hWndPrevious = nullptr;
	}
}

CMyDialog::~CMyDialog()
{
	TRACE_DESTRUCTOR(CLASS);
	if (m_lpNonModalParent) m_lpNonModalParent->Release();
}

BEGIN_MESSAGE_MAP(CMyDialog, CDialog)
	ON_WM_MEASUREITEM()
	ON_WM_DRAWITEM()
END_MESSAGE_MAP()

void CMyDialog::SetStatusHeight(int iHeight)
{
	m_iStatusHeight = iHeight;
}

int CMyDialog::GetStatusHeight() const
{
	return m_iStatusHeight;
}

// Performs an NC hittest using coordinates from WM_MOUSE* messages
int NCHitTestMouse(HWND hWnd, LPARAM lParam)
{
	POINT pt = { GET_X_LPARAM(lParam), GET_Y_LPARAM(lParam) };
	(void) ::MapWindowPoints(hWnd, nullptr, &pt, 1); // Map our client point to the screen
	lParam = MAKELONG(pt.x, pt.y);
	return static_cast<int>(::SendMessage(hWnd, WM_NCHITTEST, NULL, lParam));
}

bool DepressSystemButton(HWND hWnd, int iHitTest)
{
	auto bDepressed = true;
	DrawSystemButtons(hWnd, nullptr, iHitTest);
	SetCapture(hWnd);
	for (;;)
	{
		MSG msg = { nullptr };
		if (::PeekMessage(&msg, hWnd, WM_MOUSEFIRST, WM_MOUSELAST, PM_REMOVE))
		{
			switch (msg.message)
			{
			case WM_LBUTTONUP:
			{
				if (bDepressed)
					DrawSystemButtons(hWnd, nullptr, NULL);
				ReleaseCapture();
				if (NCHitTestMouse(hWnd, msg.lParam) == iHitTest) return true;
				return false;
			}

			case WM_MOUSEMOVE:
			{
				if (NCHitTestMouse(hWnd, msg.lParam) == iHitTest)
				{
					DrawSystemButtons(hWnd, nullptr, iHitTest);
				}
				else
				{
					DrawSystemButtons(hWnd, nullptr, NULL);
				}
			}
			break;
			}
		}
	}
}

#define WM_NCUAHDRAWCAPTION 0x00AE
#define WM_NCUAHDRAWFRAME 0x00AF

LRESULT CMyDialog::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	LRESULT lRes = 0;
	if (HandleControlUI(message, wParam, lParam, &lRes)) return lRes;

	switch (message)
	{
	case WM_NCUAHDRAWCAPTION:
	case WM_NCUAHDRAWFRAME:
		return 0;
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
		DrawWindowFrame(m_hWnd, !!wParam, m_iStatusHeight);
		break;
	case WM_SETTEXT:
		CDialog::WindowProc(message, wParam, lParam);
		DrawWindowFrame(m_hWnd, true, m_iStatusHeight);
		return true;
	case WM_NCPAINT:
		DrawWindowFrame(m_hWnd, true, m_iStatusHeight);
		return 0;
	case WM_CREATE:
		// Ensure all windows group together by enforcing a consistent App User Model ID.
		// We don't use SetCurrentProcessExplicitAppUserModelID because logging on to MAPI somehow breaks this.
		if (pfnSHGetPropertyStoreForWindow)
		{
			IPropertyStore* pps = nullptr;
			auto hRes = pfnSHGetPropertyStoreForWindow(m_hWnd, IID_PPV_ARGS(&pps));
			if (SUCCEEDED(hRes) && pps) {
				PROPVARIANT var = { 0 };
				var.vt = VT_LPWSTR;
				var.pwszVal = L"Microsoft.MFCMAPI";

				(void)pps->SetValue(PKEY_AppUserModel_ID, var);
			}

			if (pps) pps->Release();
		}

		if (pfnSetWindowTheme) (void)pfnSetWindowTheme(m_hWnd, L"", L"");
		{
			// These calls force Windows to initialize the system menu for this window.
			// This avoids repaints whenever the system menu is later accessed.
			// We eliminate classic mode visual artifacts with this call.
			(void) ::GetSystemMenu(m_hWnd, false);
			MENUBARINFO mbi = { 0 };
			mbi.cbSize = sizeof mbi;
			(void) ::GetMenuBarInfo(m_hWnd, OBJID_SYSMENU, 0, &mbi);
		}
		break;
	}

	return CDialog::WindowProc(message, wParam, lParam);
}

void CMyDialog::DisplayParentedDialog(CParentWnd* lpNonModalParent, UINT iAutoCenterWidth)
{
	auto hRes = S_OK;
	m_iAutoCenterWidth = iAutoCenterWidth;
	m_lpszTemplateName = MAKEINTRESOURCE(IDD_BLANK_DIALOG);

	m_lpNonModalParent = lpNonModalParent;
	if (m_lpNonModalParent) m_lpNonModalParent->AddRef();

	m_hwndCenteringWindow = GetActiveWindow();

	auto hInst = AfxFindResourceHandle(m_lpszTemplateName, RT_DIALOG);
	HRSRC hResource = nullptr;
	EC_D(hResource, ::FindResource(hInst, m_lpszTemplateName, RT_DIALOG));
	if (hResource)
	{
		HGLOBAL hTemplate = nullptr;
		EC_D(hTemplate, LoadResource(hInst, hResource));
		if (hTemplate)
		{
			auto lpDialogTemplate = static_cast<LPCDLGTEMPLATE>(LockResource(hTemplate));
			EC_B(CreateDlgIndirect(lpDialogTemplate, m_lpNonModalParent, hInst));
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
	if (m_hWndPrevious && !m_lpNonModalParent && m_hwndCenteringWindow && IDD_BLANK_DIALOG == reinterpret_cast<intptr_t>(m_lpszTemplateName))
	{
		// Cheap cascade effect
		RECT rc = { 0 };
		(void) ::GetWindowRect(m_hWndPrevious, &rc);
		LONG lOffset = GetSystemMetrics(SM_CXSMSIZE);
		(void) ::SetWindowPos(m_hWnd, nullptr, rc.left + lOffset, rc.top + lOffset, 0, 0, SWP_NOSIZE | SWP_NOZORDER);
	}
	else
	{
		CenterWindow(m_hwndCenteringWindow);
	}

	return false;
}

// Measure menu item widths
void CMyDialog::OnMeasureItem(int /*nIDCtl*/, _In_ LPMEASUREITEMSTRUCT lpMeasureItemStruct)
{
	MeasureItem(lpMeasureItemStruct);
}

// Draw menu items
void CMyDialog::OnDrawItem(int /*nIDCtl*/, _In_ LPDRAWITEMSTRUCT lpDrawItemStruct)
{
	DrawItem(lpDrawItemStruct);
}