// FakeSplitter.cpp : implementation file
//

#include "stdafx.h"
#include "FakeSplitter.h"
#include "BaseDialog.h"
#include "UIFunctions.h"

/////////////////////////////////////////////////////////////////////////////
// CFakeSplitter

static TCHAR* CLASS = _T("CFakeSplitter");

enum FakesSplitHitTestValue
{
	noHit                   = 0,
	SplitterHit             = 1
};

CFakeSplitter::CFakeSplitter(
							 _In_ CBaseDialog *lpHostDlg
							 ):CWnd()
{
	TRACE_CONSTRUCTOR(CLASS);
	HRESULT hRes = S_OK;
	CRect pRect;

	m_bTracking = false;

	m_lpHostDlg = lpHostDlg;
	m_lpHostDlg->AddRef();

	m_lpHostDlg->GetClientRect(pRect);

	m_flSplitPercent = 0.5;

	m_iSplitWidth = 7;

	m_PaneOne = NULL;
	m_PaneTwo = NULL;

	m_SplitType = SplitHorizontal; // this doesn't mean anything yet
	m_iSplitPos = 1;

	WNDCLASSEX wc = {0};
	HINSTANCE hInst = AfxGetInstanceHandle();
	if (!(::GetClassInfoEx(hInst, _T("FakeSplitter"), &wc))) // STRING_OK
	{
		wc.cbSize  = sizeof(wc);
		wc.style   = 0; // not passing CS_VREDRAW | CS_HREDRAW fixes flicker
		wc.lpszClassName = _T("FakeSplitter"); // STRING_OK
		wc.lpfnWndProc = ::DefWindowProc;
		wc.hbrBackground = GetSysBrush(cBitmapTransBack); // helps spot flashing

		RegisterClassEx(&wc);
	}

	EC_B(Create(
		_T("FakeSplitter"), // STRING_OK
		_T("FakeSplitter"), // STRING_OK
		WS_CHILD
		| WS_CLIPSIBLINGS
		| WS_CLIPCHILDREN // required to reduce flicker
		| WS_VISIBLE,
		pRect,
		lpHostDlg,
		IDC_FAKE_SPLITTER));

	// Necessary for TAB to work. Without this, all TABS get stuck on the fake splitter control
	// instead of passing to the children. Haven't tested with nested splitters.
	EC_B(ModifyStyleEx(0,WS_EX_CONTROLPARENT));

	// Load split cursors
	EC_D(m_hSplitCursorV,::LoadCursor(GetModuleHandle(NULL), MAKEINTRESOURCE(IDC_SPLITV)));
	EC_D(m_hSplitCursorH,::LoadCursor(GetModuleHandle(NULL), MAKEINTRESOURCE(IDC_SPLITH)));
} // CFakeSplitter::CFakeSplitter

CFakeSplitter::~CFakeSplitter()
{
	TRACE_DESTRUCTOR(CLASS);
	(void)::DestroyCursor(m_hSplitCursorH);
	(void)::DestroyCursor(m_hSplitCursorV);
	DestroyWindow();
	if (m_lpHostDlg) m_lpHostDlg->Release();
} // CFakeSplitter::~CFakeSplitter

BEGIN_MESSAGE_MAP(CFakeSplitter, CWnd)
	ON_WM_MOUSEMOVE()
	ON_WM_SIZE()
	ON_WM_PAINT()
END_MESSAGE_MAP()

_Check_return_ LRESULT CFakeSplitter::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
	case WM_HELP:
		return true;
		break;
	case WM_ERASEBKGND:
		{
			return true;
			break;
		}
	case WM_LBUTTONUP:
	case WM_CANCELMODE: // Called if focus changes while we're adjusting the splitter
		{
			StopTracking();
			return NULL;
			break;
		}
	case WM_LBUTTONDOWN:
		{
			if (m_bTracking) return true;
			StartTracking(HitTest(LOWORD(lParam),HIWORD(lParam)));
			return NULL;
			break;
		}
	case WM_SETCURSOR:
		{
			if (LOWORD(lParam) == HTCLIENT &&
				(HWND) wParam == this->m_hWnd &&
				!m_bTracking) return true; // we will handle it in the mouse move
			break;
		}

	} // end switch
	return CWnd::WindowProc(message,wParam,lParam);
} // CFakeSplitter::WindowProc

void CFakeSplitter::SetPaneOne(CWnd* PaneOne)
{
	m_PaneOne = PaneOne;
} // CFakeSplitter::SetPaneOne

void CFakeSplitter::SetPaneTwo(CWnd* PaneTwo)
{
	m_PaneTwo = PaneTwo;
} // CFakeSplitter::SetPaneTwo

void CFakeSplitter::OnSize(UINT /*nType*/, int cx, int cy)
{
	HRESULT hRes = S_OK;
	HDWP hdwp = NULL;

	CalcSplitPos();

	WC_D(hdwp, BeginDeferWindowPos(2));

	if (hdwp)
	{
		if (m_PaneOne && m_PaneOne->m_hWnd)
		{
			CRect r1;
			if (SplitHorizontal == m_SplitType)
			{
				r1.SetRect(0,0,m_iSplitPos,cy);
			}
			else
			{
				r1.SetRect(0,0,cx,m_iSplitPos);
			}
			DeferWindowPos(hdwp,m_PaneOne->m_hWnd,0,0,0,r1.Width(),r1.Height(),SWP_NOZORDER);
		}
		if (m_PaneTwo && m_PaneTwo->m_hWnd)
		{
			CRect r2;
			if (SplitHorizontal == m_SplitType)
			{
				r2.SetRect(
					m_iSplitPos+m_iSplitWidth, // new x
					0, // new y
					cx, // bottom right corner
					cy); // bottom right corner
			}
			else
			{
				r2.SetRect(
					0, // new x
					m_iSplitPos+m_iSplitWidth, // new y
					cx, // bottom right corner
					cy); // bottom right corner
			}
			DeferWindowPos(hdwp,m_PaneTwo->m_hWnd,0,r2.left,r2.top,r2.Width(),r2.Height(),SWP_NOZORDER);
		}
		EC_B(EndDeferWindowPos(hdwp));
	}

	// Invalidate our splitter region to force a redraw
	if (SplitHorizontal == m_SplitType)
	{
		InvalidateRect(CRect(m_iSplitPos,0,m_iSplitPos+m_iSplitWidth,cy),false);
	}
	else
	{
		InvalidateRect(CRect(0,m_iSplitPos,cx,m_iSplitPos+m_iSplitWidth),false);
	}
} // CFakeSplitter::OnSize

void CFakeSplitter::CalcSplitPos()
{
	int iCurSpan;
	CRect rect;
	GetClientRect(rect);
	if (SplitHorizontal == m_SplitType)
	{
		iCurSpan = rect.Width();
	}
	else
	{
		iCurSpan = rect.Height();
	}

	m_iSplitPos = (int) ((FLOAT) iCurSpan * m_flSplitPercent);
	if (m_iSplitPos + m_iSplitWidth + 1 >= iCurSpan)
	{
		m_iSplitPos = m_iSplitPos - m_iSplitWidth - 1;
	}
} // CFakeSplitter::CalcSplitPos

void CFakeSplitter::SetPercent(FLOAT iNewPercent)
{
	if (iNewPercent < 0.0 || iNewPercent > 1.0)
		return;
	m_flSplitPercent = iNewPercent;

	CalcSplitPos();

	CRect rect;
	GetClientRect(rect);

	// Recalculate our layout
	OnSize(0,rect.Width(),rect.Height());
} // CFakeSplitter::SetPercent

void CFakeSplitter::SetSplitType(SplitType stSplitType)
{
	m_SplitType = stSplitType;
} // CFakeSplitter::SetSplitType

_Check_return_ int CFakeSplitter::HitTest(LONG x, LONG y)
{
	LONG lTestPos;

	if (SplitHorizontal == m_SplitType)
	{
		lTestPos = x;
	}
	else
	{
		lTestPos = y;
	}

	if (lTestPos >= m_iSplitPos &&
		lTestPos <= m_iSplitPos + m_iSplitWidth)
		return SplitterHit;
	else
		return noHit;
} // CFakeSplitter::HitTest

void CFakeSplitter::OnMouseMove(UINT /*nFlags*/, CPoint point)
{
	HRESULT hRes = S_OK;
	// If we don't have GetCapture, then we don't want to track right now.
	if (GetCapture() != this)
		StopTracking();

	if (SplitterHit == HitTest(point.x, point.y))
	{
		// This looks backwards, but it is not. A horizontal split needs the vertical cursor
		::SetCursor(SplitHorizontal == m_SplitType?m_hSplitCursorV:m_hSplitCursorH);
	}

	if (m_bTracking)
	{
		CRect Rect;
		FLOAT flNewPercent;
		GetWindowRect(Rect);

		if (SplitHorizontal == m_SplitType)
		{
			flNewPercent = point.x / (FLOAT) Rect.Width();
		}
		else
		{
			flNewPercent = point.y / (FLOAT) Rect.Height();
		}
		SetPercent(flNewPercent);

		// Force child windows to refresh now
		EC_B(RedrawWindow(NULL, NULL, RDW_ALLCHILDREN | RDW_UPDATENOW));
	}
} // CFakeSplitter::OnMouseMove

void CFakeSplitter::StartTracking(int ht)
{
	HRESULT hRes = S_OK;
	if (ht == noHit)
		return;

	// steal focus and capture
	SetCapture();
	SetFocus();

	// make sure no updates are pending
	// CSplitterWnd does this...not sure why
	EC_B(RedrawWindow(NULL, NULL, RDW_ALLCHILDREN | RDW_UPDATENOW));

	// set tracking state and appropriate cursor
	m_bTracking = true;

	// Force redraw to get our tracking gripper
	SetPercent(m_flSplitPercent);
} // CFakeSplitter::StartTracking

void CFakeSplitter::StopTracking()
{
	if (!m_bTracking)
		return;

	ReleaseCapture();

	m_bTracking = false;

	// Force redraw to get our non-tracking gripper
	SetPercent(m_flSplitPercent);
} // CFakeSplitter::StopTracking

void CFakeSplitter::OnPaint()
{
	PAINTSTRUCT ps = {0};
	::BeginPaint(m_hWnd, &ps);
	if (ps.hdc)
	{
		RECT rcWin = {0};
		::GetClientRect(m_hWnd, &rcWin);
		CDoubleBuffer db;
		HDC hdc = ps.hdc;
		db.Begin(hdc, &rcWin);

		RECT rcSplitter = rcWin;
		::FillRect(hdc, &rcSplitter, GetSysBrush(cBackground));

		POINT pts[2]; // 0 is left top, 1 is right bottom
		if (SplitHorizontal == m_SplitType)
		{
			pts[0].x = m_iSplitPos + m_iSplitWidth / 2;
			pts[0].y = rcSplitter.top;
			pts[1].x = pts[0].x;
			pts[1].y = rcSplitter.bottom;
		}
		else
		{
			pts[0].x = rcSplitter.left;
			pts[0].y = m_iSplitPos + m_iSplitWidth / 2;
			pts[1].x = rcSplitter.right;
			pts[1].y = pts[0].y;
		}

		// Draw the splitter bar
		HGDIOBJ hpenOld = ::SelectObject(hdc, GetPen(m_bTracking?cSolidPen:cDashedPen));
		::MoveToEx(hdc, pts[0].x, pts[0].y, NULL);
		::LineTo(hdc, pts[1].x, pts[1].y);
		(void) ::SelectObject(hdc, hpenOld);

		db.End(hdc);
	}

	::EndPaint(m_hWnd, &ps);
} // CFakeSplitter::OnPaint