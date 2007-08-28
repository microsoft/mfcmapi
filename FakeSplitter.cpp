// FakeSplitter.cpp : implementation file
//

#include "stdafx.h"
#include "Error.h"

#include "FakeSplitter.h"

#include "BaseDialog.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

/////////////////////////////////////////////////////////////////////////////
// CFakeSplitter

static TCHAR* CLASS = _T("CFakeSplitter");

enum FakesSplitHitTestValue
{
	noHit                   = 0,
	SplitterHit             = 1
};


CFakeSplitter::CFakeSplitter(
							 CBaseDialog *lpHostDlg
							 ):CWnd()
{
	TRACE_CONSTRUCTOR(CLASS);
	HRESULT hRes = S_OK;
	CRect pRect;

	m_bTracking = FALSE;

	m_lpHostDlg = lpHostDlg;
	m_lpHostDlg->AddRef();

	m_lpHostDlg->GetClientRect(pRect);

	m_flSplitPercent = 0.5;

	m_iSplitWidth = 6;

	m_PaneOne = NULL;
	m_PaneTwo = NULL;

	m_SplitType = SplitHorizontal;//this doesn't mean anything yet
	m_iSplitPos = 1;

	WNDCLASSEX wc = {0};
	HINSTANCE hInst = AfxGetInstanceHandle();
	if(!(::GetClassInfoEx(hInst, _T("FakeSplitter"), &wc)))// STRING_OK
	{
		wc.cbSize  = sizeof(wc);
		wc.style   = 0; //CS_VREDRAW | CS_HREDRAW;//not passing these fixes flicker
		wc.lpszClassName = _T("FakeSplitter");// STRING_OK
		wc.lpfnWndProc = ::DefWindowProc;

		WC_D(wc.hbrBackground, (HBRUSH) GetStockObject(BLACK_BRUSH));//helps spot flashing

		RegisterClassEx(&wc);
	}

	EC_B(Create(
		_T("FakeSplitter"),// STRING_OK
		_T("FakeSplitter"),// STRING_OK
		WS_CHILD
		| WS_CLIPSIBLINGS
		| WS_CLIPCHILDREN //required to reduce flicker
//		| WS_BORDER
		| WS_VISIBLE,
		pRect,
		lpHostDlg,
		IDC_FAKE_SPLITTER));

	//Necessary for TAB to work. Without this, all TABS get stuck on the fake splitter control
	//instead of passing to the children. Haven't tested with nested splitters.
	EC_B(ModifyStyleEx(0,WS_EX_CONTROLPARENT));
}

CFakeSplitter::~CFakeSplitter()
{
	TRACE_DESTRUCTOR(CLASS);
	DestroyWindow();
	if (m_lpHostDlg) m_lpHostDlg->Release();
}

BEGIN_MESSAGE_MAP(CFakeSplitter, CWnd)
//{{AFX_MSG_MAP(CFakeSplitter)
	ON_WM_MOUSEMOVE()
	ON_WM_SIZE()
	ON_WM_PAINT()
//}}AFX_MSG_MAP
END_MESSAGE_MAP()

//Dummy code in case I wanna step into the window handling
LRESULT CFakeSplitter::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	LRESULT lResult = NULL;
	DebugPrint(DBGWindowProc,_T("CFakeSplitter::WindowProc message = 0x%x, wParam = 0x%X, lParam = 0x%X\n"),message,wParam,lParam);

	switch (message)
	{
	case WM_HELP:
		return TRUE;
		break;
	case WM_ERASEBKGND:
		{
			return TRUE;
			break;
		}
	case WM_LBUTTONUP:
	case WM_CANCELMODE://Called if focus changes while we're adjusting the splitter
		{
			StopTracking();
			return lResult;
			break;
		}
	case WM_LBUTTONDOWN:
		{
			if (m_bTracking) return TRUE;
			StartTracking(HitTest(LOWORD(lParam),HIWORD(lParam)));
			return lResult;
			break;
		}
	case WM_SETCURSOR:
		{
			if (LOWORD(lParam) == HTCLIENT &&
				(HWND) wParam == this->m_hWnd &&
				!m_bTracking) return TRUE;// we will handle it in the mouse move
			break;
		}

	}//end switch
	return CWnd::WindowProc(message,wParam,lParam);
}

void CFakeSplitter::SetPaneOne(CWnd* PaneOne)
{
	m_PaneOne = PaneOne;
}

void CFakeSplitter::SetPaneTwo(CWnd* PaneTwo)
{
	m_PaneTwo = PaneTwo;
}

void CFakeSplitter::OnSize(UINT /*nType*/, int cx, int cy)
{
	HRESULT hRes = S_OK;
	HDWP hdwp = NULL;

	CalcSplitPos();

	WC_D(hdwp, BeginDeferWindowPos(2));

	cy--;
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
				m_iSplitPos+m_iSplitWidth,//new x
				0,//new y
				cx,// - iSplitPos-m_iSplitWidth,//new width
				cy);//new height
		}
		else
		{
			r2.SetRect(
				0,//new x
				m_iSplitPos+m_iSplitWidth,//new y
				cx,//new width
				cy);// - iSplitPos-m_iSplitWidth);//new height
		}
		DeferWindowPos(hdwp,m_PaneTwo->m_hWnd,0,r2.left,r2.top,r2.Width(),r2.Height(),SWP_NOZORDER);
	}
	EC_B(EndDeferWindowPos(hdwp));

	//Invalidate our splitter region to force a redraw
	if (SplitHorizontal == m_SplitType)
	{
		InvalidateRect(CRect(m_iSplitPos,0,m_iSplitPos+m_iSplitWidth,cy),false);
	}
	else
	{
		InvalidateRect(CRect(0,m_iSplitPos,cx,m_iSplitPos+m_iSplitWidth),false);
	}
}

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
}

BOOL CFakeSplitter::SetPercent(FLOAT iNewPercent)
{
	if (iNewPercent < 0.0 || iNewPercent > 1.0)
		return FALSE;
	m_flSplitPercent = iNewPercent;

	CalcSplitPos();

	CRect rect;
	GetClientRect(rect);

	//Recalculate our layout
	OnSize(0,rect.Width(),rect.Height());
	return TRUE;
}

int CFakeSplitter::HitTest(LONG x, LONG y)
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
}

void CFakeSplitter::OnMouseMove(UINT /*nFlags*/,CPoint point)
{
	HRESULT hRes = S_OK;
	//If we don't have GetCapture, then we don't want to track right now.
	if (GetCapture() != this)
		StopTracking();

	if (SplitterHit == HitTest(point.x, point.y))
	{
		HCURSOR hSplitCursor;
		if (SplitHorizontal == m_SplitType)
		{
			EC_D(hSplitCursor,::LoadCursor(NULL, IDC_SIZEWE));
		}
		else
		{
			EC_D(hSplitCursor,::LoadCursor(NULL, IDC_SIZENS));
		}
		::SetCursor(hSplitCursor);
		EC_B(::DestroyCursor(hSplitCursor)); // destroy after being set
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

		//Force child windows to refresh now
		EC_B(RedrawWindow(NULL, NULL, RDW_ALLCHILDREN | RDW_UPDATENOW));
	}
}

void CFakeSplitter::StartTracking(int ht)
{
	HRESULT hRes = S_OK;
	ASSERT_VALID(this);
	if (ht == noHit)
		return;

	// steal focus and capture
	SetCapture();
	SetFocus();

	// make sure no updates are pending
	//CSplitterWnd does this...not sure why
	EC_B(RedrawWindow(NULL, NULL, RDW_ALLCHILDREN | RDW_UPDATENOW));

	// set tracking state and appropriate cursor
	m_bTracking = TRUE;
}

void CFakeSplitter::StopTracking()
{
	ASSERT_VALID(this);

	if (!m_bTracking)
		return;

	ReleaseCapture();

	m_bTracking = FALSE;
}

//See CSplitterWnd::OnPaint to see where I swiped this code.
void CFakeSplitter::OnPaint()
{
	HRESULT hRes = S_OK;
	PAINTSTRUCT ps;
	CDC* dc = NULL;
	EC_D(dc,BeginPaint(&ps));

	if (dc)
	{
		//DebugPrintEx(DBGGeneric,CLASS,_T("OnPaint"),_T("%d,%d,%d,%d\n"),ps.rcPaint.left,ps.rcPaint.right,ps.rcPaint.top,ps.rcPaint.bottom);

		CRect rect = ps.rcPaint;

		//Shouldn't need to worry about this now - InvalidateRect took care of it for us
		//Draw the splitter bar
		if (SplitHorizontal == m_SplitType)
		{
			rect.left = m_iSplitPos;
			rect.right = m_iSplitPos + m_iSplitWidth;
		}
		else
		{
			rect.top = m_iSplitPos;
			rect.bottom = m_iSplitPos + m_iSplitWidth;
		}

		dc->Draw3dRect(rect, ::GetSysColor(COLOR_BTNHIGHLIGHT), ::GetSysColor(COLOR_BTNSHADOW));
		//dc->Draw3dRect(rect, RGB(255, 0, 0), RGB(0, 255, 0));//red/green

		if (SplitHorizontal == m_SplitType)
		{
			rect.InflateRect(-1,0,-1,0);
		}
		else
		{
			rect.InflateRect(0,-1,0,-1);
		}

		// Draw3dRect only draws the edges: fill the middle
		dc->FillSolidRect(rect, ::GetSysColor(COLOR_BTNFACE));
	}

	EndPaint(&ps);
}