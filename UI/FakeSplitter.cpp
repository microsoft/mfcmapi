#include <StdAfx.h>
#include <UI/FakeSplitter.h>
#include <UI/Dialogs/BaseDialog.h>
#include <UI/UIFunctions.h>
#include <UI/DoubleBuffer.h>

namespace controls
{
	static std::wstring CLASS = L"CFakeSplitter";

	enum FakesSplitHitTestValue
	{
		noHit = 0,
		SplitterHit = 1
	};

	CFakeSplitter::~CFakeSplitter()
	{
		TRACE_DESTRUCTOR(CLASS);
		(void) DestroyCursor(m_hSplitCursorH);
		(void) DestroyCursor(m_hSplitCursorV);
		CWnd::DestroyWindow();
	}

	void CFakeSplitter::Init(HWND hWnd)
	{
		WNDCLASSEX wc = {0};
		const auto hInst = AfxGetInstanceHandle();
		if (!::GetClassInfoEx(hInst, _T("FakeSplitter"), &wc)) // STRING_OK
		{
			wc.cbSize = sizeof wc;
			wc.style = 0; // not passing CS_VREDRAW | CS_HREDRAW fixes flicker
			wc.lpszClassName = _T("FakeSplitter"); // STRING_OK
			wc.lpfnWndProc = ::DefWindowProc;
			wc.hbrBackground = ui::GetSysBrush(ui::cBackground); // helps spot flashing

			RegisterClassEx(&wc);
		}

		// WS_CLIPCHILDREN is used to reduce flicker
		EC_B_S(CreateEx(
			0,
			_T("FakeSplitter"), // STRING_OK
			_T("FakeSplitter"), // STRING_OK
			WS_CHILD | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | WS_VISIBLE,
			0,
			0,
			0,
			0,
			hWnd,
			reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_FAKE_SPLITTER)),
			nullptr));

		// Necessary for TAB to work. Without this, all TABS get stuck on the fake splitter control
		// instead of passing to the children. Haven't tested with nested splitters.
		EC_B_S(ModifyStyleEx(0, WS_EX_CONTROLPARENT));

		// Load split cursors
		m_hSplitCursorV = EC_D(HCURSOR, ::LoadCursor(GetModuleHandle(nullptr), MAKEINTRESOURCE(IDC_SPLITV)));
		m_hSplitCursorH = EC_D(HCURSOR, ::LoadCursor(GetModuleHandle(nullptr), MAKEINTRESOURCE(IDC_SPLITH)));
	}

	BEGIN_MESSAGE_MAP(CFakeSplitter, CWnd)
	ON_WM_MOUSEMOVE()
	ON_WM_SIZE()
	ON_WM_PAINT()
	END_MESSAGE_MAP()

	LRESULT CFakeSplitter::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
	{
		switch (message)
		{
		case WM_HELP:
			return true;
		case WM_ERASEBKGND:
			return true;
		case WM_LBUTTONUP:
		case WM_CANCELMODE: // Called if focus changes while we're adjusting the splitter
			StopTracking();
			return NULL;
		case WM_LBUTTONDOWN:
			if (m_bTracking) return true;
			StartTracking(HitTest(LOWORD(lParam), HIWORD(lParam)));
			return NULL;
		case WM_SETCURSOR:
			if (LOWORD(lParam) == HTCLIENT && reinterpret_cast<HWND>(wParam) == this->m_hWnd && !m_bTracking)
				return true; // we will handle it in the mouse move
			break;
		}

		return CWnd::WindowProc(message, wParam, lParam);
	}

	void CFakeSplitter::SetPaneOne(HWND paneOne)
	{
		m_PaneOne = paneOne;
		if (m_PaneOne)
		{
			m_iSplitWidth = 7;
		}
		else
		{
			m_iSplitWidth = 0;
		}
	}

	void CFakeSplitter::SetPaneTwo(HWND paneTwo) { m_PaneTwo = paneTwo; }

	void CFakeSplitter::OnSize(UINT /*nType*/, int cx, int cy)
	{
		CalcSplitPos();

		const auto hdwp = WC_D(HDWP, BeginDeferWindowPos(2));
		if (hdwp)
		{
			if (m_PaneOne || m_ViewPaneOne)
			{
				CRect r1;
				if (SplitHorizontal == m_SplitType)
				{
					r1.SetRect(0, 0, m_iSplitPos, cy);
				}
				else
				{
					r1.SetRect(0, 0, cx, m_iSplitPos);
				}

				if (m_PaneOne)
				{
					DeferWindowPos(hdwp, m_PaneOne, nullptr, 0, 0, r1.Width(), r1.Height(), SWP_NOZORDER);
				}

				if (m_ViewPaneOne)
				{
					m_ViewPaneOne->DeferWindowPos(hdwp, 0, 0, r1.Width(), r1.Height());
				}
			}

			if (m_PaneTwo || m_ViewPaneTwo)
			{
				CRect r2;
				if (SplitHorizontal == m_SplitType)
				{
					r2.SetRect(
						m_iSplitPos + m_iSplitWidth, // new x
						0, // new y
						cx, // bottom right corner
						cy); // bottom right corner
				}
				else
				{
					r2.SetRect(
						0, // new x
						m_iSplitPos + m_iSplitWidth, // new y
						cx, // bottom right corner
						cy); // bottom right corner
				}

				if (m_PaneTwo)
				{
					DeferWindowPos(hdwp, m_PaneTwo, nullptr, r2.left, r2.top, r2.Width(), r2.Height(), SWP_NOZORDER);
				}

				if (m_ViewPaneTwo)
				{
					m_ViewPaneTwo->DeferWindowPos(hdwp, r2.left, r2.top, r2.Width(), r2.Height());
				}
			}

			EC_B_S(EndDeferWindowPos(hdwp));
		}

		if ((m_PaneOne || m_ViewPaneOne) && (m_PaneTwo || m_ViewPaneTwo))
		{
			// Invalidate our splitter region to force a redraw
			if (SplitHorizontal == m_SplitType)
			{
				InvalidateRect(CRect(m_iSplitPos, 0, m_iSplitPos + m_iSplitWidth, cy), false);
			}
			else
			{
				InvalidateRect(CRect(0, m_iSplitPos, cx, m_iSplitPos + m_iSplitWidth), false);
			}
		}
	}

	void CFakeSplitter::CalcSplitPos()
	{
		if (!m_PaneOne && !m_ViewPaneOne)
		{
			m_iSplitPos = 0;
			return;
		}

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

		m_iSplitPos = static_cast<int>(static_cast<FLOAT>(iCurSpan) * m_flSplitPercent);
		if (m_iSplitPos + m_iSplitWidth + 1 >= iCurSpan)
		{
			m_iSplitPos = m_iSplitPos - m_iSplitWidth - 1;
		}
	}

	void CFakeSplitter::SetPercent(FLOAT iNewPercent)
	{
		if (iNewPercent < 0.0 || iNewPercent > 1.0) return;
		m_flSplitPercent = iNewPercent;

		CalcSplitPos();

		CRect rect;
		GetClientRect(rect);

		// Recalculate our layout
		OnSize(0, rect.Width(), rect.Height());
	}

	void CFakeSplitter::SetSplitType(SplitType stSplitType) { m_SplitType = stSplitType; }

	_Check_return_ int CFakeSplitter::HitTest(LONG x, LONG y) const
	{
		if (!m_PaneOne && !m_ViewPaneOne) return noHit;

		LONG lTestPos;

		if (SplitHorizontal == m_SplitType)
		{
			lTestPos = x;
		}
		else
		{
			lTestPos = y;
		}

		if (lTestPos >= m_iSplitPos && lTestPos <= m_iSplitPos + m_iSplitWidth) return SplitterHit;

		return noHit;
	}

	void CFakeSplitter::OnMouseMove(UINT /*nFlags*/, CPoint point)
	{
		if (!m_PaneOne && !m_ViewPaneOne) return;

		// If we don't have GetCapture, then we don't want to track right now.
		if (GetCapture() != this) StopTracking();

		if (SplitterHit == HitTest(point.x, point.y))
		{
			// This looks backwards, but it is not. A horizontal split needs the vertical cursor
			SetCursor(SplitHorizontal == m_SplitType ? m_hSplitCursorV : m_hSplitCursorH);
		}

		if (m_bTracking)
		{
			CRect Rect;
			FLOAT flNewPercent;
			GetWindowRect(Rect);

			if (SplitHorizontal == m_SplitType)
			{
				flNewPercent = point.x / static_cast<FLOAT>(Rect.Width());
			}
			else
			{
				flNewPercent = point.y / static_cast<FLOAT>(Rect.Height());
			}
			SetPercent(flNewPercent);

			// Force child windows to refresh now
			EC_B_S(RedrawWindow(nullptr, nullptr, RDW_ALLCHILDREN | RDW_UPDATENOW));
		}
	}

	void CFakeSplitter::StartTracking(int ht)
	{
		if (ht == noHit) return;

		// steal focus and capture
		SetCapture();
		SetFocus();

		// make sure no updates are pending
		// CSplitterWnd does this...not sure why
		EC_B_S(RedrawWindow(nullptr, nullptr, RDW_ALLCHILDREN | RDW_UPDATENOW));

		// set tracking state and appropriate cursor
		m_bTracking = true;

		// Force redraw to get our tracking gripper
		SetPercent(m_flSplitPercent);
	}

	void CFakeSplitter::StopTracking()
	{
		if (!m_bTracking) return;

		ReleaseCapture();

		m_bTracking = false;

		// Force redraw to get our non-tracking gripper
		SetPercent(m_flSplitPercent);
	}

	void CFakeSplitter::OnPaint()
	{
		PAINTSTRUCT ps = {nullptr};
		::BeginPaint(m_hWnd, &ps);
		if (ps.hdc)
		{
			RECT rcWin = {0};
			::GetClientRect(m_hWnd, &rcWin);
			ui::CDoubleBuffer db;
			auto hdc = ps.hdc;
			db.Begin(hdc, rcWin);

			auto rcSplitter = rcWin;
			FillRect(hdc, &rcSplitter, ui::GetSysBrush(ui::cBackground));

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
			const auto hpenOld = SelectObject(hdc, ui::GetPen(m_bTracking ? ui::cSolidPen : ui::cDashedPen));
			MoveToEx(hdc, pts[0].x, pts[0].y, nullptr);
			LineTo(hdc, pts[1].x, pts[1].y);
			(void) SelectObject(hdc, hpenOld);

			db.End(hdc);
		}

		::EndPaint(m_hWnd, &ps);
	}
} // namespace controls