#include <StdAfx.h>
#include <UI/Controls/PaneHeader/PaneHeader2.h>
#include <UI/UIFunctions.h>
#include <core/utility/strings.h>
#include <UI/DoubleBuffer.h>
#include <core/utility/output.h>

namespace controls
{
	static std::wstring CLASS = L"PaneHeader2";

	PaneHeader2::~PaneHeader2()
	{
		TRACE_DESTRUCTOR(CLASS);
		CWnd::DestroyWindow();
	}

	void PaneHeader2::Initialize(HWND hWnd, _In_opt_ HDC hdc, _In_ UINT nid)
	{
		m_hwndParent = hWnd;

		// Assign a nID to the collapse button that is IDD_COLLAPSE more than the control's nID
		m_nID = nid;
		m_nIDCollapse = nid + IDD_COLLAPSE;

		WNDCLASSEX wc = {};
		const auto hInst = AfxGetInstanceHandle();
		if (!::GetClassInfoEx(hInst, _T("PaneHeader"), &wc)) // STRING_OK
		{
			wc.cbSize = sizeof wc;
			wc.style = 0; // not passing CS_VREDRAW | CS_HREDRAW fixes flicker
			wc.lpszClassName = _T("PaneHeader"); // STRING_OK
			wc.lpfnWndProc = ::DefWindowProc;
			wc.hbrBackground = GetSysBrush(ui::uiColor::Background); // helps spot flashing

			RegisterClassEx(&wc);
		}

		// WS_CLIPCHILDREN is used to reduce flicker
		EC_B_S(CreateEx(
			0,
			_T("PaneHeader"), // STRING_OK
			_T("PaneHeader"), // STRING_OK
			WS_CHILD | WS_CLIPSIBLINGS | WS_CLIPCHILDREN | WS_VISIBLE,
			0,
			0,
			0,
			0,
			m_hwndParent,
			reinterpret_cast<HMENU>(static_cast<INT_PTR>(IDC_PANE_HEADER)),
			nullptr));

		// Necessary for TAB to work. Without this, all TABS get stuck on the control
		// instead of passing to the children.
		EC_B_S(ModifyStyleEx(0, WS_EX_CONTROLPARENT));

		const auto sizeText = ui::GetTextExtentPoint32(hdc, m_szLabel);
		m_iLabelWidth = sizeText.cx;
		output::DebugPrint(
			output::dbgLevel::Draw,
			L"PaneHeader::Initialize m_iLabelWidth:%d \"%ws\"\n",
			m_iLabelWidth,
			m_szLabel.c_str());
	}

	BEGIN_MESSAGE_MAP(PaneHeader2, CWnd)
	ON_WM_SIZE()
	ON_WM_PAINT()
	ON_WM_CREATE()
	END_MESSAGE_MAP()

	int PaneHeader2::OnCreate(LPCREATESTRUCT /*lpCreateStruct*/)
	{
		EC_B_S(m_leftLabel.Create(
			WS_CHILD | WS_CLIPSIBLINGS | ES_READONLY | WS_VISIBLE | WS_TABSTOP, CRect(0, 0, 0, 0), this, m_nID));
		::SetWindowTextW(m_leftLabel.m_hWnd, m_szLabel.c_str());
		ui::SubclassLabel(m_leftLabel.m_hWnd);
		if (m_bCollapsible)
		{
			StyleLabel(m_leftLabel.m_hWnd, ui::uiLabelStyle::PaneHeaderLabel);
		}

		EC_B_S(m_rightLabel.Create(
			WS_CHILD | WS_CLIPSIBLINGS | ES_READONLY | WS_VISIBLE | WS_TABSTOP,
			CRect(0, 0, 0, 0),
			this,
			IDD_COUNTLABEL));
		ui::SubclassLabel(m_rightLabel.m_hWnd);
		StyleLabel(m_rightLabel.m_hWnd, ui::uiLabelStyle::PaneHeaderText);

		SetRightLabel(L"Hello world");

		if (m_bCollapsible)
		{
			EC_B_S(m_CollapseButton.Create(
				nullptr, WS_TABSTOP | WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE, CRect(0, 0, 0, 0), this, m_nIDCollapse));
			StyleButton(
				m_CollapseButton.m_hWnd, m_bCollapsed ? ui::uiButtonStyle::UpArrow : ui::uiButtonStyle::DownArrow);
		}

		// If we need an action button, go ahead and create it
		if (m_nIDAction)
		{
			EC_B_S(m_actionButton.Create(
				strings::wstringTotstring(m_szActionButton).c_str(),
				WS_TABSTOP | WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE,
				CRect(0, 0, 0, 0),
				this,
				m_nIDAction));
			StyleButton(m_actionButton.m_hWnd, ui::uiButtonStyle::Unstyled);
		}

		m_bInitialized = true;
		return 0;
	}

	LRESULT PaneHeader2::WindowProc(const UINT message, const WPARAM wParam, const LPARAM lParam)
	{
		LRESULT lRes = 0;
		if (ui::HandleControlUI(message, wParam, lParam, &lRes)) return lRes;

		switch (message)
		{
		case WM_CLOSE:
			::SendMessage(m_hwndParent, message, wParam, lParam);
			return true;
		case WM_HELP:
			return true;
		case WM_COMMAND:
		{
			const auto nCode = HIWORD(wParam);
			if (EN_CHANGE == nCode || CBN_SELCHANGE == nCode || CBN_EDITCHANGE == nCode || BN_CLICKED == nCode)
			{
				::SendMessage(m_hwndParent, message, wParam, lParam);
			}

			break;
		}
		}

		return CWnd::WindowProc(message, wParam, lParam);
	}

	void PaneHeader2::OnSize(UINT /*nType*/, const int cx, const int cy)
	{
		auto hdwp = WC_D(HDWP, BeginDeferWindowPos(2));
		if (hdwp)
		{
			hdwp = EC_D(HDWP, DeferWindowPos(hdwp, 0, 0, cx, cy));
			EC_B_S(EndDeferWindowPos(hdwp));
		}
	}

	// Draw our collapse button and label, if needed.
	// Draws everything to GetFixedHeight()
	HDWP PaneHeader2::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		const _In_ int x,
		const _In_ int y,
		const _In_ int width,
		const _In_ int height)
	{
		if (!m_bInitialized) return hWinPosInfo;
		output::DebugPrint(
			output::dbgLevel::Draw,
			L"PaneHeader2::DeferWindowPos x:%d y:%d width:%d height:%d v:%d\n",
			x,
			y,
			width,
			height,
			IsWindowVisible());
		InvalidateRect(CRect(x, y, width, height), false);

		//hWinPosInfo = ui::DeferWindowPos(
		//	hWinPosInfo, GetSafeHwnd(), x, y, width, height, L"PaneHeader::DeferWindowPos", m_szLabel.c_str());

		//auto curX = x;
		const auto actionButtonWidth = m_actionButtonWidth ? m_actionButtonWidth + 2 * m_iMargin : 0;
		const auto actionButtonAndGutterWidth = actionButtonWidth ? actionButtonWidth + m_iSideMargin : 0;
		if (m_bCollapsible)
		{
			//	hWinPosInfo = ui::DeferWindowPos(
			//		hWinPosInfo,
			//		m_CollapseButton.GetSafeHwnd(),
			//		curX,
			//		y,
			//		width - actionButtonAndGutterWidth,
			//		height,
			//		L"PaneHeader::DeferWindowPos::collapseButton");
			//	curX += m_iButtonHeight;
		}

		//hWinPosInfo = ui::DeferWindowPos(
		//	hWinPosInfo, GetSafeHwnd(), curX, y, m_iLabelWidth, height, L"PaneHeader::DeferWindowPos::leftLabel");

		if (!m_bCollapsed)
		{
			// Drop the count on top of the label we drew above
			if (m_rightLabel.GetSafeHwnd())
			{
				hWinPosInfo = ui::DeferWindowPos(
					hWinPosInfo,
					m_rightLabel.GetSafeHwnd(),
					x + width - m_rightLabelWidth - actionButtonAndGutterWidth,
					y,
					m_rightLabelWidth,
					height,
					L"PaneHeader::DeferWindowPos::rightLabel");
			}
		}

		if (m_nIDAction)
		{
			// Drop the action button next to the label we drew above
			hWinPosInfo = ui::DeferWindowPos(
				hWinPosInfo,
				m_actionButton.GetSafeHwnd(),
				x + width - actionButtonWidth,
				y,
				actionButtonWidth,
				height,
				L"PaneHeader::DeferWindowPos::actionButton");
		}

		output::DebugPrint(output::dbgLevel::Draw, L"PaneHeader::DeferWindowPos end\n");
		return hWinPosInfo;
	}

	void PaneHeader2::OnPaint()
	{
		auto ps = PAINTSTRUCT{};
		::BeginPaint(m_hWnd, &ps);
		if (ps.hdc)
		{
			auto rcWin = RECT{};
			::GetClientRect(m_hWnd, &rcWin);
			ui::CDoubleBuffer db;
			auto hdc = ps.hdc;
			db.Begin(hdc, rcWin);
			FillRect(hdc, &rcWin, GetSysBrush(ui::uiColor::Background));
			db.End(hdc);
		}

		::EndPaint(m_hWnd, &ps);
	}

	int PaneHeader2::GetMinWidth()
	{
		auto cx = m_iLabelWidth;
		if (m_bCollapsible) cx += m_iButtonHeight;
		if (m_rightLabelWidth)
		{
			cx += m_iSideMargin;
			cx += m_rightLabelWidth;
		}

		if (m_actionButtonWidth)
		{
			cx += m_iSideMargin;
			cx += m_actionButtonWidth + 2 * m_iMargin;
		}

		return cx;
	}

	void PaneHeader2::SetRightLabel(const std::wstring szLabel)
	{
		//if (!m_bInitialized) return;
		EC_B_S(::SetWindowTextW(m_rightLabel.m_hWnd, szLabel.c_str()));

		const auto hdc = ::GetDC(m_rightLabel.GetSafeHwnd());
		const auto hfontOld = SelectObject(hdc, ui::GetSegoeFont());
		const auto sizeText = ui::GetTextExtentPoint32(hdc, szLabel);
		static_cast<void>(SelectObject(hdc, hfontOld));
		::ReleaseDC(m_rightLabel.GetSafeHwnd(), hdc);
		m_rightLabelWidth = sizeText.cx;
	}

	void PaneHeader2::SetActionButton(const std::wstring szActionButton)
	{
		// Don't bother if we never enabled the button
		if (m_nIDAction == 0) return;

		EC_B_S(::SetWindowTextW(m_actionButton.GetSafeHwnd(), szActionButton.c_str()));

		m_szActionButton = szActionButton;
		const auto hdc = ::GetDC(m_actionButton.GetSafeHwnd());
		const auto hfontOld = SelectObject(hdc, ui::GetSegoeFont());
		const auto sizeText = ui::GetTextExtentPoint32(hdc, szActionButton);
		static_cast<void>(SelectObject(hdc, hfontOld));
		::ReleaseDC(m_actionButton.GetSafeHwnd(), hdc);
		m_actionButtonWidth = sizeText.cx;
	}

	bool PaneHeader2::HandleChange(UINT nID)
	{
		// Collapse buttons have a nID IDD_COLLAPSE higher than nID of the pane they toggle.
		// So if we get asked about one that matches, we can assume it's time to toggle our collapse.
		if (m_nIDCollapse == nID)
		{
			OnToggleCollapse();
			return true;
		}

		return false;
	}

	void PaneHeader2::OnToggleCollapse()
	{
		m_bCollapsed = !m_bCollapsed;

		StyleButton(m_CollapseButton.m_hWnd, m_bCollapsed ? ui::uiButtonStyle::UpArrow : ui::uiButtonStyle::DownArrow);
		WC_B_S(m_rightLabel.ShowWindow(m_bCollapsed ? SW_HIDE : SW_SHOW));

		// Trigger a redraw
		::PostMessage(m_hwndParent, WM_COMMAND, IDD_RECALCLAYOUT, NULL);
	}

	void PaneHeader2::SetMargins(
		int iMargin,
		int iSideMargin,
		int iLabelHeight, // Height of the label
		int iButtonHeight) // Height of button
	{
		m_iMargin = iMargin;
		m_iSideMargin = iSideMargin;
		m_iLabelHeight = iLabelHeight;
		m_iButtonHeight = iButtonHeight;
	}
} // namespace controls