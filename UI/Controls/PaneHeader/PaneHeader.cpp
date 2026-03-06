#include <StdAfx.h>
#include <UI/Controls/PaneHeader/PaneHeader.h>
#include <UI/UIFunctions.h>
#include <core/utility/strings.h>
#include <UI/DoubleBuffer.h>
#include <core/utility/output.h>
#include <core/utility/registry.h>

namespace controls
{
	static std::wstring CLASS = L"PaneHeader";

	PaneHeader::~PaneHeader()
	{
		TRACE_DESTRUCTOR(CLASS);
		CWnd::DestroyWindow();
	}

	void PaneHeader::Initialize(HWND hWnd, _In_opt_ HDC hdc, _In_ UINT nid)
	{
		m_hWndParent = hWnd;

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
			m_hWndParent,
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

	BEGIN_MESSAGE_MAP(PaneHeader, CWnd)
	ON_WM_CREATE()
	END_MESSAGE_MAP()

	int PaneHeader::OnCreate(LPCREATESTRUCT /*lpCreateStruct*/)
	{
		EC_B_S(m_leftLabel.Create(nullptr, WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE, CRect(0, 0, 0, 0), this, m_nID));
		::SetWindowTextW(m_leftLabel.m_hWnd, m_szLabel.c_str());
		ui::SubclassLabel(m_leftLabel.m_hWnd);
		if (m_bCollapsible)
		{
			StyleLabel(m_leftLabel.m_hWnd, ui::uiLabelStyle::PaneHeaderLabel);
		}

		EC_B_S(m_rightLabel.Create(
			nullptr, WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE, CRect(0, 0, 0, 0), this, IDD_RIGHTLABEL));
		ui::SubclassLabel(m_rightLabel.m_hWnd);
		StyleLabel(m_rightLabel.m_hWnd, ui::uiLabelStyle::PaneHeaderText);

		if (m_bCollapsible)
		{
			EC_B_S(m_CollapseButton.Create(
				nullptr,
				WS_TABSTOP | WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE | BS_NOTIFY,
				CRect(0, 0, 0, 0),
				this,
				m_nIDCollapse));
			StyleButton(
				m_CollapseButton.m_hWnd, m_bCollapsed ? ui::uiButtonStyle::UpArrow : ui::uiButtonStyle::DownArrow);
		}

		// If we need an action button, go ahead and create it
		if (m_nIDAction)
		{
			EC_B_S(m_actionButton.Create(
				strings::wstringTotstring(m_szActionButton).c_str(),
				WS_TABSTOP | WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE | BS_NOTIFY,
				CRect(0, 0, 0, 0),
				this,
				m_nIDAction));
			StyleButton(m_actionButton.m_hWnd, ui::uiButtonStyle::Unstyled);
		}

		m_bInitialized = true;
		return 0;
	}

	LRESULT PaneHeader::WindowProc(const UINT message, const WPARAM wParam, const LPARAM lParam)
	{
		LRESULT lRes = 0;
		if (ui::HandleControlUI(message, wParam, lParam, &lRes)) return lRes;

		switch (message)
		{
		case WM_PAINT:
		{
			auto ps = PAINTSTRUCT{};
			::BeginPaint(m_hWnd, &ps);
			if (ps.hdc)
			{
				auto rcWin = RECT{};
				::GetClientRect(m_hWnd, &rcWin);
				const auto background = registry::uiDiag ? ui::GetSysBrush(ui::uiColor::TestGreen)
														 : ui::GetSysBrush(ui::uiColor::Background);
				FillRect(ps.hdc, &rcWin, background);
			}

			::EndPaint(m_hWnd, &ps);
			return 0;
		}
		case WM_WINDOWPOSCHANGED:
			RecalcLayout();
			return 0;
		case WM_CLOSE:
			::SendMessage(m_hWndParent, message, wParam, lParam);
			return true;
		case WM_HELP:
			return true;
		case WM_COMMAND:
		{
			const auto nCode = HIWORD(wParam);
			switch (nCode)
			{
			// Pass button clicks up to parent
			case BN_CLICKED:
				// Pass focus notifications up to parent
			case BN_SETFOCUS:
			case EN_SETFOCUS:
				return ::SendMessage(m_hWndParent, message, wParam, lParam);
			}

			break;
		}
		}

		return CWnd::WindowProc(message, wParam, lParam);
	}

	int PaneHeader::GetFixedHeight() const noexcept
	{
		int iHeight = 0;
		if (m_bCollapsible || !m_szLabel.empty()) iHeight = max(m_iButtonHeight, m_iLabelHeight);
		if (!m_bCollapsed && iHeight) iHeight += m_iSmallHeightMargin;

		return iHeight;
	}

	// Position collapse button, labels, action button.
	void PaneHeader::RecalcLayout()
	{
		if (!m_bInitialized) return;
		auto hWinPosInfo = WC_D(HDWP, BeginDeferWindowPos(4));
		if (hWinPosInfo)
		{
			auto rcWin = RECT{};
			::GetClientRect(m_hWnd, &rcWin);
			const auto width = rcWin.right - rcWin.left;
			const auto height = rcWin.bottom - rcWin.top;
			output::DebugPrint(
				output::dbgLevel::Draw,
				L"PaneHeader::DeferWindowPos width:%d height:%d v:%d l:\"%ws\"\n",
				width,
				height,
				IsWindowVisible(),
				m_szLabel.c_str());
			auto curX = 0;
			const auto actionButtonWidth =
				!m_bCollapsed && m_actionButtonWidth ? m_actionButtonWidth + 2 * m_iMargin : 0;
			const auto actionButtonAndGutterWidth = actionButtonWidth ? actionButtonWidth + m_iSideMargin : 0;
			const auto labelY = ((m_bCollapsible ? m_iButtonHeight : height) - m_iLabelHeight) / 2;
			if (m_bCollapsible)
			{
				hWinPosInfo = ui::DeferWindowPos(
					hWinPosInfo,
					m_CollapseButton.GetSafeHwnd(),
					curX,
					0,
					width - actionButtonAndGutterWidth,
					m_iButtonHeight,
					L"PaneHeader::DeferWindowPos::collapseButton");
				curX += m_iButtonHeight;
			}

			hWinPosInfo = ui::DeferWindowPos(
				hWinPosInfo,
				m_leftLabel.GetSafeHwnd(),
				curX,
				labelY,
				m_iLabelWidth,
				m_iLabelHeight,
				L"PaneHeader::DeferWindowPos::leftLabel");

			auto cmdShow = SW_HIDE;
			if (!m_bCollapsed && m_rightLabelWidth)
			{
				// Drop the count on top of the label we drew above
				hWinPosInfo = ui::DeferWindowPos(
					hWinPosInfo,
					m_rightLabel.GetSafeHwnd(),
					width - m_rightLabelWidth - actionButtonAndGutterWidth - m_iSideMargin - 1,
					labelY,
					m_rightLabelWidth,
					m_iLabelHeight,
					L"PaneHeader::DeferWindowPos::rightLabel");
				cmdShow = SW_SHOW;
			}

			WC_B_S(::ShowWindow(m_rightLabel.GetSafeHwnd(), cmdShow));

			if (!m_bCollapsed && m_nIDAction && m_actionButtonWidth)
			{
				// Drop the action button next to the label we drew above
				hWinPosInfo = ui::DeferWindowPos(
					hWinPosInfo,
					m_actionButton.GetSafeHwnd(),
					width - actionButtonWidth,
					0,
					actionButtonWidth,
					m_iButtonHeight,
					L"PaneHeader::DeferWindowPos::actionButton");
				cmdShow = SW_SHOW;
			}
			else
			{
				cmdShow = SW_HIDE;
			}

			WC_B_S(::ShowWindow(m_actionButton.GetSafeHwnd(), cmdShow));

			output::DebugPrint(output::dbgLevel::Draw, L"PaneHeader::DeferWindowPos end\n");
			EC_B_S(EndDeferWindowPos(hWinPosInfo));
		}
	}

	int PaneHeader::GetMinWidth()
	{
		auto cx = m_iLabelWidth;
		if (m_bCollapsible) cx += m_iButtonHeight;
		if (m_rightLabelWidth)
		{
			cx += m_iSideMargin * 2;
			cx += m_rightLabelWidth;
		}

		if (m_actionButtonWidth)
		{
			cx += m_iSideMargin;
			cx += m_actionButtonWidth + 2 * m_iMargin;
		}

		return cx;
	}

	void PaneHeader::SetRightLabel(const std::wstring szLabel)
	{
		if (!m_bInitialized) return;
		EC_B_S(::SetWindowTextW(m_rightLabel.m_hWnd, szLabel.c_str()));

		const auto hdc = ::GetDC(m_rightLabel.GetSafeHwnd());
		const auto hfontOld = SelectObject(hdc, ui::GetSegoeFont());
		const auto sizeText = ui::GetTextExtentPoint32(hdc, szLabel);
		static_cast<void>(SelectObject(hdc, hfontOld));
		::ReleaseDC(m_rightLabel.GetSafeHwnd(), hdc);
		m_rightLabelWidth = sizeText.cx;

		RecalcLayout();
	}

	void PaneHeader::SetActionButton(const std::wstring szActionButton)
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

		RecalcLayout();
	}

	bool PaneHeader::HandleChange(UINT nID)
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

	void PaneHeader::OnToggleCollapse()
	{
		if (!m_bCollapsible) return;
		m_bCollapsed = !m_bCollapsed;

		StyleButton(m_CollapseButton.m_hWnd, m_bCollapsed ? ui::uiButtonStyle::UpArrow : ui::uiButtonStyle::DownArrow);

		// Trigger a redraw
		::PostMessage(m_hWndParent, WM_COMMAND, IDD_RECALCLAYOUT, NULL);

		// When we toggle because of VK_ENTER, we don't redraw - signal one just in case
		::RedrawWindow(m_CollapseButton.m_hWnd, nullptr, nullptr, RDW_INVALIDATE | RDW_UPDATENOW);
	}

	void PaneHeader::SetMargins(
		int iMargin,
		int iSideMargin,
		int iLabelHeight, // Height of the label
		int iSmallHeightMargin,
		int iButtonHeight) // Height of button
	{
		m_iMargin = iMargin;
		m_iSideMargin = iSideMargin;
		m_iLabelHeight = iLabelHeight;
		m_iSmallHeightMargin = iSmallHeightMargin;
		m_iButtonHeight = iButtonHeight;
	}

	bool PaneHeader::containsWindow(HWND hWnd) const noexcept
	{
		if (GetSafeHwnd() == hWnd) return true;
		if (m_leftLabel.GetSafeHwnd() == hWnd) return true;
		if (m_rightLabel.GetSafeHwnd() == hWnd) return true;
		if (m_actionButton.GetSafeHwnd() == hWnd) return true;
		if (m_CollapseButton.GetSafeHwnd() == hWnd) return true;
		return false;
	}
} // namespace controls