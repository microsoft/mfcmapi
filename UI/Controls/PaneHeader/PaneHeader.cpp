#include <StdAfx.h>
#include <UI/Controls/PaneHeader/PaneHeader.h>
#include <UI/UIFunctions.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>

namespace controls
{
	void PaneHeader::Initialize(_In_ CWnd* pParent, _In_opt_ HDC hdc, _In_ UINT nid)
	{
		if (pParent) m_hWndParent = pParent->m_hWnd;
		// Assign a nID to the collapse button that is IDD_COLLAPSE more than the control's nID
		m_nIDCollapse = nid + IDD_COLLAPSE;
		// TODO: We don't save our header's nID here, but we could if we wanted

		EC_B_S(Create(
			WS_CHILD | WS_CLIPSIBLINGS | ES_READONLY | WS_VISIBLE | WS_TABSTOP, CRect(0, 0, 0, 0), pParent, nid));
		::SetWindowTextW(m_hWnd, m_szLabel.c_str());
		ui::SubclassLabel(m_hWnd);
		if (m_bCollapsible)
		{
			StyleLabel(m_hWnd, ui::uiLabelStyle::PaneHeaderLabel);
		}

		EC_B_S(m_rightLabel.Create(
			WS_CHILD | WS_CLIPSIBLINGS | ES_READONLY | WS_VISIBLE | WS_TABSTOP,
			CRect(0, 0, 0, 0),
			pParent,
			IDD_COUNTLABEL));
		ui::SubclassLabel(m_rightLabel.m_hWnd);
		StyleLabel(m_rightLabel.m_hWnd, ui::uiLabelStyle::PaneHeaderText);

		if (m_bCollapsible)
		{
			EC_B_S(m_CollapseButton.Create(
				nullptr,
				WS_TABSTOP | WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE,
				CRect(0, 0, 0, 0),
				pParent,
				m_nIDCollapse));
			StyleButton(
				m_CollapseButton.m_hWnd, m_bCollapsed ? ui::uiButtonStyle::UpArrow : ui::uiButtonStyle::DownArrow);
		}

		const auto sizeText = ui::GetTextExtentPoint32(hdc, m_szLabel);
		m_iLabelWidth = sizeText.cx;
		output::DebugPrint(
			output::dbgLevel::Draw,
			L"PaneHeader::Initialize m_iLabelWidth:%d \"%ws\"\n",
			m_iLabelWidth,
			m_szLabel.c_str());

		// If we need an action button, go ahead and create it
		if (m_nIDAction)
		{
			EC_B_S(m_actionButton.Create(
				strings::wstringTotstring(m_szActionButton).c_str(),
				WS_TABSTOP | WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE,
				CRect(0, 0, 0, 0),
				pParent,
				m_nIDAction));
			StyleButton(m_actionButton.m_hWnd, ui::uiButtonStyle::Unstyled);
		}

		m_bInitialized = true;
	}

	// Draw our collapse button and label, if needed.
	// Draws everything to GetFixedHeight()
	HDWP PaneHeader::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		const _In_ int x,
		const _In_ int y,
		const _In_ int width,
		const _In_ int height)
	{
		if (!m_bInitialized) return hWinPosInfo;
		hWinPosInfo = ui::DeferWindowPos(
			hWinPosInfo, GetSafeHwnd(), x, y, width, height, L"PaneHeader::DeferWindowPos", m_szLabel.c_str());

		auto curX = x;
		const auto actionButtonWidth = m_actionButtonWidth ? m_actionButtonWidth + 2 * m_iMargin : 0;
		const auto actionButtonAndGutterWidth = actionButtonWidth ? actionButtonWidth + m_iSideMargin : 0;
		if (m_bCollapsible)
		{
			hWinPosInfo = ui::DeferWindowPos(
				hWinPosInfo,
				m_CollapseButton.GetSafeHwnd(),
				curX,
				y,
				width - actionButtonAndGutterWidth,
				height,
				L"PaneHeader::DeferWindowPos::collapseButton");
			curX += m_iButtonHeight;
		}

		hWinPosInfo = ui::DeferWindowPos(
			hWinPosInfo, GetSafeHwnd(), curX, y, m_iLabelWidth, height, L"PaneHeader::DeferWindowPos::leftLabel");

		if (!m_bCollapsed)
		{
			// Drop the count on top of the label we drew above
			hWinPosInfo = ui::DeferWindowPos(
				hWinPosInfo,
				m_rightLabel.GetSafeHwnd(),
				x + width - m_rightLabelWidth - actionButtonAndGutterWidth,
				y,
				m_rightLabelWidth,
				height,
				L"PaneHeader::DeferWindowPos::rightLabel");
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

	int PaneHeader::GetMinWidth()
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
		m_bCollapsed = !m_bCollapsed;

		StyleButton(m_CollapseButton.m_hWnd, m_bCollapsed ? ui::uiButtonStyle::UpArrow : ui::uiButtonStyle::DownArrow);
		WC_B_S(m_rightLabel.ShowWindow(m_bCollapsed ? SW_HIDE : SW_SHOW));

		// Trigger a redraw
		::PostMessage(m_hWndParent, WM_COMMAND, IDD_RECALCLAYOUT, NULL);
	}

	void PaneHeader::SetMargins(
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