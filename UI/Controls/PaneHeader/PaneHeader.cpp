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

		EC_B_S(m_Count.Create(
			WS_CHILD | WS_CLIPSIBLINGS | ES_READONLY | WS_VISIBLE, CRect(0, 0, 0, 0), pParent, IDD_COUNTLABEL));
		ui::SubclassLabel(m_Count.m_hWnd);
		StyleLabel(m_Count.m_hWnd, ui::uiLabelStyle::PaneHeaderText);

		EC_B_S(Create(WS_CHILD | WS_CLIPSIBLINGS | ES_READONLY | WS_VISIBLE, CRect(0, 0, 0, 0), pParent, nid));
		::SetWindowTextW(m_hWnd, m_szLabel.c_str());
		ui::SubclassLabel(m_hWnd);

		if (m_bCollapsible)
		{
			StyleLabel(m_hWnd, ui::uiLabelStyle::PaneHeaderLabel);

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
	}

	// Draw our collapse button and label, if needed.
	// Draws everything to GetFixedHeight()
	void PaneHeader::DeferWindowPos(_In_ HDWP hWinPosInfo, const _In_ int x, const _In_ int y, const _In_ int width)
	{
		const auto height = GetFixedHeight();
		auto curX = x;
		if (m_bCollapsible)
		{
			::DeferWindowPos(
				hWinPosInfo, m_CollapseButton.GetSafeHwnd(), nullptr, curX, y, width, height, SWP_NOZORDER);
			curX += m_iButtonHeight;
		}

		output::DebugPrint(
			output::dbgLevel::Draw,
			L"PaneHeader::DeferWindowPos x:%d width:%d labelpos:%d labelwidth:%d \n",
			x,
			width,
			curX,
			m_iLabelWidth);

		::DeferWindowPos(hWinPosInfo, GetSafeHwnd(), nullptr, curX, y, m_iLabelWidth, height, SWP_NOZORDER);

		if (!m_bCollapsed)
		{
			// Drop the count on top of the label we drew above
			EC_B_S(::DeferWindowPos(
				hWinPosInfo,
				m_Count.GetSafeHwnd(),
				nullptr,
				x + width - m_iCountLabelWidth,
				y,
				m_iCountLabelWidth,
				height,
				SWP_NOZORDER));
		}
	}

	int PaneHeader::GetMinWidth()
	{
		auto cx = m_iLabelWidth;
		if (m_bCollapsible) cx += m_iButtonHeight;
		if (m_iCountLabelWidth)
		{
			cx += m_iSideMargin;
			cx += m_iCountLabelWidth;
		}

		return cx;
	}

	void PaneHeader::SetCount(const std::wstring szCount)
	{
		::SetWindowTextW(m_Count.m_hWnd, szCount.c_str());

		const auto hdc = ::GetDC(m_Count.GetSafeHwnd());
		const auto hfontOld = SelectObject(hdc, ui::GetSegoeFont());
		const auto sizeText = ui::GetTextExtentPoint32(hdc, szCount);
		static_cast<void>(SelectObject(hdc, hfontOld));
		::ReleaseDC(m_Count.GetSafeHwnd(), hdc);
		m_iCountLabelWidth = sizeText.cx + m_iSideMargin;
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
		WC_B_S(m_Count.ShowWindow(m_bCollapsed ? SW_HIDE : SW_SHOW));

		// Trigger a redraw
		::PostMessage(m_hWndParent, WM_COMMAND, IDD_RECALCLAYOUT, NULL);
	}

	void PaneHeader::SetMargins(
		int iSideMargin,
		int iLabelHeight, // Height of the label
		int iButtonHeight) // Height of button
	{
		m_iSideMargin = iSideMargin;
		m_iLabelHeight = iLabelHeight;
		m_iButtonHeight = iButtonHeight;
	}
} // namespace controls