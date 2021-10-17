#include <StdAfx.h>
#include <UI/Controls/ViewHeader/ViewHeader.h>
#include <UI/UIFunctions.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>

namespace controls
{
	// Draw our collapse button and label, if needed.
	// Draws everything to GetLabelHeight()
	void ViewHeader::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		const _In_ int x,
		const _In_ int y,
		const _In_ int width,
		const _In_ int /*height*/)
	{
		const auto labelHeight = GetFixedHeight();
		auto curX = x;
		if (m_bCollapsible)
		{
			StyleButton(
				m_CollapseButton.m_hWnd, m_bCollapsed ? ui::uiButtonStyle::UpArrow : ui::uiButtonStyle::DownArrow);
			::DeferWindowPos(
				hWinPosInfo, m_CollapseButton.GetSafeHwnd(), nullptr, curX, y, width, labelHeight, SWP_NOZORDER);
			curX += m_iButtonHeight;
		}

		output::DebugPrint(
			output::dbgLevel::Draw,
			L"ViewHeader::DeferWindowPos x:%d width:%d labelpos:%d labelwidth:%d \n",
			x,
			width,
			curX,
			m_iLabelWidth);

		::DeferWindowPos(hWinPosInfo, GetSafeHwnd(), nullptr, curX, y, m_iLabelWidth, labelHeight, SWP_NOZORDER);
	}

	void ViewHeader::Initialize(_In_ CWnd* pParent, _In_opt_ HDC hdc, _In_ bool bCollapsible, _In_ UINT nidParent)
	{
		m_bCollapsible = bCollapsible;
		if (pParent) m_hWndParent = pParent->m_hWnd;
		// We compute nID for our view, the label, and collapse button all from the pane's base ID.
		const UINT iCurIDLabel = IDC_PROP_CONTROL_ID_BASE + 2 * m_paneID;
		// Assign a nID to the collapse button that is IDD_COLLAPSE more than the control's nID
		m_nIDCollapse = nidParent + IDD_COLLAPSE;

		EC_B_S(Create(WS_CHILD | WS_CLIPSIBLINGS | ES_READONLY | WS_VISIBLE, CRect(0, 0, 0, 0), pParent, iCurIDLabel));
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
		}

		const auto sizeText = ui::GetTextExtentPoint32(hdc, m_szLabel);
		m_iLabelWidth = sizeText.cx;
		output::DebugPrint(
			output::dbgLevel::Draw,
			L"ViewHeader::Initialize m_iLabelWidth:%d \"%ws\"\n",
			m_iLabelWidth,
			m_szLabel.c_str());
	}

	bool ViewHeader::HandleChange(UINT nID)
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

	void ViewHeader::OnToggleCollapse()
	{
		m_bCollapsed = !m_bCollapsed;

		// Trigger a redraw
		::PostMessage(m_hWndParent, WM_COMMAND, IDD_RECALCLAYOUT, NULL);
	}

	void ViewHeader::SetMargins(
		int iMargin,
		int iSideMargin,
		int iLabelHeight, // Height of the label
		int iSmallHeightMargin,
		int iLargeHeightMargin,
		int iButtonHeight, // Height of buttons below the control
		int iEditHeight) // height of an edit control
	{
		m_iMargin = iMargin;
		m_iSideMargin = iSideMargin;
		m_iLabelHeight = iLabelHeight;
		m_iSmallHeightMargin = iSmallHeightMargin;
		m_iLargeHeightMargin = iLargeHeightMargin;
		m_iButtonHeight = iButtonHeight;
		m_iEditHeight = iEditHeight;
	}

	void ViewHeader::SetAddInLabel(const std::wstring& szLabel) { m_szLabel = szLabel; }

	void ViewHeader::UpdateButtons() {}
} // namespace controls