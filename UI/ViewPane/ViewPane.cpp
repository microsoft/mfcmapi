#include <StdAfx.h>
#include <UI/ViewPane/ViewPane.h>
#include <UI/UIFunctions.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>

namespace viewpane
{
	// Draw our collapse button and label, if needed.
	// Draws everything to GetLabelHeight()
	void ViewPane::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		const _In_ int x,
		const _In_ int y,
		const _In_ int width,
		const _In_ int /*height*/)
	{
		const auto labelHeight = GetLabelHeight();
		auto curX = x;
		if (m_bCollapsible)
		{
			StyleButton(m_CollapseButton.m_hWnd, m_bCollapsed ? ui::bsUpArrow : ui::bsDownArrow);
			::DeferWindowPos(
				hWinPosInfo, m_CollapseButton.GetSafeHwnd(), nullptr, curX, y, width, labelHeight, SWP_NOZORDER);
			curX += m_iButtonHeight;
		}

		output::DebugPrint(
			DBGDraw,
			L"ViewPane::DeferWindowPos x:%d width:%d labelpos:%d labelwidth:%d \n",
			x,
			width,
			curX,
			m_iLabelWidth);

		::DeferWindowPos(
			hWinPosInfo, m_Label.GetSafeHwnd(), nullptr, curX, y, m_iLabelWidth, labelHeight, SWP_NOZORDER);
	}

	void ViewPane::Initialize(_In_ CWnd* pParent, _In_opt_ HDC hdc)
	{
		if (pParent) m_hWndParent = pParent->m_hWnd;
		// We compute nID for our view, the label, and collapse button all from the pane's base ID.
		const UINT iCurIDLabel = IDC_PROP_CONTROL_ID_BASE + 2 * m_paneID;
		m_nID = IDC_PROP_CONTROL_ID_BASE + 2 * m_paneID + 1;

		EC_B_S(m_Label.Create(
			WS_CHILD | WS_CLIPSIBLINGS | ES_READONLY | WS_VISIBLE, CRect(0, 0, 0, 0), pParent, iCurIDLabel));
		SetWindowTextW(m_Label.m_hWnd, m_szLabel.c_str());
		ui::SubclassLabel(m_Label.m_hWnd);

		if (m_bCollapsible)
		{
			StyleLabel(m_Label.m_hWnd, ui::lsPaneHeaderLabel);

			// Assign a nID to the collapse button that is IDD_COLLAPSE more than the control's nID
			EC_B_S(m_CollapseButton.Create(
				NULL,
				WS_TABSTOP | WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE,
				CRect(0, 0, 0, 0),
				pParent,
				IDD_COLLAPSE + m_nID));
		}

		const auto sizeText = ui::GetTextExtentPoint32(hdc, m_szLabel);
		m_iLabelWidth = sizeText.cx;
		output::DebugPrint(
			DBGDraw, L"ViewPane::Initialize m_iLabelWidth:%d \"%ws\"\n", m_iLabelWidth, m_szLabel.c_str());
	}

	ULONG ViewPane::HandleChange(UINT nID)
	{
		// Collapse buttons have a nID IDD_COLLAPSE higher than nID of the pane they toggle.
		// So if we get asked about one that matches, we can assume it's time to toggle our collapse.
		if (IDD_COLLAPSE + m_nID == nID)
		{
			OnToggleCollapse();
			return m_paneID;
		}

		return static_cast<ULONG>(-1);
	}

	void ViewPane::OnToggleCollapse()
	{
		m_bCollapsed = !m_bCollapsed;

		// Trigger a redraw
		::PostMessage(m_hWndParent, WM_COMMAND, IDD_RECALCLAYOUT, NULL);
	}

	void ViewPane::SetMargins(
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

	void ViewPane::SetAddInLabel(const std::wstring& szLabel) { m_szLabel = szLabel; }

	void ViewPane::UpdateButtons() {}
} // namespace viewpane