#include <StdAfx.h>
#include <UI/ViewPane/ViewPane.h>
#include <UI/UIFunctions.h>
#include <Interpret/String.h>

namespace viewpane
{
	ViewPane::ViewPane()
	{
		m_iControl = -1;

		m_nID = 0;
		m_bCollapsible = false;
		m_bCollapsed = false;
		m_bReadOnly = true;
		m_bInitialized = false;
		m_iMargin = 0;
		m_iSideMargin = 0;
		m_iLabelHeight = 0;
		m_iSmallHeightMargin = 0;
		m_iLargeHeightMargin = 0;
		m_iButtonHeight = 0;
		m_iEditHeight = 0;
		m_iSmallHeightMargin = 0;
		m_iLabelWidth = 0;

		m_hWndParent = nullptr;
	}

	void ViewPane::SetWindowPos(int x, int y, int width, int height)
	{
		const auto hdwp = WC_D(HDWP, BeginDeferWindowPos(2));
		if (hdwp)
		{
			DeferWindowPos(hdwp, x, y, width, height);
			EC_B_S(EndDeferWindowPos(hdwp));
		}
	}

	void ViewPane::DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int /*height*/)
	{
		if (m_bCollapsible)
		{
			StyleButton(m_CollapseButton.m_hWnd, m_bCollapsed ? ui::bsUpArrow : ui::bsDownArrow);
			::DeferWindowPos(
				hWinPosInfo, m_CollapseButton.GetSafeHwnd(), nullptr, x, y, width, m_iLabelHeight, SWP_NOZORDER);
			x += m_iButtonHeight;
		}

		output::DebugPrint(
			DBGDraw,
			L"ViewPane::DeferWindowPos x:%d width:%d labelpos:%d labelwidth:%d \n",
			x,
			width,
			x + m_iButtonHeight,
			m_iLabelWidth);

		::DeferWindowPos(
			hWinPosInfo, m_Label.GetSafeHwnd(), nullptr, x, y, m_iLabelWidth, m_iLabelHeight, SWP_NOZORDER);
	}

	void ViewPane::SetLabel(UINT uidLabel, bool bReadOnly)
	{
		m_bReadOnly = bReadOnly;

		if (uidLabel)
		{
			m_szLabel = strings::loadstring(uidLabel);
		}
	}

	void ViewPane::Initialize(int iControl, _In_ CWnd* pParent, _In_opt_ HDC /*hdc*/)
	{
		m_iControl = iControl;
		if (pParent) m_hWndParent = pParent->m_hWnd;
		const UINT iCurIDLabel = IDC_PROP_CONTROL_ID_BASE + 2 * m_iControl;
		m_nID = IDC_PROP_CONTROL_ID_BASE + 2 * m_iControl + 1;

		EC_B_S(m_Label.Create(
			WS_CHILD | WS_CLIPSIBLINGS | ES_READONLY | WS_VISIBLE, CRect(0, 0, 0, 0), pParent, iCurIDLabel));
		SetWindowTextW(m_Label.m_hWnd, m_szLabel.c_str());
		ui::SubclassLabel(m_Label.m_hWnd);

		if (m_bCollapsible)
		{
			StyleLabel(m_Label.m_hWnd, ui::lsPaneHeaderLabel);

			EC_B_S(m_CollapseButton.Create(
				NULL,
				WS_TABSTOP | WS_CHILD | WS_CLIPSIBLINGS | WS_VISIBLE,
				CRect(0, 0, 0, 0),
				pParent,
				IDD_COLLAPSE + iControl));
		}
	}

	bool ViewPane::IsDirty() { return false; }

	int ViewPane::GetMinWidth(_In_ HDC hdc)
	{
		const auto sizeText = ui::GetTextExtentPoint32(hdc, m_szLabel);
		m_iLabelWidth = sizeText.cx;
		output::DebugPrint(
			DBGDraw, L"ViewPane::GetMinWidth m_iLabelWidth:%d \"%ws\"\n", m_iLabelWidth, m_szLabel.c_str());
		return m_iLabelWidth;
	}

	int ViewPane::GetLines() { return 0; }

	ULONG ViewPane::HandleChange(UINT nID)
	{
		if (static_cast<UINT>(IDD_COLLAPSE + m_iControl) == nID)
		{
			OnToggleCollapse();
			return m_iControl;
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

	bool ViewPane::MatchID(UINT nID) const { return nID == m_nID; }

	void ViewPane::UpdateButtons() {}
} // namespace viewpane