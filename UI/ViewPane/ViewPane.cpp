#include <StdAfx.h>
#include <UI/ViewPane/ViewPane.h>
#include <UI/UIFunctions.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>

namespace viewpane
{
	// Draw our header, if needed.
	// Draws everything to GetLabelHeight()
	void ViewPane::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		const _In_ int x,
		const _In_ int y,
		const _In_ int width,
		const _In_ int /*height*/)
	{
		const auto labelHeight = GetLabelHeight();
		output::DebugPrint(
			output::dbgLevel::Draw,
			L"ViewPane::DeferWindowPos x:%d width:%d \n",
			x,
			width);

		m_Header.DeferWindowPos(hWinPosInfo, x, y, width, labelHeight);
	}

	void ViewPane::Initialize(_In_ CWnd* pParent, _In_opt_ HDC hdc)
	{
		if (pParent) m_hWndParent = pParent->m_hWnd;
		// We compute nID for our view, the label, and collapse button all from the pane's base ID.
		const UINT iCurIDLabel = IDC_PROP_CONTROL_ID_BASE + 2 * m_paneID;
		m_nID = IDC_PROP_CONTROL_ID_BASE + 2 * m_paneID + 1;
		// Assign a nID to the collapse button that is IDD_COLLAPSE more than the control's nID
		m_nIDCollapse = m_nID + IDD_COLLAPSE;

		m_Header.Initialize(pParent, hdc, m_bCollapsible, m_nID);
	}

	ULONG ViewPane::HandleChange(UINT nID)
	{
		if (m_Header.HandleChange(nID))
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
		m_iSmallHeightMargin = iSmallHeightMargin;
		m_iLargeHeightMargin = iLargeHeightMargin;
		m_iButtonHeight = iButtonHeight;
		m_iEditHeight = iEditHeight;

		m_Header.SetMargins(
			iMargin, iSideMargin, iLabelHeight, iSmallHeightMargin, iLargeHeightMargin, iButtonHeight, iEditHeight);
	}

	void ViewPane::SetAddInLabel(const std::wstring& szLabel) { m_Header.SetLabel(szLabel); }

	void ViewPane::UpdateButtons() {}
} // namespace viewpane