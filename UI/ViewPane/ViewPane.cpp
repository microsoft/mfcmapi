#include <StdAfx.h>
#include <UI/ViewPane/ViewPane.h>
#include <UI/UIFunctions.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>

namespace viewpane
{
	// Draw our header, if needed.
	HDWP ViewPane::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		const _In_ int x,
		const _In_ int y,
		const _In_ int width,
		const _In_ int height)
	{
		output::DebugPrint(
			output::dbgLevel::Draw,
			L"ViewPane::DeferWindowPos x:%d y:%d width:%d height:%d v:%d\n",
			x,
			y,
			width,
			height,
			m_Header.IsWindowVisible());
		hWinPosInfo = ui::DeferWindowPos(
			hWinPosInfo,
			m_Header.GetSafeHwnd(),
			x,
			y,
			width,
			m_Header.GetFixedHeight(),
			L"ViewPane::DeferWindowPos::header");
		m_Header.OnSize(NULL, width, m_Header.GetFixedHeight());
		output::DebugPrint(output::dbgLevel::Draw, L"ViewPane::DeferWindowPos end\n");
		return hWinPosInfo;
	}

	void ViewPane::Initialize(_In_ CWnd* pParent, _In_opt_ HDC hdc)
	{
		if (pParent) m_hWndParent = pParent->m_hWnd;
		// We compute nID for our view and the header from the pane's base ID.
		const UINT nidHeader = IDC_PROP_CONTROL_ID_BASE + 2 * m_paneID;
		m_nID = IDC_PROP_CONTROL_ID_BASE + 2 * m_paneID + 1;
		m_Header.Initialize(pParent->GetSafeHwnd(), hdc, nidHeader);
	}

	ULONG ViewPane::HandleChange(UINT nID)
	{
		if (m_Header.HandleChange(nID))
		{
			return m_paneID;
		}

		return static_cast<ULONG>(-1);
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

		m_Header.SetMargins(iMargin, iSideMargin, iLabelHeight, iButtonHeight);
	}

	void ViewPane::UpdateButtons() {}
} // namespace viewpane