#include <StdAfx.h>
#include <UI/ViewPane/SplitterPane.h>

namespace viewpane
{
	SplitterPane* SplitterPane::CreateVerticalPane(const int paneID, const UINT uidLabel)
	{
		const auto pane = CreateHorizontalPane(paneID, uidLabel);
		pane->m_bVertical = true;

		return pane;
	}

	SplitterPane* SplitterPane::CreateHorizontalPane(const int paneID, const UINT uidLabel)
	{
		const auto pane = new (std::nothrow) SplitterPane();
		pane->SetLabel(uidLabel);
		if (uidLabel)
		{
			pane->m_bCollapsible = true;
		}

		pane->m_paneID = paneID;
		return pane;
	}

	int SplitterPane::GetMinWidth(_In_ HDC hdc)
	{
		if (m_bVertical)
		{
			return max(m_PaneOne->GetMinWidth(hdc), m_PaneTwo->GetMinWidth(hdc));
		}
		else
		{
			return m_PaneOne->GetMinWidth(hdc) + m_PaneTwo->GetMinWidth(hdc) + m_lpSplitter
					   ? m_lpSplitter->GetSplitWidth()
					   : 0;
		}
	}

	int SplitterPane::GetFixedHeight()
	{
		if (m_bVertical)
		{
			return m_PaneOne->GetFixedHeight() + m_PaneTwo->GetFixedHeight() + m_lpSplitter
					   ? m_lpSplitter->GetSplitWidth()
					   : 0;
		}
		else
		{
			return max(m_PaneOne->GetFixedHeight(), m_PaneTwo->GetFixedHeight());
		}
	}

	int SplitterPane::GetLines()
	{
		if (m_bVertical)
		{
			return m_PaneOne->GetLines() + m_PaneTwo->GetLines();
		}
		else
		{
			return max(m_PaneOne->GetLines(), m_PaneTwo->GetLines());
		}
	}

	ULONG SplitterPane::HandleChange(const UINT nID)
	{
		// See if the panes can handle the change first
		auto paneID = m_PaneOne->HandleChange(nID);
		if (paneID != static_cast<ULONG>(-1)) return paneID;

		paneID = m_PaneTwo->HandleChange(nID);
		if (paneID != static_cast<ULONG>(-1)) return paneID;

		return ViewPane::HandleChange(nID);
	}

	void SplitterPane::SetMargins(
		const int iMargin,
		const int iSideMargin,
		const int iLabelHeight, // Height of the label
		const int iSmallHeightMargin,
		const int iLargeHeightMargin,
		const int iButtonHeight, // Height of buttons below the control
		const int iEditHeight) // height of an edit control
	{
		if (m_PaneOne)
		{
			m_PaneOne->SetMargins(
				iMargin, iSideMargin, iLabelHeight, iSmallHeightMargin, iLargeHeightMargin, iButtonHeight, iEditHeight);
		}

		if (m_PaneTwo)
		{
			m_PaneTwo->SetMargins(
				iMargin, iSideMargin, iLabelHeight, iSmallHeightMargin, iLargeHeightMargin, iButtonHeight, iEditHeight);
		}
	}

	void SplitterPane::Initialize(_In_ CWnd* pParent, _In_ HDC hdc)
	{
		ViewPane::Initialize(pParent, nullptr);

		m_lpSplitter = new controls::CFakeSplitter();

		if (m_lpSplitter)
		{
			m_lpSplitter->Init(pParent->GetSafeHwnd());
			m_lpSplitter->SetSplitType(m_bVertical ? controls::SplitVertical : controls::SplitHorizontal);
			m_PaneOne->Initialize(m_lpSplitter, hdc);
			m_PaneTwo->Initialize(m_lpSplitter, hdc);
			m_lpSplitter->SetPaneOne(m_PaneOne);
			m_lpSplitter->SetPaneTwo(m_PaneTwo);
		}

		m_bInitialized = true;
	}

	void SplitterPane::DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		_In_ const int x,
		_In_ const int y,
		_In_ const int width,
		_In_ const int height)
	{
		output::DebugPrint(
			DBGDraw, L"SplitterPane::DeferWindowPos x:%d y:%d width:%d height: %d\n", x, y, width, height);
		::DeferWindowPos(hWinPosInfo, m_lpSplitter->GetSafeHwnd(), nullptr, x, y, width, height, SWP_NOZORDER);
		m_lpSplitter->OnSize(NULL, width, height);
	}
} // namespace viewpane