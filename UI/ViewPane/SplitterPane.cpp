#include <StdAfx.h>
#include <UI/ViewPane/SplitterPane.h>

namespace viewpane
{
	SplitterPane* SplitterPane::CreateVerticalPane(int paneID)
	{
		const auto pane = CreateHorizontalPane(paneID);
		pane->m_bVertical = true;

		return pane;
	}

	SplitterPane* SplitterPane::CreateHorizontalPane(int paneID)
	{
		const auto pane = new (std::nothrow) SplitterPane();
		pane->m_paneID = paneID;
		return pane;
	}

	int SplitterPane::GetMinWidth(_In_ HDC hdc)
	{
		if (m_bVertical)
		{
			return m_PaneOne->GetMinWidth(hdc) + m_PaneTwo->GetMinWidth(hdc) + m_lpSplitter
					   ? m_lpSplitter->GetSplitWidth()
					   : 0;
		}
		else
		{
			return max(m_PaneOne->GetMinWidth(hdc), m_PaneTwo->GetMinWidth(hdc));
		}
	}

	int SplitterPane::GetFixedHeight()
	{
		if (m_bVertical)
		{
			return max(m_PaneOne->GetFixedHeight(), m_PaneTwo->GetFixedHeight());
		}
		else
		{
			return m_PaneOne->GetFixedHeight() + m_PaneTwo->GetFixedHeight() + m_lpSplitter
					   ? m_lpSplitter->GetSplitWidth()
					   : 0;
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

	ULONG SplitterPane::HandleChange(UINT nID)
	{
		auto pane = m_PaneOne->GetPaneByNID(nID);
		if (pane) return pane->GetID();

		pane = m_PaneTwo->GetPaneByNID(nID);
		if (pane) return pane->GetID();

		return ViewPane::HandleChange(nID);
	}

	void SplitterPane::SetMargins(
		int iMargin,
		int iSideMargin,
		int iLabelHeight, // Height of the label
		int iSmallHeightMargin,
		int iLargeHeightMargin,
		int iButtonHeight, // Height of buttons below the control
		int iEditHeight) // height of an edit control
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

	void SplitterPane::DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height)
	{
		output::DebugPrint(
			DBGDraw, L"SplitterPane::DeferWindowPos x:%d y:%d width:%d height: %d\n", x, y, width, height);
		::DeferWindowPos(hWinPosInfo, m_lpSplitter->GetSafeHwnd(), nullptr, x, y, width, height, SWP_NOZORDER);
		m_lpSplitter->OnSize(NULL, width, height);
	}
} // namespace viewpane