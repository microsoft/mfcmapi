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
		(void) ViewPane::GetMinWidth(hdc);
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
		auto iHeight = 0;
		// TODO: Better way to find the top pane
		if (0 != m_paneID) iHeight += m_iSmallHeightMargin; // Top margin

		if (m_bCollapsible)
		{
			// Our expand/collapse button
			iHeight += m_iButtonHeight;
		}

		// A small margin between our button and the splitter control, if we're collapsible and not collapsed
		if (!m_bCollapsed && m_bCollapsible)
		{
			iHeight += m_iSmallHeightMargin;
		}

		if (!m_bCollapsed)
		{
			if (m_bVertical)
			{
				iHeight += m_PaneOne->GetFixedHeight() + m_PaneTwo->GetFixedHeight() +
						   (m_lpSplitter ? m_lpSplitter->GetSplitWidth() : 0);
			}
			else
			{
				iHeight += max(m_PaneOne->GetFixedHeight(), m_PaneTwo->GetFixedHeight());
			}
		}

		if (m_bCollapsed)
		{
			iHeight += m_iSmallHeightMargin; // Bottom margin
		}

		return iHeight;
	}

	int SplitterPane::GetLines()
	{
		if (!m_bCollapsed)
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

		return 0;
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
		ViewPane::SetMargins(
			iMargin, iSideMargin, iLabelHeight, iSmallHeightMargin, iLargeHeightMargin, iButtonHeight, iEditHeight);

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
		_In_ int y,
		_In_ const int width,
		_In_ int height)
	{
		output::DebugPrint(
			DBGDraw, L"SplitterPane::DeferWindowPos x:%d y:%d width:%d height: %d\n", x, y, width, height);

		if (0 != m_paneID)
		{
			y += m_iSmallHeightMargin;
			height -= m_iSmallHeightMargin;
		}

		ViewPane::DeferWindowPos(hWinPosInfo, x, y, width, height);

		if (m_bCollapsed)
		{
			EC_B_S(m_lpSplitter->ShowWindow(SW_HIDE));
		}
		else
		{
			if (m_bCollapsible)
			{
				y += m_iLabelHeight + m_iSmallHeightMargin;
				height -= m_iLabelHeight + m_iSmallHeightMargin;
			}

			EC_B_S(m_lpSplitter->ShowWindow(SW_SHOW));
			::DeferWindowPos(hWinPosInfo, m_lpSplitter->GetSafeHwnd(), nullptr, x, y, width, height, SWP_NOZORDER);
			m_lpSplitter->OnSize(NULL, width, height);
		}
	}
} // namespace viewpane