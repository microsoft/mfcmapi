#pragma once
#include <UI/ViewPane/ViewPane.h>
#include <UI/FakeSplitter.h>

namespace viewpane
{
	class SplitterPane : public ViewPane
	{
	public:
		static SplitterPane* CreateHorizontalPane(int paneID);
		static SplitterPane* CreateVerticalPane(int paneID);
		void SetPaneOne(ViewPane* paneOne) { m_PaneOne = paneOne; }
		void SetPaneTwo(ViewPane* paneTwo) { m_PaneTwo = paneTwo; }
		// Return a pane with a matching paneID.
		ViewPane* GetPaneByID(int id) override
		{
			if (m_paneID == id) return this;
			const auto pane = m_PaneOne->GetPaneByID(id);
			if (pane) return pane;
			return m_PaneTwo->GetPaneByID(id);
		}
		// Return a pane with a matching nID.
		ViewPane* GetPaneByNID(UINT nID) override
		{
			if (m_nID == nID) return this;
			const auto pane = m_PaneOne->GetPaneByNID(nID);
			if (pane) return pane;
			return m_PaneTwo->GetPaneByNID(nID);
		}

	private:
		void Initialize(_In_ CWnd* pParent, _In_ HDC hdc) override;
		void DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height) override;
		void CommitUIValues() override {}
		int GetMinWidth(_In_ HDC hdc) override;
		int GetFixedHeight() override;
		int GetLines() override;
		ULONG HandleChange(UINT nID) override;
		void SetMargins(
			int iMargin,
			int iSideMargin,
			int iLabelHeight, // Height of the label
			int iSmallHeightMargin,
			int iLargeHeightMargin,
			int iButtonHeight, // Height of buttons below the control
			int iEditHeight) override; // height of an edit control

		controls::CFakeSplitter* m_lpSplitter{};
		ViewPane* m_PaneOne{};
		ViewPane* m_PaneTwo{};
		bool m_bVertical{};
	};
} // namespace viewpane