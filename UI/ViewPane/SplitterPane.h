#pragma once
#include <UI/ViewPane/ViewPane.h>
#include <UI/FakeSplitter.h>

namespace viewpane
{
	class SplitterPane : public ViewPane
	{
	public:
		SplitterPane() : m_bVertical(false) {}

		static SplitterPane* CreateHorizontalPane(int paneID);
		static SplitterPane* CreateVerticalPane(int paneID);
		void SetPaneOne(ViewPane* paneOne)
		{
			m_PaneOneControl = paneOne->GetID();
			m_PaneOne = paneOne;
		}
		void SetPaneTwo(ViewPane* paneTwo)
		{
			m_PaneTwoControl = paneTwo->GetID();
			m_PaneTwo = paneTwo;
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
		int m_PaneOneControl{};
		ViewPane* m_PaneOne{};
		int m_PaneTwoControl{};
		ViewPane* m_PaneTwo{};
		bool m_bVertical{};
	};
} // namespace viewpane