#pragma once
#include <UI/ViewPane/ViewPane.h>
#include <UI/FakeSplitter.h>

namespace viewpane
{
	class SplitterPane : public ViewPane
	{
	public:
		static SplitterPane* CreateHorizontalPane(int paneID, UINT uidLabel);
		static SplitterPane* CreateVerticalPane(int paneID, UINT uidLabel);
		~SplitterPane() { delete m_lpSplitter; }
		ViewPane* GetPaneOne() { return m_PaneOne; }
		void SetPaneOne(ViewPane* paneOne) { m_PaneOne = paneOne; }
		ViewPane* GetPaneTwo() { return m_PaneTwo; }
		void SetPaneTwo(ViewPane* paneTwo) { m_PaneTwo = paneTwo; }
		void Initialize(_In_ CWnd* pParent, _In_ HDC hdc) override;
		void DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height) override;
		int GetFixedHeight() override;
		int GetLines() override;
		void SetMargins(
			int iMargin,
			int iSideMargin,
			int iLabelHeight, // Height of the label
			int iSmallHeightMargin,
			int iLargeHeightMargin,
			int iButtonHeight, // Height of buttons below the control
			int iEditHeight) override; // height of an edit control
		void ShowWindow(const int nCmdShow)
		{
			if (m_lpSplitter) m_lpSplitter->ShowWindow(nCmdShow);
		}

	private:
		void CommitUIValues() override {}
		int GetMinWidth() override;
		ULONG HandleChange(UINT nID) override;

		controls::CFakeSplitter* m_lpSplitter{};
		ViewPane* m_PaneOne{};
		ViewPane* m_PaneTwo{};
		bool m_bVertical{};
	};
} // namespace viewpane