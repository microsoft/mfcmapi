#pragma once
#include <UI/ViewPane/ViewPane.h>
#include <UI/Controls/FakeSplitter.h>

namespace viewpane
{
	class SplitterPane : public ViewPane
	{
	public:
		static std::shared_ptr<SplitterPane> CreateHorizontalPane(int paneID, UINT uidLabel);
		static std::shared_ptr<SplitterPane> CreateVerticalPane(int paneID, UINT uidLabel);
		std::shared_ptr<ViewPane> GetPaneOne() noexcept { return m_PaneOne; }
		void SetPaneOne(std::shared_ptr<ViewPane> paneOne) { m_PaneOne = paneOne; }
		std::shared_ptr<ViewPane> GetPaneTwo() noexcept { return m_PaneTwo; }
		void SetPaneTwo(std::shared_ptr<ViewPane> paneTwo) noexcept { m_PaneTwo = paneTwo; }
		void Initialize(_In_ CWnd* pParent, _In_ HDC hdc) override;
		HDWP DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height) override;
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
		void ShowWindow(const int nCmdShow)
		{
			if (m_lpSplitter) m_lpSplitter->ShowWindow(nCmdShow);
		}
		bool containsWindow(HWND hWnd) const noexcept override;

	private:
		void CommitUIValues() override {}
		int GetMinWidth() override;

		std::shared_ptr<controls::CFakeSplitter> m_lpSplitter{};
		std::shared_ptr<ViewPane> m_PaneOne{};
		std::shared_ptr<ViewPane> m_PaneTwo{};
		bool m_bVertical{};
	};
} // namespace viewpane