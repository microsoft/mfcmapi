#pragma once
#include <UI/ViewPane/ViewPane.h>
#include <UI/Controls/FakeSplitter.h>
// List splitter control which hosts two controls either horizontally or vertically

namespace dialog
{
	class CBaseDialog;
}

namespace controls
{
	// Implementation of a lite Splitter class.
	// Liberal code sharing from the CSplitterWnd class
	class CFakeSplitter2 : public CWnd
	{
	public:
		CFakeSplitter2() = default;
		~CFakeSplitter2();

		void Init(HWND hWnd);
		void SetPaneOne(HWND /*paneOne*/) noexcept {}
		void SetPaneTwo(HWND /*paneTwo*/) noexcept {}
		void SetPaneOne(std::shared_ptr<viewpane::ViewPane> paneOne) noexcept { m_ViewPaneOne = paneOne; }
		void SetPaneTwo(std::shared_ptr<viewpane::ViewPane> paneTwo) noexcept {}

		void SetRightLabel(const std::wstring szLabel);
		void SetPercent(FLOAT /*iNewPercent*/) {}
		void SetSplitType(splitType /*stSplitType*/) noexcept {}
		void OnSize(UINT nType, int cx, int cy);
		HDWP DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height);
		int GetSplitWidth() const noexcept {}

		// Callbacks
		std::function<int()> PaneOneMinSpanCallback = nullptr;
		std::function<int()> PaneTwoMinSpanCallback = nullptr;

	private:
		void OnPaint();
		int OnCreate(LPCREATESTRUCT lpCreateStruct);
		LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;

		CEdit m_rightLabel; // Label on right of header
		int m_rightLabelWidth{0}; // The width of the string

		HWND m_PaneOne{};
		HWND m_hwndParent{};
		std::shared_ptr<viewpane::ViewPane> m_ViewPaneOne{};

		DECLARE_MESSAGE_MAP()
	};
} // namespace controls