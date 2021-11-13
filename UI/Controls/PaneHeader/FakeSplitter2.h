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
		void SetPaneOne(HWND paneOne) noexcept;
		void SetPaneTwo(HWND paneTwo) noexcept;
		void SetPaneOne(std::shared_ptr<viewpane::ViewPane> paneOne) noexcept { m_ViewPaneOne = paneOne; }
		void SetPaneTwo(std::shared_ptr<viewpane::ViewPane> paneTwo) noexcept {}

		void SetPercent(FLOAT iNewPercent);
		void SetSplitType(splitType stSplitType) noexcept;
		void OnSize(UINT nType, int cx, int cy);
		HDWP DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height);
		int GetSplitWidth() const noexcept {}

		// Callbacks
		std::function<int()> PaneOneMinSpanCallback = nullptr;
		std::function<int()> PaneTwoMinSpanCallback = nullptr;

	private:
		void OnPaint();
		LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;
		_Check_return_ int HitTest(LONG x, LONG y) const noexcept;

		HWND m_PaneOne{};
		HWND m_hwndParent{};
		std::shared_ptr<viewpane::ViewPane> m_ViewPaneOne{};
		splitType m_SplitType{splitType::horizontal};
		HCURSOR m_hSplitCursorV{};
		HCURSOR m_hSplitCursorH{};

		DECLARE_MESSAGE_MAP()
	};
} // namespace controls