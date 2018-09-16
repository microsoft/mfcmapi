#pragma once
#include <UI/ViewPane/ViewPane.h>
// List splitter control which hosts two controls either horizontally or vertically

namespace dialog
{
	class CBaseDialog;
}

namespace controls
{
	enum SplitType
	{
		SplitVertical = 0,
		SplitHorizontal = 1
	};

	// Implementation of a lite Splitter class.
	// Liberal code sharing from the CSplitterWnd class
	class CFakeSplitter : public CWnd
	{
	public:
		CFakeSplitter() = default;
		virtual ~CFakeSplitter();

		void Init(HWND hWnd);
		void SetPaneOne(HWND paneOne);
		void SetPaneTwo(HWND paneTwo);
		void SetPaneOne(viewpane::ViewPane* paneOne)
		{
			m_ViewPaneOne = paneOne;
			if (m_ViewPaneOne)
			{
				m_iSplitWidth = 7;
			}
			else
			{
				m_iSplitWidth = 0;
			}
		}

		void SetPaneTwo(viewpane::ViewPane* paneTwo) { m_ViewPaneTwo = paneTwo; }

		void SetPercent(FLOAT iNewPercent);
		void SetSplitType(SplitType stSplitType);
		void OnSize(UINT nType, int cx, int cy);
		int GetSplitWidth() const { return m_iSplitWidth; }

	private:
		void OnPaint();
		void OnMouseMove(UINT nFlags, CPoint point);
		LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;
		_Check_return_ int HitTest(LONG x, LONG y) const;

		// starting and stopping tracking
		void StartTracking(int ht);
		void StopTracking();
		void CalcSplitPos();

		bool m_bTracking{};
		FLOAT m_flSplitPercent{0.5};
		HWND m_PaneOne{};
		HWND m_PaneTwo{};
		HWND m_hwndParent{};
		viewpane::ViewPane* m_ViewPaneOne{};
		viewpane::ViewPane* m_ViewPaneTwo{};
		int m_iSplitWidth{};
		int m_iSplitPos{1};
		SplitType m_SplitType{SplitHorizontal};
		HCURSOR m_hSplitCursorV{};
		HCURSOR m_hSplitCursorH{};

		DECLARE_MESSAGE_MAP()
	};
} // namespace controls