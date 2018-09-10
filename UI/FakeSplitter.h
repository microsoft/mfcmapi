#pragma once
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
		CFakeSplitter(HWND hWnd);
		virtual ~CFakeSplitter();

		void SetPaneOne(CWnd* PaneOne);
		void SetPaneTwo(CWnd* PaneTwo);
		void SetPercent(FLOAT iNewPercent);
		void SetSplitType(SplitType stSplitType);

	private:
		void OnSize(UINT nType, int cx, int cy);
		void OnPaint();
		void OnMouseMove(UINT nFlags, CPoint point);
		LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
		_Check_return_ int HitTest(LONG x, LONG y) const;

		// starting and stopping tracking
		void StartTracking(int ht);
		void StopTracking();
		void CalcSplitPos();

		bool m_bTracking{};
		FLOAT m_flSplitPercent{0.5};
		CWnd* m_PaneOne{};
		CWnd* m_PaneTwo{};
		int m_iSplitWidth;
		int m_iSplitPos{1};
		SplitType m_SplitType{SplitHorizontal};
		HCURSOR m_hSplitCursorV{};
		HCURSOR m_hSplitCursorH{};

		DECLARE_MESSAGE_MAP()
	};
} // namespace controls