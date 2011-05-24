#pragma once
// FakeSplitter.h : List splitter control which hosts two controls either horizontally or vertically

class CBaseDialog;

enum SplitType
{
	SplitVertical   = 0,
	SplitHorizontal = 1
};

// Implementation of a lite Splitter class.
// Liberal code sharing from the CSplitterWnd class
class CFakeSplitter : public CWnd
{
public:
	CFakeSplitter(
		_In_ CBaseDialog* lpHostDlg);
	virtual ~CFakeSplitter();

	void SetPaneOne(CWnd* PaneOne);
	void SetPaneTwo(CWnd* PaneTwo);
	void SetPercent(FLOAT iNewPercent);
	void SetSplitType(SplitType stSplitType);

private:
	void    OnSize(UINT nType, int cx, int cy);
	void    OnPaint();
	void    OnMouseMove(UINT nFlags,CPoint point);
	_Check_return_ LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	_Check_return_ int     HitTest(LONG x, LONG y);

	// starting and stopping tracking
	void StartTracking(int ht);
	void StopTracking();
	void CalcSplitPos();

	CBaseDialog*	m_lpHostDlg;
	bool			m_bTracking;
	FLOAT			m_flSplitPercent;
	CWnd*			m_PaneOne;
	CWnd*			m_PaneTwo;
	int				m_iSplitWidth;
	int				m_iSplitPos;
	SplitType		m_SplitType;

	DECLARE_MESSAGE_MAP()
};