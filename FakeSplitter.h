#pragma once
// FakeSplitter.h : header file
//

//forward definitions
class CBaseDialog;

/////////////////////////////////////////////////////////////////////////////
// FakeSplitter window

enum SplitType
{
	SplitVertical   = 0,
	SplitHorizontal = 1
};

//Implementation of a lite Splitter class.
//Liberal code sharing from the CSplitterWnd class
class CFakeSplitter : public CWnd
{
public:
	CFakeSplitter(
		CBaseDialog *lpHostDlg);
	~CFakeSplitter();

	BOOL SetPercent(FLOAT iNewPercent);

	void SetPaneOne(CWnd* PaneOne);
	void SetPaneTwo(CWnd* PaneTwo);

	// Generated message map functions
protected:
	//{{AFX_MSG(CFakeSplitter)
	afx_msg void OnSize(UINT nType, int cx, int cy);
	afx_msg void OnPaint();
	afx_msg void OnMouseMove(UINT nFlags,CPoint point);
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()
	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);

private:
	int HitTest(LONG x, LONG y);

	// starting and stopping tracking
	void StartTracking(int ht);
	void StopTracking();
	void CalcSplitPos();

	CBaseDialog*		m_lpHostDlg;
	BOOL				m_bTracking;
	FLOAT				m_flSplitPercent;
	CWnd*				m_PaneOne;
	CWnd*				m_PaneTwo;
	int					m_iSplitWidth;
	int					m_iSplitPos;

public:
	SplitType			m_SplitType;
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
