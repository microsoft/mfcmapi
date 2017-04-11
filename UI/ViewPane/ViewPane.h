#pragma once

class ViewPane
{
public:
	ViewPane();
	virtual ~ViewPane() = default;

	void SetLabel(UINT uidLabel, bool bReadOnly);
	virtual void Initialize(int iControl, _In_ CWnd* pParent, _In_opt_ HDC hdc);
	virtual void SetWindowPos(int x, int y, int width, int height) = 0;
	virtual void CommitUIValues() = 0;
	virtual bool IsDirty();
	virtual int GetMinWidth(_In_ HDC hdc);
	virtual int GetFixedHeight() = 0;
	virtual int GetLines();
	virtual ULONG HandleChange(UINT nID);
	void OnToggleCollapse();

	virtual void SetMargins(
		int iMargin,
		int iSideMargin,
		int iLabelHeight, // Height of the label
		int iSmallHeightMargin,
		int iLargeHeightMargin,
		int iButtonHeight, // Height of buttons below the control
		int iEditHeight); // Height of an edit control
	void SetAddInLabel( const wstring& szLabel);
	bool MatchID(UINT nID) const;
	virtual void UpdateButtons();

protected:
	int m_iControl; // Number of the view pane in the view - used for callbacks and layout
	bool m_bInitialized;
	bool m_bReadOnly;
	wstring m_szLabel; // Text to push into UI in Initialize
	int m_iLabelWidth; // The width of the label
	CEdit m_Label;
	UINT m_nID; // Id for matching change notifications back to controls
	HWND m_hWndParent;
	bool m_bCollapsible;
	bool m_bCollapsed;
	CButton m_CollapseButton;

	// Margins
	int m_iMargin;
	int m_iSideMargin;
	int m_iLabelHeight; // Height of the label
	int m_iSmallHeightMargin;
	int m_iLargeHeightMargin;
	int m_iButtonHeight; // Height of buttons below the control
	int m_iEditHeight; // Height of an edit control
};