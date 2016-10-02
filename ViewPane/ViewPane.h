#pragma once
// ViewPane.h : header file

#include "ViewPane.h"

enum __ViewPaneFlags
{
	vpNone = 0x0000, // None
	vpDirty = 0x0001, // Pane has been edited
	vpReadonly = 0x0002, // Pane is read only
	vpCollapsible = 0x0004, // Pane can be collapsed and needs a collapse button
};

enum __ViewTypes
{
	CTRL_UNKNOWN,
	CTRL_CHECKPANE,
	CTRL_LISTPANE,
	CTRL_DROPDOWNPANE,
	CTRL_TEXTPANE,
	CTRL_COLLAPSIBLETEXTPANE,
	CTRL_COUNTEDTEXTPANE,
	CTRL_SMARTVIEWPANE,
};

class ViewPane
{
public:
	ViewPane(UINT uidLabel, bool bReadOnly);
	virtual ~ViewPane();

	virtual bool IsType(__ViewTypes vType);
	virtual void Initialize(int iControl, _In_ CWnd* pParent, _In_opt_ HDC hdc);
	virtual void SetWindowPos(int x, int y, int width, int height) = 0;
	virtual void CommitUIValues();
	virtual ULONG GetFlags();
	virtual int GetMinWidth(_In_ HDC hdc);
	virtual int GetFixedHeight() = 0;
	virtual int GetLines() = 0;
	virtual ULONG HandleChange(UINT nID);
	void OnToggleCollapse();

	virtual void SetMargins(
		int iMargin,
		int iSideMargin,
		int iLabelHeight, // Height of the label
		int iSmallHeightMargin,
		int iLargeHeightMargin,
		int iButtonHeight, // Height of buttons below the control
		int iEditHeight); // height of an edit control
	void SetAddInLabel(wstring szLabel);
	bool MatchID(UINT nID) const;

protected:
	int m_iControl; // Number of the view pane in the view - used for callbacks and layout
	bool m_bInitialized;
	bool m_bReadOnly;
	bool m_bUseLabelControl; // whether we have a label to display
	wstring m_szLabel; // Text to push into UI in Initialize
	int m_iLabelWidth; // The width of the label
	CEdit m_Label;
	UINT m_nID; // id for matching change notifications back to controls
	HWND m_hWndParent;
	bool m_bCollapsed;
	CButton m_CollapseButton;

	// Margins
	int m_iMargin;
	int m_iSideMargin;
	int m_iLabelHeight; // Height of the label
	int m_iSmallHeightMargin;
	int m_iLargeHeightMargin;
	int m_iButtonHeight; // Height of buttons below the control
	int m_iEditHeight; // height of an edit control

private:
	UINT m_uidLabel; // Label to load
};