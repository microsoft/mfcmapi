#pragma once
// SortHeader.h : header file for our custom sort list control header

#include "InterpretProp.h"

typedef struct _HeaderData	FAR * LPHEADERDATA;

typedef struct _HeaderData
{
	ULONG	ulTagArrayRow;
	ULONG	ulPropTag;
	BOOL	bIsAB;
	TCHAR	szTipString[TAG_MAX_LEN];
} HeaderData;

class CSortHeader : public CHeaderCtrl
{
// Construction
public:
	CSortHeader();

// Attributes
public:

// Operations
public:

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CSortHeader)
	public:
	afx_msg LRESULT	msgOnSaveColumnOrder(WPARAM wParam, LPARAM lParam);
	//}}AFX_VIRTUAL

// Implementation
public:
	BOOL Init(CHeaderCtrl *pHeader, HWND hwndParent);
	virtual ~CSortHeader();

	// Generated message map functions
protected:
	DECLARE_MESSAGE_MAP()
	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	void	RegisterHeaderTooltip();
	HWND	m_hwndTip;
	TOOLINFO m_ti;
	HWND	m_hwndParent;

};