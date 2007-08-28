#pragma once
// AbDlg.h : header file
//

class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "ContentsTableDlg.h"

/////////////////////////////////////////////////////////////////////////////
// CAbDlg dialog

class CAbDlg : public CContentsTableDlg
{
	// Construction
public:

	CAbDlg(
		CParentWnd* pParentWnd,
		CMapiObjects *lpMapiObjects,
		LPABCONT lpAdrBook);
	virtual ~CAbDlg();

	// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(CAbDlg)
	afx_msg void OnInitMenu(CMenu* pMenu);
	afx_msg void OnDeleteSelectedItem();
	afx_msg void OnDisplayDetails();
	afx_msg void OnOpenContact();
	afx_msg void OnOpenManager();
	afx_msg void OnOpenOwner();
	//}}AFX_MSG
	// Generated message map functions
	DECLARE_MESSAGE_MAP()
	BOOL CreateDialogAndMenu(UINT nIDMenuResource);
	virtual void OnCreatePropertyStringRestriction();
	virtual BOOL	HandleCopy();
	virtual BOOL	HandlePaste();

	virtual void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
