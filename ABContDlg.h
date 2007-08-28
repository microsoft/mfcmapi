#pragma once
// AbContDlg.h : header file
//

//forward definitions
class CHierarchyTableTreeCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "HierarchyTableDlg.h"

/////////////////////////////////////////////////////////////////////////////
// CAbContDlg dialog

class CAbContDlg : public CHierarchyTableDlg
{
	// Construction
public:
	CAbContDlg(
		CParentWnd* pParentWnd,
		CMapiObjects *lpMapiObjects);

	~CAbContDlg();

	// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(CAbContDlg)
	afx_msg void OnSetPAB();
	//}}AFX_MSG
	// Generated message map functions
	DECLARE_MESSAGE_MAP()

	virtual void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
