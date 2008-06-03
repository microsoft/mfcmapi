#pragma once
// ProfileListDlg.h : header file
//

class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "ContentsTableDlg.h"

/////////////////////////////////////////////////////////////////////////////
// CProfileListDlg dialog

class CProfileListDlg : public CContentsTableDlg
{
	// Construction
public:

	CProfileListDlg(
		CParentWnd* pParentWnd,
		CMapiObjects *lpMapiObjects,
		LPMAPITABLE lpMAPITable);
	virtual ~CProfileListDlg();

	virtual void	OnRefreshView();

	// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(CProfileListDlg)
	afx_msg void OnInitMenu(CMenu* pMenu);
	afx_msg void OnGetMAPISVC();
	afx_msg void OnAddServicesToMAPISVC();
	afx_msg void OnRemoveServicesFromMAPISVC();
	afx_msg void OnAddExchangeToProfile();
	afx_msg void OnAddPSTToProfile();
	afx_msg void OnAddUnicodePSTToProfile();
	afx_msg void OnAddServiceToProfile();
	afx_msg void OnCreateProfile();
	afx_msg void OnLaunchProfileWizard();
	afx_msg void OnDeleteSelectedItem();
	afx_msg void OnDisplayItem();
	afx_msg void OnGetProfileServiceVersion();
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()

	void AddPSTToProfile(BOOL bUnicodePST);
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
