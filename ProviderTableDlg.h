#pragma once
// ProviderTableDlg.h : header file
//

class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "ContentsTableDlg.h"

/////////////////////////////////////////////////////////////////////////////
// CProviderTableDlg dialog

class CProviderTableDlg : public CContentsTableDlg
{
	// Construction
public:

	CProviderTableDlg(
		CParentWnd* pParentWnd,
		CMapiObjects *lpMapiObjects,
		LPMAPITABLE lpMAPITable,
		LPPROVIDERADMIN lpProviderAdmin);
	virtual ~CProviderTableDlg();

	virtual HRESULT	OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, LPMAPIPROP* lppMAPIProp);

	// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(CProviderTableDlg)
	afx_msg void OnInitMenu(CMenu* pMenu);
	afx_msg void OnDisplayItem();
	afx_msg void OnOpenProfileSection();
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()

	virtual void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);

	LPPROVIDERADMIN m_lpProviderAdmin;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
