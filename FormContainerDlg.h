#pragma once
// FormContainerDlg.h : header file
//

class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "ContentsTableDlg.h"

/////////////////////////////////////////////////////////////////////////////
// CFormContainerDlg dialog

class CFormContainerDlg : public CContentsTableDlg
{
	// Construction
public:

	CFormContainerDlg(
		CParentWnd* pParentWnd,
		CMapiObjects *lpMapiObjects,
		LPMAPIFORMCONTAINER lpFormContainer);
	virtual ~CFormContainerDlg();

	virtual void	OnRefreshView();
	virtual HRESULT	OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, LPMAPIPROP* lppMAPIProp);

	// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(CFormContainerDlg)
	afx_msg void OnInitMenu(CMenu* pMenu);
	afx_msg void OnDeleteSelectedItem();
	afx_msg void OnInstallForm();
	afx_msg void OnRemoveForm();
	afx_msg void OnResolveMessageClass();
	afx_msg void OnResolveMultipleMessageClasses();
	afx_msg void OnCalcFormPropSet();
	afx_msg void OnGetDisplay();
	//}}AFX_MSG
	// Generated message map functions
	DECLARE_MESSAGE_MAP()

	LPMAPIFORMCONTAINER m_lpFormContainer;

private:
	virtual BOOL	OnInitDialog();
	virtual void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
