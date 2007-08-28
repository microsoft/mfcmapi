#pragma once
// PublicFolderTableDlg.h : header file
//

//forward definitions
class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "ContentsTableDlg.h"

/////////////////////////////////////////////////////////////////////////////
// CPublicFolderTableDlg dialog

class CPublicFolderTableDlg : public CContentsTableDlg
{
	// Construction
public:
	CPublicFolderTableDlg(
		CParentWnd* pParentWnd,
		CMapiObjects *lpMapiObjects,
		LPCTSTR lpszServerName,
		LPMAPITABLE	lpMAPITable);
	virtual ~CPublicFolderTableDlg();

	virtual HRESULT	OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, LPMAPIPROP* lppMAPIProp);

	// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(CPublicFolderTableDlg)
	afx_msg void OnDisplayItem();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
	BOOL CreateDialogAndMenu(UINT nIDMenuResource);
	virtual void OnCreatePropertyStringRestriction();

	LPTSTR m_lpszServerName;
};

