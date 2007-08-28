#pragma once
// RecipientsDlg.h : header file
//

//forward definitions
class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "ContentsTableDlg.h"

/////////////////////////////////////////////////////////////////////////////
// CRecipientsDlg dialog

class CRecipientsDlg : public CContentsTableDlg
{
	// Construction
public:
	CRecipientsDlg(
		CParentWnd* pParentWnd,
		CMapiObjects *lpMapiObjects,
		LPMAPITABLE	lpMAPITable,
		LPMESSAGE lpMessage);
	virtual ~CRecipientsDlg();

//	virtual HRESULT	OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, LPMAPIPROP* lppMAPIProp);

	// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(CRecipientsDlg)
	afx_msg void OnInitMenu(CMenu* pMenu);
	afx_msg void OnDeleteSelectedItem();
	afx_msg void OnModifyRecipients();
	afx_msg void OnRecipOptions();
	afx_msg void OnSaveChanges();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

	LPMESSAGE		m_lpMessage;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
