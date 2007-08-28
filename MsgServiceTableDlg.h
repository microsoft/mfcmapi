#pragma once
// MsgServiceTableDlg.h : header file
//

class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "ContentsTableDlg.h"

/////////////////////////////////////////////////////////////////////////////
// CMsgServiceTableDlg dialog

class CMsgServiceTableDlg : public CContentsTableDlg
{
	// Construction
public:

	CMsgServiceTableDlg(
		CParentWnd* pParentWnd,
		CMapiObjects *lpMapiObjects,
		LPCSTR szProfileName);
	virtual ~CMsgServiceTableDlg();

	virtual HRESULT	OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, LPMAPIPROP* lppMAPIProp);

	virtual void	OnRefreshView();

	// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(CMsgServiceTableDlg)
	afx_msg void OnInitMenu(CMenu* pMenu);
	afx_msg void OnDisplayItem();
	afx_msg void OnConfigureMsgService();
	afx_msg void OnDeleteSelectedItem();
	afx_msg void OnOpenProfileSection();
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()

	virtual void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);

	LPSERVICEADMIN	m_lpServiceAdmin;
	LPSTR			m_szProfileName;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
