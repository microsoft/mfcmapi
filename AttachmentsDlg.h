#pragma once
// AttachmentsDlg.h : header file
//

//forward definitions
class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "ContentsTableDlg.h"

/////////////////////////////////////////////////////////////////////////////
// CAttachmentsDlg dialog

class CAttachmentsDlg : public CContentsTableDlg
{
	// Construction
public:
	CAttachmentsDlg(
		CParentWnd* pParentWnd,
		CMapiObjects *lpMapiObjects,
		LPMAPITABLE	lpMAPITable,
		LPMESSAGE lpMessage);
	virtual ~CAttachmentsDlg();

	virtual HRESULT	OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, LPMAPIPROP* lppMAPIProp);

	// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(CAttachmentsDlg)
	afx_msg void OnInitMenu(CMenu* pMenu);
	afx_msg void OnDeleteSelectedItem();
	afx_msg void OnModifySelectedItem();
	afx_msg void OnSaveChanges();
	afx_msg void OnSaveToFile();
	afx_msg void OnViewEmbeddedMessageProps();
	afx_msg void OnUseMapiModify();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

	virtual void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);

	LPATTACH		m_lpAttach;
	LPMESSAGE		m_lpMessage;
	BOOL			m_bDisplayAttachAsEmbeddedMessage;
	BOOL			m_bUseMapiModifyOnEmbeddedMessage;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
