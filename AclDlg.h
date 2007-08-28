#pragma once
// AclDlg.h : header file
//

//forward definitions
class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "ContentsTableDlg.h"

/////////////////////////////////////////////////////////////////////////////
// CAclDlg dialog

class CAclDlg : public CContentsTableDlg
{
	// Construction
public:
	CAclDlg(
		CParentWnd* pParentWnd,
		CMapiObjects *lpMapiObjects,
		LPEXCHANGEMODIFYTABLE lpExchTbl,
		BOOL bFreeBusyVisible);
	virtual ~CAclDlg();

	virtual HRESULT	OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, LPMAPIPROP* lppMAPIProp);

	virtual void OnRefreshView();

	// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(CAclDlg)
	afx_msg void OnInitMenu(CMenu* pMenu);
	afx_msg void OnDeleteSelectedItem();
	afx_msg void OnModifySelectedItem();
	afx_msg void OnAddItem();

	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()

	HRESULT GetSelectedItems(ULONG ulFlags, ULONG ulRowFlags, LPROWLIST* lppRowList);
	virtual void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);
	LPEXCHANGEMODIFYTABLE m_lpExchTbl;

private:
	ULONG m_ulTableFlags;
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
