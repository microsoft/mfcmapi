#pragma once
// HierarchyTableDlg.h : header file
//

class CHierarchyTableTreeCtrl;
class CParentWnd;
class CMapiObjects;

#include "BaseDialog.h"

/////////////////////////////////////////////////////////////////////////////
// CHierarchyTableDlg dialog

class CHierarchyTableDlg : public CBaseDialog
{
friend class CHierarchyTableTreeCtrl;
	// Construction
public:
	CHierarchyTableDlg(
		CParentWnd* pParentWnd,
		CMapiObjects *lpMapiObjects,
		UINT uidTitle,
		LPUNKNOWN lpRootContainer,
		ULONG nIDContextMenu,
		ULONG ulAddInContext
		);
	virtual ~CHierarchyTableDlg();

	BOOL CreateDialogAndMenu(UINT nIDMenuResource);

	// Implementation
protected:
	__mfcmapiDeletedItemsEnum	m_bShowingDeletedFolders;
	CHierarchyTableTreeCtrl*	m_lpHierarchyTableTreeCtrl;

	virtual BOOL	OnInitDialog();
	virtual void	OnInitMenu(CMenu* pMenu);
	BOOL PreTranslateMessage(MSG* pMsg);

	void	OnCancel();

	// Generated message map functions
	//{{AFX_MSG(CHierarchyTableDlg)
	afx_msg void OnDisplayItem();
	afx_msg void OnRefreshView();
	afx_msg void OnDisplayHierarchyTable();
	afx_msg void OnEditSearchCriteria();
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()

	virtual BOOL HandleAddInMenu(WORD wMenuSelect);
	virtual void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);

private:
	UINT m_nIDContextMenu;
};
