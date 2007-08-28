#pragma once
// HierarchyTableTreeCtrl.h : header file
//

//forward definitions
class CMapiObjects;
class CHierarchyTableDlg;

#include "enums.h"

/////////////////////////////////////////////////////////////////////////////
// CHierarchyTableTreeCtrl window

class CHierarchyTableTreeCtrl : public CTreeCtrl
{
friend class CHierarchyTableDlg;
public:
	CHierarchyTableTreeCtrl(
		CWnd* pCreateParent,
		CMapiObjects *lpMapiObjects,
		CHierarchyTableDlg *lpHostDlg,
		__mfcmapiDeletedItemsEnum bShowingDeletedFolders);
	virtual ~CHierarchyTableTreeCtrl();

	STDMETHODIMP_(ULONG)	AddRef();
 	STDMETHODIMP_(ULONG)	Release();

	BOOL	m_bItemSelected;

	LPMAPICONTAINER	GetSelectedContainer(__mfcmapiModifyEnum bModify);
	SortListData* GetSelectedItemData();
	LPSBinary	GetSelectedItemEID();
	HRESULT		RefreshHierarchyTable();
	HRESULT		LoadHierarchyTable(LPMAPICONTAINER lpMAPIContainer);

	// Generated message map functions
protected:
	//{{AFX_MSG(CHierarchyTableTreeCtrl)
	afx_msg void OnDblclk(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnEndLabelEdit(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnItemExpanding(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnDeleteItem(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnGetDispInfo(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnSelChanged(NMHDR* pNMHDR, LRESULT* pResult);
	afx_msg void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);
	afx_msg UINT OnGetDlgCode();
	afx_msg LRESULT		msgOnAddItem(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT		msgOnDeleteItem(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT		msgOnModifyItem(WPARAM wParam, LPARAM lParam);
	afx_msg LRESULT		msgOnRefreshTable(WPARAM wParam, LPARAM lParam);
	//}}AFX_MSG

	afx_msg void OnContextMenu(CWnd *pWnd, CPoint pos);
	afx_msg void OnRightClick(NMHDR* pNMHDR, LRESULT* pResult);
	DECLARE_MESSAGE_MAP()
	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
private:
	__mfcmapiDeletedItemsEnum	m_bShowingDeletedFolders;
	LONG			m_cRef;
	CHierarchyTableDlg*	m_lpHostDlg;
	CMapiObjects*	m_lpMapiObjects;
	LPMAPICONTAINER m_lpContainer;
	ULONG			m_ulContainerType;

	void	GetContainer(HTREEITEM Item, __mfcmapiModifyEnum bModify, LPMAPICONTAINER *lpContainer);
	HRESULT AddRootNode(LPMAPICONTAINER lpMAPIContainer);
	void	AddNode(
		ULONG			cProps,
		LPSPropValue	lpProps,
		LPCTSTR szName,
		LPSBinary	lpEntryID,
		LPSBinary	lpInstanceKey,
		ULONG		bSubfolders,
		ULONG		ulContainerFlags,
		HTREEITEM	hParent,
		BOOL		bDoNotifs);
	void	AddNode(LPSRow lpsRow, HTREEITEM hParent, BOOL bGetTable);
	HRESULT	ExpandNode(HTREEITEM Parent);
	void	UpdateSelectionUI(HTREEITEM hItem);
	LPMAPITABLE GetHierarchyTable(HTREEITEM hItem,LPMAPICONTAINER lpMAPIContainer,BOOL bGetTable);
	HTREEITEM FindNode(LPSBinary lpInstance, HTREEITEM hParent);
};

/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
