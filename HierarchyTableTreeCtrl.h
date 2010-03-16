#pragma once
// HierarchyTableTreeCtrl.h : header file

class CMapiObjects;
class CHierarchyTableDlg;

#include "enums.h"

class CHierarchyTableTreeCtrl : public CTreeCtrl
{
public:
	CHierarchyTableTreeCtrl(
		CWnd* pCreateParent,
		CMapiObjects* lpMapiObjects,
		CHierarchyTableDlg *lpHostDlg,
		ULONG ulDisplayFlags,
		UINT nIDContextMenu);
	virtual ~CHierarchyTableTreeCtrl();

	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	// Initialization
	HRESULT LoadHierarchyTable(LPMAPICONTAINER lpMAPIContainer);
	HRESULT RefreshHierarchyTable();

	// Selected item accessors
	LPMAPICONTAINER GetSelectedContainer(__mfcmapiModifyEnum bModify);
	LPSBinary       GetSelectedItemEID();
	SortListData*   GetSelectedItemData();
	BOOL            IsItemSelected();

private:
	// Overrides from base class
	LRESULT     WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	void        OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);

	HRESULT     ExpandNode(HTREEITEM Parent);
	HTREEITEM   FindNode(LPSBinary lpInstance, HTREEITEM hParent);
	void        GetContainer(HTREEITEM Item, __mfcmapiModifyEnum bModify, LPMAPICONTAINER *lpContainer);
	LPMAPITABLE GetHierarchyTable(HTREEITEM hItem,LPMAPICONTAINER lpMAPIContainer,BOOL bGetTable);
	void        OnContextMenu(CWnd *pWnd, CPoint pos);
	void        OnDblclk(NMHDR* pNMHDR, LRESULT* pResult);
	void        OnDeleteItem(NMHDR* pNMHDR, LRESULT* pResult);
	void        OnEndLabelEdit(NMHDR* pNMHDR, LRESULT* pResult);
	void        OnGetDispInfo(NMHDR* pNMHDR, LRESULT* pResult);
	UINT        OnGetDlgCode();
	void        OnItemExpanding(NMHDR* pNMHDR, LRESULT* pResult);
	void        OnRightClick(NMHDR* pNMHDR, LRESULT* pResult);
	void        OnSelChanged(NMHDR* pNMHDR, LRESULT* pResult);
	void        UpdateSelectionUI(HTREEITEM hItem);

	// Node insertion
	HRESULT AddRootNode(LPMAPICONTAINER lpMAPIContainer);
	void AddNode(
		ULONG			cProps,
		LPSPropValue	lpProps,
		LPCTSTR szName,
		LPSBinary	lpEntryID,
		LPSBinary	lpInstanceKey,
		ULONG		bSubfolders,
		ULONG		ulContainerFlags,
		HTREEITEM	hParent,
		BOOL		bDoNotifs);
	void AddNode(LPSRow lpsRow, HTREEITEM hParent, BOOL bGetTable);

	// Custom messages
	LRESULT msgOnAddItem(WPARAM wParam, LPARAM lParam);
	LRESULT msgOnDeleteItem(WPARAM wParam, LPARAM lParam);
	LRESULT msgOnModifyItem(WPARAM wParam, LPARAM lParam);
	LRESULT msgOnRefreshTable(WPARAM wParam, LPARAM lParam);

	LONG				m_cRef;
	CHierarchyTableDlg*	m_lpHostDlg;
	CMapiObjects*		m_lpMapiObjects;
	LPMAPICONTAINER		m_lpContainer;
	ULONG				m_ulContainerType;
	ULONG				m_ulDisplayFlags;
	UINT				m_nIDContextMenu;
	BOOL				m_bItemSelected;

	DECLARE_MESSAGE_MAP()
};