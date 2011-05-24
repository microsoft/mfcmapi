#pragma once
// HierarchyTableTreeCtrl.h : header file

class CMapiObjects;
class CHierarchyTableDlg;

#include "enums.h"

class CHierarchyTableTreeCtrl : public CTreeCtrl
{
public:
	CHierarchyTableTreeCtrl(
		_In_ CWnd* pCreateParent,
		_In_ CMapiObjects* lpMapiObjects,
		_In_ CHierarchyTableDlg *lpHostDlg,
		ULONG ulDisplayFlags,
		UINT nIDContextMenu);
	virtual ~CHierarchyTableTreeCtrl();

	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	// Initialization
	_Check_return_ HRESULT LoadHierarchyTable(_In_ LPMAPICONTAINER lpMAPIContainer);
	_Check_return_ HRESULT RefreshHierarchyTable();

	// Selected item accessors
	_Check_return_ LPMAPICONTAINER GetSelectedContainer(__mfcmapiModifyEnum bModify);
	_Check_return_ LPSBinary       GetSelectedItemEID();
	_Check_return_ SortListData*   GetSelectedItemData();
	_Check_return_ bool            IsItemSelected();

private:
	// Overrides from base class
	_Check_return_ LRESULT     WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	void        OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);

	_Check_return_ HRESULT     ExpandNode(HTREEITEM Parent);
	_Check_return_ HTREEITEM   FindNode(_In_ LPSBinary lpInstance, HTREEITEM hParent);
	void        GetContainer(HTREEITEM Item, __mfcmapiModifyEnum bModify, _In_ LPMAPICONTAINER *lpContainer);
	_Check_return_ LPMAPITABLE GetHierarchyTable(HTREEITEM hItem, _In_opt_ LPMAPICONTAINER lpMAPIContainer, bool bGetTable);
	void        OnContextMenu(_In_ CWnd *pWnd, CPoint pos);
	void        OnDblclk(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
	void        OnDeleteItem(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
	void        OnEndLabelEdit(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
	void        OnGetDispInfo(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
	_Check_return_ UINT        OnGetDlgCode();
	void        OnItemExpanding(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
	void        OnRightClick(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
	void        OnSelChanged(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
	void        UpdateSelectionUI(HTREEITEM hItem);

	// Node insertion
	_Check_return_ HRESULT AddRootNode(_In_ LPMAPICONTAINER lpMAPIContainer);
	void AddNode(
		ULONG			cProps,
		_In_opt_ LPSPropValue	lpProps,
		_In_ LPCTSTR szName,
		_In_opt_ LPSBinary	lpEntryID,
		_In_opt_ LPSBinary	lpInstanceKey,
		ULONG		bSubfolders,
		ULONG		ulContainerFlags,
		HTREEITEM	hParent,
		bool		bDoNotifs);
	void AddNode(_In_ LPSRow lpsRow, HTREEITEM hParent, bool bGetTable);

	// Custom messages
	_Check_return_ LRESULT msgOnAddItem(WPARAM wParam, LPARAM lParam);
	_Check_return_ LRESULT msgOnDeleteItem(WPARAM wParam, LPARAM lParam);
	_Check_return_ LRESULT msgOnModifyItem(WPARAM wParam, LPARAM lParam);
	_Check_return_ LRESULT msgOnRefreshTable(WPARAM wParam, LPARAM lParam);

	LONG				m_cRef;
	CHierarchyTableDlg*	m_lpHostDlg;
	CMapiObjects*		m_lpMapiObjects;
	LPMAPICONTAINER		m_lpContainer;
	ULONG				m_ulContainerType;
	ULONG				m_ulDisplayFlags;
	UINT				m_nIDContextMenu;
	bool				m_bItemSelected;

	DECLARE_MESSAGE_MAP()
};