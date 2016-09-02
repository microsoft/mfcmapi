#pragma once
#include "enums.h"
#include "SortList\SortListData.h"

class CMapiObjects;
class CHierarchyTableDlg;

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
	_Check_return_ LPSBinary GetSelectedItemEID() const;
	_Check_return_ SortListData* GetSelectedItemData() const;
	_Check_return_ bool IsItemSelected() const;

private:
	// Overrides from base class
	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;
	void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);

	_Check_return_ HRESULT ExpandNode(HTREEITEM Parent);
	_Check_return_ HTREEITEM FindNode(_In_ LPSBinary lpInstance, HTREEITEM hParent) const;
	void GetContainer(HTREEITEM Item, __mfcmapiModifyEnum bModify, _In_ LPMAPICONTAINER *lpContainer) const;
	_Check_return_ LPMAPITABLE GetHierarchyTable(HTREEITEM hItem, _In_opt_ LPMAPICONTAINER lpMAPIContainer, bool bGetTable);
	void OnContextMenu(_In_ CWnd *pWnd, CPoint pos);
	void OnCustomDraw(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
	void OnDblclk(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
	void OnDeleteItem(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
	void OnEndLabelEdit(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
	void OnGetDispInfo(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
	_Check_return_ UINT OnGetDlgCode();
	void OnItemExpanding(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
	void OnRightClick(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
	void OnSelChanged(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
	void UpdateSelectionUI(HTREEITEM hItem);

	// Node insertion
	_Check_return_ HRESULT AddRootNode(_In_ LPMAPICONTAINER lpMAPIContainer);
	void AddNode(
		_In_ wstring szName,
		HTREEITEM hParent,
		SortListData* lpData,
		bool bGetTable);
	void AddNode(_In_ LPSRow lpsRow, HTREEITEM hParent, bool bGetTable);

	// Custom messages
	_Check_return_ LRESULT msgOnAddItem(WPARAM wParam, LPARAM lParam);
	_Check_return_ LRESULT msgOnDeleteItem(WPARAM wParam, LPARAM lParam);
	_Check_return_ LRESULT msgOnModifyItem(WPARAM wParam, LPARAM lParam);
	_Check_return_ LRESULT msgOnRefreshTable(WPARAM wParam, LPARAM lParam);

	LONG m_cRef;
	CHierarchyTableDlg* m_lpHostDlg;
	CMapiObjects* m_lpMapiObjects;
	LPMAPICONTAINER m_lpContainer;
	ULONG m_ulContainerType;
	ULONG m_ulDisplayFlags;
	UINT m_nIDContextMenu;
	bool m_bItemSelected;
	bool m_bShuttingDown;
	HTREEITEM m_hItemCurHover;
	bool m_HoverButton;

	DECLARE_MESSAGE_MAP()
};