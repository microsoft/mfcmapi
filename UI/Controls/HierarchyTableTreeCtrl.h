#pragma once
#include <Enums.h>
#include <UI/Controls/SortList/SortListData.h>

namespace cache
{
	class CMapiObjects;
}

namespace dialog
{
	class CHierarchyTableDlg;
}

namespace controls
{
	class CHierarchyTableTreeCtrl : public CTreeCtrl
	{
	public:
		CHierarchyTableTreeCtrl(
			_In_ CWnd* pCreateParent,
			_In_ cache::CMapiObjects* lpMapiObjects,
			_In_ dialog::CHierarchyTableDlg *lpHostDlg,
			ULONG ulDisplayFlags,
			UINT nIDContextMenu);
		virtual ~CHierarchyTableTreeCtrl();

		STDMETHODIMP_(ULONG) AddRef();
		STDMETHODIMP_(ULONG) Release();

		// Initialization
		_Check_return_ HRESULT LoadHierarchyTable(_In_ LPMAPICONTAINER lpMAPIContainer);
		_Check_return_ HRESULT RefreshHierarchyTable();

		// Selected item accessors
		_Check_return_ LPMAPICONTAINER GetSelectedContainer(__mfcmapiModifyEnum bModify) const;
		_Check_return_ LPSBinary GetSelectedItemEID() const;
		_Check_return_ sortlistdata::SortListData* GetSelectedItemData() const;
		_Check_return_ sortlistdata::SortListData* GetSortListData(HTREEITEM iItem) const;
		_Check_return_ bool IsItemSelected() const;

	private:
		// Overrides from base class
		LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;
		void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);

		_Check_return_ HRESULT ExpandNode(HTREEITEM hParent) const;
		_Check_return_ HTREEITEM FindNode(_In_ LPSBinary lpInstance, HTREEITEM hParent) const;
		void GetContainer(HTREEITEM Item, __mfcmapiModifyEnum bModify, _In_ LPMAPICONTAINER * lppContainer) const;
		_Check_return_ LPMAPITABLE GetHierarchyTable(HTREEITEM hItem, _In_opt_ LPMAPICONTAINER lpMAPIContainer, bool bRegNotifs) const;
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
		void UpdateSelectionUI(HTREEITEM hItem) const;

		// Node insertion
		_Check_return_ HRESULT AddRootNode(_In_ LPMAPICONTAINER lpMAPIContainer) const;
		void AddNode(
			_In_ const std::wstring& szName,
			HTREEITEM hParent,
			sortlistdata::SortListData* lpData,
			bool bGetTable) const;
		void AddNode(_In_ LPSRow lpsRow, HTREEITEM hParent, bool bGetTable) const;

		// Custom messages
		_Check_return_ LRESULT msgOnAddItem(WPARAM wParam, LPARAM lParam);
		_Check_return_ LRESULT msgOnDeleteItem(WPARAM wParam, LPARAM lParam);
		_Check_return_ LRESULT msgOnModifyItem(WPARAM wParam, LPARAM lParam);
		_Check_return_ LRESULT msgOnRefreshTable(WPARAM wParam, LPARAM lParam);

		LONG m_cRef;
		dialog::CHierarchyTableDlg* m_lpHostDlg;
		cache::CMapiObjects* m_lpMapiObjects;
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
}