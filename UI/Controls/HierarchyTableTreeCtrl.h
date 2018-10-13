#pragma once
#include <Enums.h>
#include <UI/Controls/SortList/SortListData.h>
#include <UI/Controls/StyleTreeCtrl.h>

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
	class CHierarchyTableTreeCtrl : public StyleTreeCtrl
	{
	public:
		virtual ~CHierarchyTableTreeCtrl();

		// Initialization
		void Create(
			_In_ CWnd* pCreateParent,
			_In_ cache::CMapiObjects* lpMapiObjects,
			_In_ dialog::CHierarchyTableDlg* lpHostDlg,
			ULONG ulDisplayFlags,
			UINT nIDContextMenu);
		void LoadHierarchyTable(_In_ LPMAPICONTAINER lpMAPIContainer);

		// Selected item accessors
		_Check_return_ LPMAPICONTAINER GetSelectedContainer(__mfcmapiModifyEnum bModify) const;
		_Check_return_ LPSBinary GetSelectedItemEID() const;
		_Check_return_ sortlistdata::SortListData* GetSelectedItemData() const;
		_Check_return_ sortlistdata::SortListData* GetSortListData(HTREEITEM iItem) const;

	private:
		// Overrides from base class
		LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;

		// Callback functions
		void OnItemAdded(HTREEITEM hItem) const;
		bool HasChildren(_In_ HTREEITEM hItem) const;
		void OnItemSelected(HTREEITEM hItem) const;
		bool HandleKeyDown(UINT nChar, bool bShiftPressed, bool bCtrlPressed, bool bMenuPressed);
		void FreeNodeData(LPARAM lpData) const;
		void ExpandNode(HTREEITEM hParent) const;
		void OnRefresh() const;
		void OnLabelEdit(HTREEITEM hItem, LPTSTR szText);
		void OnDisplaySelectedItem();
		void OnLastChildDeleted(LPARAM /*lpData*/);
		void HandleContextMenu(int x, int y);

		_Check_return_ HTREEITEM FindNode(_In_ LPSBinary lpInstance, HTREEITEM hParent) const;
		_Check_return_ LPMAPICONTAINER GetContainer(HTREEITEM Item, __mfcmapiModifyEnum bModify) const;
		void Advise(HTREEITEM hItem, sortlistdata::SortListData* lpData) const;
		_Check_return_ LPMAPITABLE
		GetHierarchyTable(HTREEITEM hItem, _In_opt_ LPMAPICONTAINER lpMAPIContainer, bool bRegNotifs) const;

		// Node management
		void AddRootNode() const;
		void AddNode(_In_ LPSRow lpsRow, HTREEITEM hParent, const std::function<void(HTREEITEM hItem)>& callback) const;

		// Custom messages
		_Check_return_ LRESULT msgOnAddItem(WPARAM wParam, LPARAM lParam);
		_Check_return_ LRESULT msgOnDeleteItem(WPARAM wParam, LPARAM lParam);
		_Check_return_ LRESULT msgOnModifyItem(WPARAM wParam, LPARAM lParam);
		_Check_return_ LRESULT msgOnRefreshTable(WPARAM wParam, LPARAM lParam);

		dialog::CHierarchyTableDlg* m_lpHostDlg{nullptr};
		cache::CMapiObjects* m_lpMapiObjects{nullptr};
		LPMAPICONTAINER m_lpContainer{nullptr};
		ULONG m_ulContainerType{NULL};
		ULONG m_ulDisplayFlags{dfNormal};
	};
} // namespace controls