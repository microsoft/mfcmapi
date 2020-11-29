#pragma once
#include <UI/enums.h>
#include <core/sortlistdata/sortListData.h>
#include <UI/Controls/StyleTree/StyleTreeCtrl.h>
#include <UI/Dialogs/BaseDialog.h>

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
		~CHierarchyTableTreeCtrl();

		// Initialization
		void Create(
			_In_ CWnd* pCreateParent,
			_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
			_In_ dialog::CBaseDialog* lpHostDlg,
			tableDisplayFlags displayFlags,
			UINT nIDContextMenu);
		void LoadHierarchyTable(_In_ LPMAPICONTAINER lpMAPIContainer);

		// Selected item accessors
		_Check_return_ LPMAPICONTAINER GetSelectedContainer(modifyType bModify) const;
		_Check_return_ std::vector<BYTE> GetSelectedItemEID() const;
		_Check_return_ sortlistdata::sortListData* GetSelectedItemData() const;
		_Check_return_ sortlistdata::sortListData* GetSortListData(HTREEITEM iItem) const;

	private:
		// Overrides from base class
		LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;

		// Callback functions
		void OnItemAdded(HTREEITEM hItem) const;
		bool HasChildren(_In_ HTREEITEM hItem) const;
		void OnItemSelected(HTREEITEM hItem) const;
		bool HandleKeyDown(UINT nChar, bool bShiftPressed, bool bCtrlPressed, bool bMenuPressed);
		void FreeNodeData(LPARAM lpData) const noexcept;
		void ExpandNode(HTREEITEM hParent) const;
		void OnRefresh() const;
		void OnLabelEdit(HTREEITEM hItem, LPTSTR szText);
		void OnDisplaySelectedItem();
		void OnLastChildDeleted(LPARAM /*lpData*/) noexcept;
		void HandleContextMenu(int x, int y);
		void OnCustomDraw(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult, _In_ HTREEITEM hItemCurHover) noexcept;

		_Check_return_ HTREEITEM FindNode(_In_ const SBinary& lpInstance, HTREEITEM hParent) const;
		_Check_return_ LPMAPICONTAINER GetContainer(HTREEITEM Item, modifyType bModify) const;
		void Advise(HTREEITEM hItem, sortlistdata::sortListData* lpData) const;
		_Check_return_ LPMAPITABLE
		GetHierarchyTable(HTREEITEM hItem, _In_opt_ LPMAPICONTAINER lpMAPIContainer, bool bRegNotifs) const;

		// Node management
		void AddRootNode() const;
		void
		AddNode(_In_ LPSRow lpsRow, HTREEITEM hParent, const std::function<void(HTREEITEM hItem)>& itemAddedCallback)
			const;

		// Custom messages
		_Check_return_ LRESULT msgOnAddItem(WPARAM wParam, LPARAM lParam);
		_Check_return_ LRESULT msgOnDeleteItem(WPARAM wParam, LPARAM lParam);
		_Check_return_ LRESULT msgOnModifyItem(WPARAM wParam, LPARAM lParam);
		_Check_return_ LRESULT msgOnRefreshTable(WPARAM wParam, LPARAM lParam);

		dialog::CBaseDialog* m_lpHostDlg{};
		std::shared_ptr<cache::CMapiObjects> m_lpMapiObjects{};
		LPMAPICONTAINER m_lpContainer{};
		ULONG m_ulContainerType{};
		tableDisplayFlags m_displayFlags{tableDisplayFlags::dfNormal};
		UINT m_nIDContextMenu{};
	};
} // namespace controls