#pragma once

namespace controls
{
	class StyleTreeCtrl : public CTreeCtrl
	{
	public:
		void Create(_In_ CWnd* pCreateParent, UINT nIDContextMenu);
		_Check_return_ bool IsItemSelected() const { return m_bItemSelected; }
		void Refresh();

	protected:
		// Node management
		// Removes any existing node data and replaces it with lpData
		void SetNodeData(HWND hWnd, HTREEITEM hItem, LPARAM lpData) const;
		void OnSelChanged(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);

		UINT m_nIDContextMenu{0};

	private:
		// Overrides for derived controls to customize behavior
		// Override to provide custom deletion of node data
		virtual void FreeNodeData(LPARAM /*lpData*/) const {};
		virtual void OnItemSelected(HTREEITEM /*hItem*/) const {};
		virtual bool HasChildren(_In_ HTREEITEM /*hItem*/) const { return true; }
		virtual void ExpandNode(HTREEITEM /*hParent*/) const {}
		virtual void OnRefresh() const {}

		void OnGetDispInfo(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);

		// Overrides from base class
		LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;

		void OnCustomDraw(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
		void OnRightClick(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
		void OnItemExpanding(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
		_Check_return_ UINT OnGetDlgCode();
		virtual void HandleContextMenu(const int /*x*/, const int /*y*/) {}
		void OnContextMenu(_In_ CWnd* pWnd, CPoint pos);

		HTREEITEM m_hItemCurHover{nullptr};
		bool m_HoverButton{false};
		bool m_bItemSelected{false};

		// TODO: Kill this and use WindowProc instead
		DECLARE_MESSAGE_MAP()
	};
} // namespace controls