#pragma once

namespace controls
{
	class StyleTreeCtrl : public CTreeCtrl
	{
	public:
		void Create(_In_ CWnd* pCreateParent, bool bReadOnly);
		_Check_return_ bool IsItemSelected() const { return m_bItemSelected; }
		void Refresh();
		HTREEITEM
		AddChildNode(
			_In_ const std::wstring& szName,
			HTREEITEM hParent,
			LPARAM lpData,
			const std::function<void(HTREEITEM hItem)>& callback) const;

		// Callbacks
		std::function<bool(HTREEITEM hItem)> HasChildrenCallback = nullptr;
		std::function<void(HTREEITEM hItem)> ItemSelectedCallback = nullptr;
		std::function<bool(UINT nChar, bool bShiftPressed, bool bCtrlPressed, bool bMenuPressed)> KeyDownCallback =
			[&](auto nChar, auto, auto, auto) -> bool {
			if (nChar == VK_ESCAPE)
			{
				::SendMessage(this->GetParent()->GetSafeHwnd(), WM_CLOSE, NULL, NULL);
				return true;
			}

			return false;
		};
		std::function<void(LPARAM lpData)> FreeNodeDataCallback = nullptr;
		std::function<void(HTREEITEM hParent)> ExpandNodeCallback = nullptr;
		std::function<void()> OnRefreshCallback = nullptr;
		std::function<void(HTREEITEM hItem, LPTSTR szText)> OnLabelEditCallback = nullptr;
		std::function<void()> OnDisplaySelectedItemCallback = nullptr;
		std::function<void(LPARAM lpData)> OnLastChildDeletedCallback = nullptr;
		std::function<void(int x, int y)> HandleContextMenuCallback = nullptr;

	protected:
		LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;
		BOOL PreTranslateMessage(MSG* pMsg) override;

		// Node management
		// Removes any existing node data and replaces it with lpData
		void SetNodeData(HWND hWnd, HTREEITEM hItem, LPARAM lpData) const;
		void OnSelChanged(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);

		bool m_bSortNodes{false};
		bool m_bShuttingDown{false};

	private:
		void OnGetDispInfo(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);

		// Overrides from base class
		void OnKeyDown(UINT nChar, UINT nRepCnt, UINT nFlags);

		void OnCustomDraw(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
		void OnRightClick(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
		void OnItemExpanding(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
		_Check_return_ UINT OnGetDlgCode();
		void OnContextMenu(_In_ HWND hwnd, int x, int y);
		void OnEndLabelEdit(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
		void OnDeleteItem(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
		void OnDblclk(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);

		HTREEITEM m_hItemCurHover{nullptr};
		bool m_HoverButton{false};
		bool m_bItemSelected{false};
		bool m_bReadOnly{false};

		DECLARE_MESSAGE_MAP()
	};
} // namespace controls