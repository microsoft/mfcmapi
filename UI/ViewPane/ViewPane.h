#pragma once

namespace viewpane
{
	class ViewPane
	{
	public:
		virtual ~ViewPane() = default;

		void SetLabel(UINT uidLabel, bool bReadOnly);
		virtual void Initialize(_In_ CWnd* pParent, _In_opt_ HDC hdc);
		virtual void DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height);

		virtual void CommitUIValues() = 0;
		virtual bool IsDirty();
		virtual int GetMinWidth(_In_ HDC hdc);
		virtual int GetFixedHeight() = 0;
		virtual int GetLines();
		virtual ULONG HandleChange(UINT nID);
		void OnToggleCollapse();

		virtual void SetMargins(
			int iMargin,
			int iSideMargin,
			int iLabelHeight, // Height of the label
			int iSmallHeightMargin,
			int iLargeHeightMargin,
			int iButtonHeight, // Height of buttons below the control
			int iEditHeight); // Height of an edit control
		void SetAddInLabel(const std::wstring& szLabel);
		virtual void UpdateButtons();
		ULONG GetID() { return m_iControl; }
		void SetID(int id) { m_iControl = id; }
		// Return a pane with a matching control #. Can be overriden to allow container panes to return subpanes.
		virtual ViewPane* GetPaneByID(int iControl) { return m_iControl == iControl ? this : nullptr; }
		// Return a pane with a matching nID. Can be overriden to allow container panes to return subpanes.
		virtual ViewPane* GetPaneByNID(UINT nID) { return nID == m_nID ? this : nullptr; }

	protected:
		int m_iControl{-1}; // ID of the view pane in the view - used for callbacks and layout
		bool m_bInitialized{false};
		bool m_bReadOnly{true};
		std::wstring m_szLabel; // Text to push into UI in Initialize
		int m_iLabelWidth; // The width of the label
		CEdit m_Label;
		UINT m_nID{0}; // NID for matching change notifications back to controls. Also used for Create calls.
		HWND m_hWndParent{nullptr};
		bool m_bCollapsible{false};
		bool m_bCollapsed{false};
		CButton m_CollapseButton;

		// Margins
		int m_iMargin{0};
		int m_iSideMargin{0};
		int m_iLabelHeight{0}; // Height of the label
		int m_iSmallHeightMargin{0};
		int m_iLargeHeightMargin{0};
		int m_iButtonHeight{0}; // Height of buttons below the control
		int m_iEditHeight{0}; // Height of an edit control
	};
} // namespace viewpane