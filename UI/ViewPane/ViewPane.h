#pragma once
#include <core/utility/strings.h>

namespace viewpane
{
	class ViewPane
	{
	public:
		virtual ~ViewPane() = default;

		void SetLabel(const UINT uidLabel) { m_szLabel = strings::loadstring(uidLabel); }
		void SetReadOnly(const bool bReadOnly) noexcept { m_bReadOnly = bReadOnly; }

		virtual void Initialize(_In_ CWnd* pParent, _In_opt_ HDC hdc);
		virtual HDWP DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height);

		virtual void CommitUIValues() = 0;
		virtual bool IsDirty() { return false; }
		virtual int GetMinWidth() { return (m_bCollapsible ? m_iButtonHeight : 0) + m_iLabelWidth; }
		virtual int GetFixedHeight() = 0;
		virtual int GetLines() { return 0; }
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
		int GetID() const noexcept { return m_paneID; }
		UINT GetNID() const noexcept { return m_nID; }

	protected:
		// Returns the height of our label, accounting for an expand/collapse button
		// Will return 0 if we have no label or button
		int GetLabelHeight() const noexcept
		{
			if (m_bCollapsible || !m_szLabel.empty()) return max(m_iButtonHeight, m_iLabelHeight);

			return 0;
		}
		int m_paneID{-1}; // ID of the view pane in the view - used for callbacks and layout
		bool m_bInitialized{};
		bool m_bReadOnly{true};
		std::wstring m_szLabel; // Text to push into UI in Initialize
		int m_iLabelWidth{}; // The width of the label
		CEdit m_Label;
		UINT m_nID{}; // NID for matching change notifications back to controls. Also used for Create calls.
		HWND m_hWndParent{};
		bool m_bCollapsible{};
		bool m_bCollapsed{};
		CButton m_CollapseButton;

		// Margins
		int m_iMargin{};
		int m_iSideMargin{};
		int m_iLabelHeight{}; // Height of the label
		int m_iSmallHeightMargin{};
		int m_iLargeHeightMargin{};
		int m_iButtonHeight{}; // Height of buttons below the control
		int m_iEditHeight{}; // Height of an edit control
	};
} // namespace viewpane