#pragma once
#include <core/utility/strings.h>

namespace controls
{
	class ViewHeader : public CEdit
	{
	public:
		virtual ~ViewHeader() = default;

		void SetLabel(const UINT uidLabel) { m_szLabel = strings::loadstring(uidLabel); }
		void SetLabel(const std::wstring szLabel) { m_szLabel = szLabel; }

		void Initialize(_In_ CWnd* pParent, _In_opt_ HDC hdc, _In_ bool bCollapsible, _In_ UINT nidParent);
		void DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height);

		int GetMinWidth() { return (m_bCollapsible ? m_iButtonHeight : 0) + m_iLabelWidth; }
		int GetLines() { return 0; }
		bool HandleChange(UINT nID);
		void OnToggleCollapse();
		std::function<void()> ToggleCollapseCallback = nullptr;

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
		// Returns the height of our label, accounting for an expand/collapse button
		// Will return 0 if we have no label or button
		int GetFixedHeight() const noexcept
		{
			if (m_bCollapsible || !m_szLabel.empty()) return max(m_iButtonHeight, m_iLabelHeight);

			return 0;
		}

	protected:
		int m_paneID{-1}; // ID of the view pane in the view - used for callbacks and layout
		bool m_bInitialized{};
		std::wstring m_szLabel; // Text to push into UI in Initialize
		int m_iLabelWidth{}; // The width of the label
		UINT m_nIDCollapse{}; // NID for collapse button.
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
} // namespace controls