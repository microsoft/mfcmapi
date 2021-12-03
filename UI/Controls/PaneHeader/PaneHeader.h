#pragma once
#include <core/utility/strings.h>

namespace controls
{
	class PaneHeader : public CWnd
	{
	public:
		PaneHeader() = default;
		~PaneHeader();

		void Initialize(HWND hWnd, _In_opt_ HDC hdc, _In_ UINT nid);
		int GetMinWidth();

		void SetRightLabel(const std::wstring szLabel);
		void EnableActionButton(_In_ UINT nid) { m_nIDAction = nid; }
		void SetActionButton(const std::wstring szActionButton);
		bool HandleChange(UINT nID);
		void OnToggleCollapse();

		_NODISCARD bool empty() const noexcept { return m_szLabel.empty(); }
		void SetMargins(
			int iMargin,
			int iSideMargin,
			int iLabelHeight, // Height of the label
			int iButtonHeight); // Height of buttons below the control

		void SetLabel(const UINT uidLabel) { m_szLabel = strings::loadstring(uidLabel); }
		void SetLabel(const std::wstring szLabel) { m_szLabel = szLabel; }

		// Returns the height of our control, accounting for an expand/collapse button
		// Will return 0 if we have no label or button
		int GetFixedHeight() const noexcept
		{
			if (m_bCollapsible || !m_szLabel.empty()) return max(m_iButtonHeight, m_iLabelHeight);

			return 0;
		}
		void makeCollapsible() noexcept { m_bCollapsible = true; }
		bool collapsible() const noexcept { return m_bCollapsible; }
		bool collapsed() const noexcept { return m_bCollapsed; }
		bool containsWindow(HWND hWnd) const noexcept;

	private:
		void RecalcLayout();
		int OnCreate(LPCREATESTRUCT lpCreateStruct);
		LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;
		BOOL PreTranslateMessage(MSG* pMsg);
		bool m_bInitialized{};
		std::wstring m_szLabel; // Text to push into UI in Initialize
		UINT m_nID{};
		UINT m_nIDCollapse{}; // NID for collapse button.
		HWND m_hWndParent{};
		bool m_bCollapsible{};
		bool m_bCollapsed{};
		CButton m_CollapseButton;

		// Margins
		int m_iMargin{};
		int m_iSideMargin{};
		int m_iLabelWidth{}; // The width of the label
		int m_iLabelHeight{}; // Height of the label
		int m_iButtonHeight{}; // Height of button

		CEdit m_leftLabel; // Label on left of header
		int m_leftLabelWidth{0}; // The width of the string

		CEdit m_rightLabel; // Label on right of header
		int m_rightLabelWidth{0}; // The width of the string

		CButton m_actionButton;
		std::wstring m_szActionButton;
		int m_actionButtonWidth{0};
		UINT m_nIDAction{}; // NID for action button.

		DECLARE_MESSAGE_MAP()
	};
} // namespace controls