#pragma once
#include <core/utility/strings.h>
#include <UI/Controls/PaneHeader/PaneHeader.h>

namespace viewpane
{
	class ViewPane
	{
	public:
		virtual ~ViewPane() = default;

		void SetLabel(const UINT uidLabel) { m_Header.SetLabel(uidLabel); }
		void SetLabel(const std::wstring& szLabel) { m_Header.SetLabel(szLabel); }
		void SetReadOnly(const bool bReadOnly) noexcept { m_bReadOnly = bReadOnly; }

		virtual void Initialize(_In_ CWnd* pParent, _In_opt_ HDC hdc);
		virtual HDWP DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height);

		virtual void CommitUIValues() = 0;
		virtual bool IsDirty() { return false; }
		virtual int GetMinWidth() { return m_Header.GetMinWidth(); }
		virtual int GetFixedHeight() = 0;
		virtual int GetLines() { return 0; }
		virtual ULONG HandleChange(UINT nID);

		virtual void SetMargins(
			int iMargin,
			int iSideMargin,
			int iLabelHeight, // Height of the label
			int iSmallHeightMargin,
			int iLargeHeightMargin,
			int iButtonHeight, // Height of buttons below the control
			int iEditHeight); // Height of an edit control
		virtual void UpdateButtons();
		int GetID() const noexcept { return m_paneID; }
		UINT GetNID() const noexcept { return m_nID; }
		void makeCollapsible() noexcept { m_Header.makeCollapsible(); }
		bool collapsible() const noexcept { m_Header.collapsible(); }
		bool collapsed() const noexcept { return m_Header.collapsed(); }
		virtual bool containsWindow(HWND hWnd) const noexcept { return m_Header.containsWindow(hWnd); }
		virtual RECT GetWindowRect() const noexcept
		{
			// TODO: ViewPanes which will scroll need to override this
			auto rcPane = RECT{};
			::GetWindowRect(m_Header.GetSafeHwnd(), &rcPane);
			return rcPane;
		}

	protected:
		// Returns the height of our header
		int GetHeaderHeight() const noexcept { return m_Header.GetFixedHeight(); }
		int m_paneID{-1}; // ID of the view pane in the view - used for callbacks and layout
		bool m_bInitialized{};
		bool m_bReadOnly{true};
		controls::PaneHeader m_Header;
		UINT m_nID{}; // NID for matching change notifications back to controls. Also used for Create calls.

		// Margins
		int m_iMargin{};
		int m_iSideMargin{};
		int m_iSmallHeightMargin{};
		int m_iLargeHeightMargin{};
		int m_iButtonHeight{}; // Height of buttons below the control
		int m_iEditHeight{}; // Height of an edit control

	private:
		HWND m_hWndParent{};
	};
} // namespace viewpane