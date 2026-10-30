#pragma once
#include <UI/ViewPane/ViewPane.h>
#include <UI/Themes/Theme.h>

namespace viewpane
{
	class ThemePreviewPane final : public ViewPane
	{
	public:
		static std::shared_ptr<ThemePreviewPane> Create(int paneID, UINT uidLabel);
		void SetTheme(const ui::theme::ThemeDefinition& definition, ui::theme::Appearance appearance);
		void ClearTheme();
		void SetVisible(bool visible) noexcept { m_visible = visible; }
		void Initialize(_In_ CWnd* pParent, _In_opt_ HDC hdc) override;
		HDWP DeferWindowPos(_In_ HDWP hWinPosInfo, _In_ int x, _In_ int y, _In_ int width, _In_ int height) override;
		int GetMinWidth() override;
		int GetFixedHeight() override;
		bool containsWindow(HWND hWnd) const noexcept override;
		RECT GetWindowRect() const noexcept override;

	private:
		static LRESULT CALLBACK
		PreviewProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam, UINT_PTR subclassID, DWORD_PTR data);
		void CommitUIValues() override {}

		HWND m_hWndPreview{};
		bool m_hasTheme{};
		bool m_visible{true};
		bool m_highContrast{};
		std::wstring m_themeName;
		std::array<COLORREF, static_cast<size_t>(ui::uiColor::UIEnd)> m_colors{};
	};
} // namespace viewpane
