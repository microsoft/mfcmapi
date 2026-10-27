#include <StdAfx.h>
#include <UI/Dialogs/Editors/Themes.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/ViewPane/ThemePreviewPane.h>
#include <UI/UIFunctions.h>
#include <UI/Themes/Theme.h>
#include <core/addin/mfcmapi.h>
#include <core/utility/registry.h>

namespace dialog::editor
{
	class CThemes final : public CEditor
	{
	public:
		CThemes(_In_ CWnd* parent, const std::wstring& selectedTheme);

	private:
		DECLARE_MESSAGE_MAP()
		afx_msg LRESULT OnThemeChanged(WPARAM wParam, LPARAM lParam);
		_Check_return_ ULONG HandleChange(UINT nID) override;
		void UpdatePreview();
		void ApplySelection(int selection);

		std::wstring m_selectedTheme;
		std::shared_ptr<viewpane::ThemePreviewPane> m_previewPane;
	};

	enum ThemeFields
	{
		THEME_APPEARANCE,
		THEME_ENABLED,
		THEME_SELECTION,
		THEME_PREVIEW
	};

	CThemes::CThemes(_In_ CWnd* parent, const std::wstring& selectedTheme)
		: CEditor(parent, IDS_THEME_SELECTION_TITLE, NULL, CEDITOR_BUTTON_OK),
		  m_selectedTheme(
			  ui::theme::FindTheme(selectedTheme) ? selectedTheme : std::wstring(ui::theme::GetDefaultTheme().stableId))
	{
		EnableScroll();

		const auto appearance = viewpane::DropDownPane::Create(
			THEME_APPEARANCE,
			IDS_REGKEY_UI_THEME,
			static_cast<ULONG>(registry::uiTheme.uiDropDownOptions.size()),
			registry::uiTheme.uiDropDownOptions.data(),
			true);
		auto appearanceIndex = registry::uiTheme.ulCurDWORD;
		if (appearanceIndex >= registry::uiTheme.uiDropDownOptions.size()) appearanceIndex = 0;

		appearance->SetDropDownSelection(static_cast<int>(appearanceIndex));
		AddPane(appearance);

		AddPane(viewpane::CheckPane::Create(THEME_ENABLED, IDS_THEMES_ENABLED, registry::themesEnabled, false));

		const auto& themes = ui::theme::GetThemes();
		const auto choices =
			viewpane::DropDownPane::Create(THEME_SELECTION, IDS_THEME_SELECTION_CHOICE, 0, nullptr, true);
		for (size_t index = 0; index < themes.size(); ++index)
		{
			choices->InsertDropString(std::wstring(themes[index].name), static_cast<ULONG>(index));
		}

		choices->SetDropDownSelection(
			static_cast<int>(std::distance(
				themes.begin(), std::find_if(themes.begin(), themes.end(), [this](const auto& definition) {
					return definition.stableId == m_selectedTheme;
				}))));
		AddPane(choices);

		m_previewPane = viewpane::ThemePreviewPane::Create(THEME_PREVIEW, IDS_THEME_PREVIEW);
		AddPane(m_previewPane);

		UpdatePreview();
	}

	BEGIN_MESSAGE_MAP(CThemes, CEditor)
	ON_MESSAGE(WM_MFCMAPI_THEMECHANGED, OnThemeChanged)
	END_MESSAGE_MAP()

	LRESULT CThemes::OnThemeChanged(WPARAM /*wParam*/, LPARAM /*lParam*/)
	{
		const auto appearance = std::dynamic_pointer_cast<viewpane::DropDownPane>(GetPane(THEME_APPEARANCE));
		if (appearance) appearance->SetDropDownSelection(static_cast<int>(registry::uiTheme.ulCurDWORD));

		const auto& definition = ui::theme::ResolveTheme(registry::activeTheme.szCurSTRING, true);
		m_selectedTheme = std::wstring(definition.stableId);

		const auto choices = std::dynamic_pointer_cast<viewpane::DropDownPane>(GetPane(THEME_SELECTION));
		if (choices)
		{
			const auto& themes = ui::theme::GetThemes();
			const auto selected = std::find_if(themes.begin(), themes.end(), [&definition](const auto& candidate) {
				return candidate.stableId == definition.stableId;
			});
			choices->SetDropDownSelection(static_cast<int>(std::distance(themes.begin(), selected)));
			UpdatePreview();
		}

		return 0;
	}

	_Check_return_ ULONG CThemes::HandleChange(const UINT nID)
	{
		const auto paneID = CEditor::HandleChange(nID);

		if (paneID == THEME_APPEARANCE)
		{
			const auto selection = GetDropDown(THEME_APPEARANCE);
			if (selection >= 0 && selection < static_cast<int>(registry::uiTheme.uiDropDownOptions.size()))
			{
				ui::theme::SetAppearance(static_cast<ui::uiThemePreference>(selection));
			}
		}
		else if (paneID == THEME_ENABLED)
		{
			ui::theme::SetThemesEnabled(GetCheck(THEME_ENABLED));
		}
		else if (paneID == THEME_SELECTION)
		{
			ApplySelection(GetDropDown(THEME_SELECTION));
			UpdatePreview();
		}

		return paneID;
	}

	void CThemes::ApplySelection(const int selection)
	{
		const auto& themes = ui::theme::GetThemes();
		if (selection < 0 || selection >= static_cast<int>(themes.size())) return;

		m_selectedTheme = std::wstring(themes[selection].stableId);
		ui::theme::SetSelectedTheme(m_selectedTheme);
	}

	void CThemes::UpdatePreview()
	{
		const auto selection = GetDropDown(THEME_SELECTION);
		const auto& themes = ui::theme::GetThemes();
		const auto selectedTheme = selection >= 0 && selection < static_cast<int>(themes.size())
									   ? &themes[selection]
									   : ui::theme::FindTheme(m_selectedTheme);
		if (!selectedTheme) return;

		const auto& definition = *selectedTheme;
		const auto appearance =
			ui::GetEffectiveTheme() == ui::uiTheme::Dark ? ui::theme::Appearance::Dark : ui::theme::Appearance::Light;
		m_previewPane->SetTheme(definition, appearance);
	}
	void DisplayThemesDlg(_In_ CWnd* parent)
	{
		CThemes themes(parent, registry::activeTheme.szCurSTRING);
		static_cast<void>(themes.DisplayDialog());
	}
} // namespace dialog::editor
