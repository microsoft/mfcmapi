#include <StdAfx.h>
#include <UI/Dialogs/Editors/Themes.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <UI/ViewPane/ThemePreviewPane.h>
#include <UI/UIFunctions.h>
#include <UI/Themes/Theme.h>
#include <UI/Themes/ThemeFile.h>
#include <core/addin/mfcmapi.h>
#include <core/utility/registry.h>
#include <filesystem>

namespace dialog::editor
{
	class CThemes final : public CEditor
	{
	public:
		CThemes(_In_ CWnd* parent, const std::wstring& selectedTheme);

	private:
		DECLARE_MESSAGE_MAP()
		BOOL OnInitDialog() override;
		afx_msg LRESULT OnThemeChanged(WPARAM wParam, LPARAM lParam);
		_Check_return_ ULONG HandleChange(UINT nID) override;
		void UpdatePreview();
		void ApplySelection(int selection);
		void PopulateThemes();
		void OnEditAction1() override;
		void OnOK() override;

		std::wstring m_selectedTheme;
		std::shared_ptr<viewpane::ThemePreviewPane> m_previewPane;
		bool m_reloading{};
		bool m_initialized{};
	};

	enum ThemeFields
	{
		THEME_APPEARANCE,
		THEME_ENABLED,
		THEME_AUTOMATIC,
		THEME_DIRECTORY,
		THEME_SELECTION,
		THEME_FILE_STATUS,
		THEME_PREVIEW
	};

	std::wstring GetScheduledThemeLabel()
	{
		auto localTime = SYSTEMTIME{};
		GetLocalTime(&localTime);

		const auto date = std::chrono::year{localTime.wYear} / std::chrono::month{localTime.wMonth} /
						  std::chrono::day{localTime.wDay};
		const auto& scheduled = ui::theme::ResolveTheme(
			registry::activeTheme.szCurSTRING, true, true, date, ui::theme::UsesSouthernSeasons());

		return strings::formatmessage(IDS_THEME_AUTOMATIC, std::wstring(scheduled.name).c_str());
	}

	CThemes::CThemes(_In_ CWnd* parent, const std::wstring& selectedTheme)
		: CEditor(
			  parent,
			  IDS_THEME_SELECTION_TITLE,
			  NULL,
			  CEDITOR_BUTTON_OK | CEDITOR_BUTTON_ACTION1,
			  IDS_THEME_RELOAD,
			  NULL,
			  NULL),
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
		AddPane(
			viewpane::CheckPane::Create(THEME_AUTOMATIC, GetScheduledThemeLabel(), registry::automaticThemes, false));

		const auto directory = viewpane::PathPane::Create(
			THEME_DIRECTORY, IDS_THEME_DIRECTORY, viewpane::pathType::directory, registry::themeDirectory.szCurSTRING);
		directory->SetHelpText(IDS_THEME_DIRECTORY_HELP);
		AddPane(directory);

		const auto choices =
			viewpane::DropDownPane::Create(THEME_SELECTION, IDS_THEME_SELECTION_CHOICE, 0, nullptr, true);
		AddPane(choices);

		m_previewPane = viewpane::ThemePreviewPane::Create(THEME_PREVIEW, IDS_THEME_PREVIEW);
		AddPane(m_previewPane);

		AddPane(viewpane::TextPane::CreateMultiLinePane(THEME_FILE_STATUS, IDS_THEME_FILE_STATUS, true));

		PopulateThemes();
		UpdatePreview();
	}

	BEGIN_MESSAGE_MAP(CThemes, CEditor)
	ON_MESSAGE(WM_MFCMAPI_THEMECHANGED, OnThemeChanged)
	END_MESSAGE_MAP()

	BOOL CThemes::OnInitDialog()
	{
		const auto result = CEditor::OnInitDialog();
		m_initialized = true;

		static_cast<void>(OnThemeChanged(0, 0));

		return result;
	}

	LRESULT CThemes::OnThemeChanged(WPARAM /*wParam*/, LPARAM /*lParam*/)
	{
		if (!m_initialized || m_reloading) return 0;

		const auto appearance = std::dynamic_pointer_cast<viewpane::DropDownPane>(GetPane(THEME_APPEARANCE));
		if (appearance) appearance->SetDropDownSelection(static_cast<int>(registry::uiTheme.ulCurDWORD));

		const auto automatic = std::dynamic_pointer_cast<viewpane::CheckPane>(GetPane(THEME_AUTOMATIC));
		if (automatic)
		{
			automatic->SetCheck(registry::automaticThemes);
			automatic->SetCheckLabel(GetScheduledThemeLabel());
		}

		if (!IsDirty(THEME_DIRECTORY)) SetStringW(THEME_DIRECTORY, registry::themeDirectory.szCurSTRING);

		const auto& definition = ui::theme::ResolveTheme(registry::activeTheme.szCurSTRING, true);
		m_selectedTheme = std::wstring(definition.stableId);

		PopulateThemes();
		UpdatePreview();
		OnRecalcLayout();

		return 0;
	}

	void CThemes::PopulateThemes()
	{
		const auto choices = std::dynamic_pointer_cast<viewpane::DropDownPane>(GetPane(THEME_SELECTION));
		if (!choices) return;

		choices->ClearDropStrings();

		const auto& themes = ui::theme::GetThemes();
		auto selection = 0;
		for (size_t index = 0; index < themes.size(); ++index)
		{
			const auto& theme = themes[index];
			const auto duplicates = std::count_if(
				themes.begin(), themes.end(), [&theme](const auto& candidate) { return candidate.name == theme.name; });
			const auto label = duplicates > 1 ? theme.name + L" (" + theme.stableId + L")" : theme.name;
			choices->InsertDropString(label, static_cast<ULONG>(index));

			if (theme.stableId == m_selectedTheme) selection = static_cast<int>(index);
		}

		const auto& errors = ui::theme::GetThemeFileErrors();
		auto status = std::wstring{};
		for (size_t index = 0; index < errors.size(); ++index)
		{
			const auto& error = errors[index];
			const auto filename = std::filesystem::path{error.path}.filename().wstring();

			choices->InsertDropString(
				strings::formatmessage(
					IDS_THEME_FILE_IN_ERROR, filename.empty() ? error.path.c_str() : filename.c_str()),
				static_cast<ULONG>(themes.size() + index));

			if (!status.empty()) status += L"\r\n";
			status += error.path + L": " + error.message;
		}

		choices->SetDropDownSelection(selection);
		SetStringW(THEME_FILE_STATUS, status);
		std::dynamic_pointer_cast<viewpane::TextPane>(GetPane(THEME_FILE_STATUS))->SetVisible(!status.empty());
	}

	void CThemes::OnEditAction1()
	{
		m_reloading = true;
		ui::theme::ReloadThemes(GetStringW(THEME_DIRECTORY));
		m_reloading = false;

		static_cast<void>(OnThemeChanged(0, 0));
		OnRecalcLayout();
	}

	void CThemes::OnOK()
	{
		if (IsDirty(THEME_DIRECTORY) && GetStringW(THEME_DIRECTORY) != registry::themeDirectory.szCurSTRING)
		{
			OnEditAction1();
		}

		CEditor::OnOK();
	}

	_Check_return_ ULONG CThemes::HandleChange(const UINT nID)
	{
		const auto paneID = CEditor::HandleChange(nID);

		if (!m_initialized || m_reloading) return paneID;

		if (IsDirty(THEME_DIRECTORY) && GetStringW(THEME_DIRECTORY) != registry::themeDirectory.szCurSTRING)
		{
			OnEditAction1();
		}

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
		else if (paneID == THEME_AUTOMATIC)
		{
			ui::theme::SetAutomaticThemes(GetCheck(THEME_AUTOMATIC));
		}
		else if (paneID == THEME_SELECTION)
		{
			const auto selection = GetDropDown(THEME_SELECTION);
			const auto& themes = ui::theme::GetThemes();
			if (selection >= static_cast<int>(themes.size()))
			{
				const auto& errors = ui::theme::GetThemeFileErrors();
				const auto errorIndex = static_cast<size_t>(selection) - themes.size();
				if (errorIndex < errors.size())
				{
					SetStringW(THEME_FILE_STATUS, errors[errorIndex].path + L": " + errors[errorIndex].message);
				}

				UpdatePreview();
				OnRecalcLayout();

				return paneID;
			}

			ApplySelection(selection);
			UpdatePreview();
			OnRecalcLayout();
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
		const auto selectedTheme =
			selection >= 0 && selection < static_cast<int>(themes.size()) ? &themes[selection] : nullptr;
		m_previewPane->SetVisible(selectedTheme != nullptr);
		if (!selectedTheme)
		{
			m_previewPane->ClearTheme();

			return;
		}

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
