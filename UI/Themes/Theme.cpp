#include <StdAfx.h>
#include <UI/Themes/Theme.h>
#include <UI/Themes/BuiltinThemes.h>
#include <UI/UIFunctions.h>
#include <UI/Dialogs/BaseDialog.h>
#include <core/utility/registry.h>
#include <core/utility/strings.h>
#include <algorithm>

namespace ui::theme
{
	void SetAppearance(const uiThemePreference appearance)
	{
		if (appearance != uiThemePreference::System && appearance != uiThemePreference::Light &&
			appearance != uiThemePreference::Dark)
		{
			WC_H_S(E_INVALIDARG);
			return;
		}

		if (registry::uiTheme.ulCurDWORD == static_cast<DWORD>(appearance)) return;

		registry::uiTheme = static_cast<DWORD>(appearance);
		registry::WriteToRegistry();
		ApplyThemeToAllWindows();
	}

	void ToggleAppearance()
	{
		const auto appearance = GetEffectiveTheme();
		const auto dark = appearance == uiTheme::Dark ||
						  (appearance == uiTheme::HighContrast &&
						   registry::uiTheme.ulCurDWORD == static_cast<DWORD>(uiThemePreference::Dark));

		SetAppearance(dark ? uiThemePreference::Light : uiThemePreference::Dark);
	}

	void SetSelectedTheme(const std::wstring_view stableId)
	{
		if (!FindTheme(stableId))
		{
			WC_H_S(E_INVALIDARG);
			return;
		}

		if (registry::activeTheme.szCurSTRING == stableId) return;

		registry::activeTheme = std::wstring(stableId);
		registry::WriteToRegistry();
		ApplyThemeToAllWindows();
	}

	void SetThemesEnabled(const bool enabled)
	{
		if (static_cast<bool>(registry::themesEnabled) == enabled) return;

		registry::themesEnabled = enabled;
		registry::WriteToRegistry();
		ApplyThemeToAllWindows();
	}

	void CycleTheme()
	{
		if (!registry::themesEnabled) return;

		const auto& next = GetNextTheme(registry::activeTheme.szCurSTRING);
		SetSelectedTheme(next.stableId);

		const auto activeWindow = GetActiveWindow();
		if (activeWindow && registry::themesEnabled)
		{
			dialog::CBaseDialog::UpdateStatus(
				GetAncestor(activeWindow, GA_ROOT),
				statusPane::infoText,
				strings::formatmessage(IDS_THEME_CYCLE_STATUS, std::wstring(next.name).c_str()));
		}
	}

	const std::vector<ThemeDefinition>& GetThemes()
	{
		static const std::vector<ThemeDefinition> themes(builtin::Definitions.begin(), builtin::Definitions.end());

		return themes;
	}

	const ThemeDefinition* FindTheme(const std::wstring_view stableId)
	{
		const auto& themes = GetThemes();
		const auto found = std::find_if(themes.begin(), themes.end(), [stableId](const auto& definition) {
			return definition.stableId == stableId;
		});

		return found == themes.end() ? nullptr : &*found;
	}

	const ThemeDefinition& GetDefaultTheme() { return GetThemes().front(); }

	const ThemeDefinition& ResolveTheme(const std::wstring_view selectedTheme, const bool enabled)
	{
		if (enabled)
		{
			if (const auto definition = FindTheme(selectedTheme)) return *definition;
		}

		return GetDefaultTheme();
	}

	COLORREF GetColor(const ThemeDefinition& definition, const uiTheme appearance, const uiColor role) noexcept
	{
		const auto& colors = appearance == uiTheme::Dark ? definition.dark : definition.light;
		const auto& color = colors.values.at(static_cast<size_t>(role));
		if (appearance == uiTheme::HighContrast && color.highContrast >= 0) return GetSysColor(color.highContrast);

		if (color.system >= 0) return GetSysColor(color.system);

		return color.value;
	}

	const ThemeDefinition& GetNextTheme(const std::wstring_view stableId)
	{
		const auto& themes = GetThemes();
		const auto current = FindTheme(stableId);
		const auto index = current ? static_cast<size_t>(current - themes.data()) : 0;

		return themes[(index + 1) % themes.size()];
	}

	bool TryGetColor(
		const ThemeDefinition& definition,
		const Appearance appearance,
		const ColorRole role,
		_Out_ COLORREF& color) noexcept
	{
		const auto roleIndex = static_cast<size_t>(role);
		if (roleIndex >= static_cast<size_t>(ColorRole::Count) ||
			(appearance != Appearance::Light && appearance != Appearance::Dark))
			return false;

		static constexpr std::array<uiColor, static_cast<size_t>(ColorRole::Count)> roles = {
			uiColor::Glow,
			uiColor::GlowBackground,
			uiColor::SelectedBackground,
			uiColor::GlowText,
			uiColor::FrameSelected,
			uiColor::Status,
			uiColor::StatusText,
			uiColor::MenuGutter};
		color = GetColor(definition, appearance == Appearance::Dark ? uiTheme::Dark : uiTheme::Light, roles[roleIndex]);

		return true;
	}
} // namespace ui::theme
