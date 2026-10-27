#pragma once

#include <Windows.h>
#include <array>
#include <string_view>
#include <vector>

namespace ui
{
	enum class uiThemePreference : DWORD
	{
		System,
		Light,
		Dark,
	};

	enum class uiTheme
	{
		Light,
		Dark,
		HighContrast,
	};

	enum class uiColor
	{
		Background,
		BackgroundReadOnly,
		Glow,
		GlowBackground,
		GlowText,
		FrameSelected,
		FrameUnselected,
		SelectedBackground,
		Arrow,
		Text,
		TextDisabled,
		TextReadOnly,
		BitmapTransBack,
		BitmapTransFore,
		Status,
		StatusText,
		PaneHeaderBackground,
		PaneHeaderText,
		TextHighlightBackground,
		TextHighlight,
		MenuGutter,
		TestPink,
		TestLavender,
		TestRed,
		TestGreen,
		TestOrange,
		UIEnd
	};
} // namespace ui

namespace ui::theme
{
	enum class Appearance
	{
		Light,
		Dark
	};

	enum class ColorRole
	{
		Accent,
		AccentSecondary,
		SelectionBackground,
		SelectionText,
		Focus,
		HeroBackground,
		HeroText,
		MenuGutter,
		Count
	};

	enum class ScheduleKind
	{
		None,
		AnnualDateRange
	};

	struct Schedule
	{
		ScheduleKind kind;
		unsigned int startMonth;
		unsigned int startDay;
		unsigned int endMonth;
		unsigned int endDay;
	};

	struct ThemeColor
	{
		COLORREF value;
		int system = -1;
		int highContrast = -1;
	};

	struct ThemeColors
	{
		std::array<ThemeColor, static_cast<size_t>(uiColor::UIEnd)> values;
	};

	struct ThemeDefinition
	{
		std::wstring_view stableId;
		std::wstring_view name;
		std::wstring_view category;
		std::wstring_view variant;
		Schedule schedule;
		ThemeColors light;
		ThemeColors dark;
	};

	const std::vector<ThemeDefinition>& GetThemes();
	const ThemeDefinition* FindTheme(std::wstring_view stableId);
	const ThemeDefinition& GetDefaultTheme();
	const ThemeDefinition& ResolveTheme(std::wstring_view selectedTheme, bool enabled);
	COLORREF GetColor(const ThemeDefinition& definition, uiTheme appearance, uiColor role) noexcept;
	const ThemeDefinition& GetNextTheme(std::wstring_view stableId);
	void SetAppearance(uiThemePreference appearance);
	void ToggleAppearance();
	void SetSelectedTheme(std::wstring_view stableId);
	void SetThemesEnabled(bool enabled);
	void CycleTheme();
	bool TryGetColor(
		const ThemeDefinition& definition,
		Appearance appearance,
		ColorRole role,
		_Out_ COLORREF& color) noexcept;
} // namespace ui::theme
