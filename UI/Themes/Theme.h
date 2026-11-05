#pragma once

#include <Windows.h>
#include <array>
#include <chrono>
#include <string>
#include <string_view>
#include <span>
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
		DisabledSelectedBackground,
		DisabledSelectedBorder,
		Arrow,
		Text,
		TextDisabled,
		TextDisabledSelected,
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
		TitleBarBackground,
		TitleBarText,
		TitleBarButton,
		TitleBarButtonHover,
		ColumnHeaderBackground,
		ColumnHeaderText,
		ColumnHeaderArrow,
		MenuBackground,
		MenuText,
		MenuTextDisabled,
		StatusTextOutline,
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
		int priority = 0;
		bool seasonal = false;
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

	enum class StatusBarStyle
	{
		Ribbon,
		Bands,
		Gradient
	};

	struct ThemeDefinition
	{
		std::wstring stableId;
		std::wstring name;
		Schedule schedule;
		ThemeColors light;
		ThemeColors dark;
		std::vector<COLORREF> statusRibbon{};
		StatusBarStyle statusBarStyle = StatusBarStyle::Ribbon;
	};

	struct StatusRibbonBand
	{
		RECT bounds;
		COLORREF color;
		COLORREF endColor;
	};

	std::span<const COLORREF> GetStatusRibbon(const ThemeDefinition& definition, uiTheme appearance) noexcept;
	std::vector<StatusRibbonBand>
	LayoutStatusRibbon(std::span<const COLORREF> colors, const RECT& rect, int height, bool gradient = false);

	const std::vector<ThemeDefinition>& GetThemes();
	const ThemeDefinition* FindTheme(std::wstring_view stableId);
	const ThemeDefinition& GetDefaultTheme();
	const ThemeDefinition& ResolveTheme(std::wstring_view selectedTheme, bool enabled);
	const ThemeDefinition& ResolveTheme(
		std::wstring_view selectedTheme,
		bool enabled,
		bool automatic,
		std::chrono::year_month_day date,
		bool southernHemisphere);
	const ThemeDefinition* FindScheduledTheme(
		const std::vector<ThemeDefinition>& themes,
		std::chrono::year_month_day date,
		bool southernHemisphere);
	bool UsesSouthernSeasons();
	const ThemeDefinition& ResolveActiveTheme();
	void StartThemeSchedule() noexcept;
	void StopThemeSchedule() noexcept;
	void RefreshScheduledTheme() noexcept;
	COLORREF GetColor(const ThemeDefinition& definition, uiTheme appearance, uiColor role) noexcept;
	const ThemeDefinition& GetNextTheme(std::wstring_view stableId, bool reverse = false);
	void SetAppearance(uiThemePreference appearance);
	void ToggleAppearance();
	void SetSelectedTheme(std::wstring_view stableId);
	void SetThemesEnabled(bool enabled);
	void SetAutomaticThemes(bool enabled);
	void CycleTheme(bool reverse = false);
	void ReloadThemes(const std::wstring& directory);
	bool TryGetColor(
		const ThemeDefinition& definition,
		Appearance appearance,
		ColorRole role,
		_Out_ COLORREF& color) noexcept;
} // namespace ui::theme
