#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <UI/Themes/Theme.h>
#include <algorithm>

namespace themetest
{
	TEST_CLASS(ThemeTest){public : TEST_METHOD(CatalogIdsAreUniqueAndResolveToTheirDefinitions){
		const auto& themes = ui::theme::GetThemes();
	Assert::IsFalse(themes.empty());

	for (const auto& theme : themes)
	{
		Assert::IsFalse(theme.stableId.empty());
		Assert::IsTrue(ui::theme::FindTheme(theme.stableId) == &theme);
	}
} // namespace themetest

TEST_METHOD(ResolutionUsesSelectionOnlyWhenEnabledAndFallsBackForUnknownIds)
{
	const auto& standard = ui::theme::GetDefaultTheme();
	Assert::IsTrue(&ui::theme::ResolveTheme(L"unknown", true) == &standard);
	Assert::IsTrue(&ui::theme::ResolveTheme(L"", true) == &standard);

	for (const auto& theme : ui::theme::GetThemes())
	{
		Assert::IsTrue(&ui::theme::ResolveTheme(theme.stableId, false) == &standard);
		Assert::IsTrue(&ui::theme::ResolveTheme(theme.stableId, true) == &theme);
	}
}

TEST_METHOD(ThemeDefinitionsOwnEveryColor)
{
	auto theme = ui::theme::GetDefaultTheme();
	theme.stableId = L"test.complete";
	for (size_t index = 0; index < static_cast<size_t>(ui::uiColor::UIEnd); ++index)
	{
		const auto channel = static_cast<BYTE>(index);
		theme.light.values[index] = {RGB(channel, 37, 83), -1, -1};
		theme.dark.values[index] = {RGB(149, channel, 211), -1, -1};
	}

	for (auto index = 0; index < static_cast<int>(ui::uiColor::UIEnd); ++index)
	{
		const auto role = static_cast<ui::uiColor>(index);
		const auto channel = static_cast<BYTE>(index);
		Assert::AreEqual(RGB(channel, 37, 83), ui::theme::GetColor(theme, ui::uiTheme::Light, role));
		Assert::AreEqual(RGB(149, channel, 211), ui::theme::GetColor(theme, ui::uiTheme::Dark, role));
		Assert::AreEqual(RGB(channel, 37, 83), ui::theme::GetColor(theme, ui::uiTheme::HighContrast, role));

		theme.light.values[index].system = COLOR_WINDOW;
		theme.light.values[index].highContrast = COLOR_HIGHLIGHT;
		Assert::AreEqual(GetSysColor(COLOR_WINDOW), ui::theme::GetColor(theme, ui::uiTheme::Light, role));
		Assert::AreEqual(GetSysColor(COLOR_HIGHLIGHT), ui::theme::GetColor(theme, ui::uiTheme::HighContrast, role));
	}
}

TEST_METHOD(SeasonalSchedulingCoversEveryDateInBothHemispheres)
{
	using namespace std::chrono;
	const std::wstring_view seasons[] = {L"season.winter", L"season.spring", L"season.summer", L"season.autumn"};
	for (const auto year : {2023, 2024})
	{
		for (auto date = sys_days{std::chrono::year{year} / January / 1};
			 date < sys_days{std::chrono::year{year + 1} / January / 1};
			 date += days{1})
		{
			const auto calendar = year_month_day{date};
			const auto season = unsigned(calendar.month()) % 12 / 3;
			for (const auto southern : {false, true})
			{
				const auto selected = ui::theme::FindScheduledTheme(ui::theme::GetThemes(), calendar, southern);
				Assert::IsNotNull(selected);
				Assert::IsTrue(selected->stableId == seasons[(season + (southern ? 2 : 0)) % 4]);
			}
		}
	}
}

TEST_METHOD(EventDatesDoNotShiftAndResumeTheSeasonAfterwards)
{
	using namespace std::chrono;
	auto themes = ui::theme::GetThemes();
	auto event = ui::theme::GetDefaultTheme();
	event.stableId = L"test.event";
	event.schedule = {ui::theme::ScheduleKind::AnnualDateRange, 10, 25, 10, 31, 10};
	themes.push_back(event);

	for (const auto southern : {false, true})
	{
		const auto baseline = southern ? L"season.spring" : L"season.autumn";
		Assert::IsTrue(
			ui::theme::FindScheduledTheme(themes, year{2024} / October / 24, southern)->stableId == baseline);
		for (auto day = 25; day <= 31; ++day)
		{
			Assert::IsTrue(
				ui::theme::FindScheduledTheme(themes, year{2024} / October / day, southern)->stableId ==
				event.stableId);
		}

		Assert::IsTrue(
			ui::theme::FindScheduledTheme(themes, year{2024} / November / 1, southern)->stableId == baseline);
	}
}

TEST_METHOD(OverlapsUsePriorityThenDurationThenIdNotCatalogOrder)
{
	using namespace std::chrono;
	auto season = ui::theme::GetDefaultTheme();
	season.stableId = L"test.season";
	season.schedule = {ui::theme::ScheduleKind::AnnualDateRange, 9, 1, 11, 30, 2};
	auto event = season;
	event.stableId = L"test.event";
	event.schedule = {ui::theme::ScheduleKind::AnnualDateRange, 10, 25, 10, 31, 1};
	std::vector<ui::theme::ThemeDefinition> themes{event, season};
	const auto date = year{2024} / October / 27;
	Assert::IsTrue(ui::theme::FindScheduledTheme(themes, date, false)->stableId == season.stableId);

	themes[0].schedule.priority = 2;
	Assert::IsTrue(ui::theme::FindScheduledTheme(themes, date, false)->stableId == event.stableId);

	auto tied = themes[0];
	tied.stableId = L"test.a";
	themes.push_back(tied);
	Assert::IsTrue(ui::theme::FindScheduledTheme(themes, date, false)->stableId == tied.stableId);
	std::reverse(themes.begin(), themes.end());
	Assert::IsTrue(ui::theme::FindScheduledTheme(themes, date, false)->stableId == tied.stableId);
}

TEST_METHOD(OverlapDurationUsesTheActualYearCrossingOccurrence)
{
	using namespace std::chrono;
	auto winter = ui::theme::GetDefaultTheme();
	winter.stableId = L"z.winter";
	winter.schedule = {ui::theme::ScheduleKind::AnnualDateRange, 12, 1, 3, 1};
	auto other = winter;
	other.stableId = L"a.other";
	other.schedule = {ui::theme::ScheduleKind::AnnualDateRange, 11, 1, 1, 31};
	const std::vector<ui::theme::ThemeDefinition> themes{winter, other};

	Assert::IsTrue(ui::theme::FindScheduledTheme(themes, year{2023} / January / 1, false)->stableId == winter.stableId);
	Assert::IsTrue(ui::theme::FindScheduledTheme(themes, year{2024} / January / 1, false)->stableId == other.stableId);
	Assert::IsNull(ui::theme::FindScheduledTheme(themes, year{2024} / July / 1, false));
}

TEST_METHOD(AutomaticSelectionHonorsEnablementAndLeavesManualFallbackIndependent)
{
	using namespace std::chrono;
	const auto date = year{2024} / July / 1;
	const auto& manual = ui::theme::GetDefaultTheme();
	Assert::IsTrue(&ui::theme::ResolveTheme(manual.stableId, true, false, date, false) == &manual);
	Assert::IsTrue(&ui::theme::ResolveTheme(L"season.spring", false, true, date, false) == &manual);
	Assert::IsTrue(ui::theme::ResolveTheme(manual.stableId, true, true, date, false).stableId == L"season.summer");
	Assert::IsTrue(ui::theme::ResolveTheme(manual.stableId, true, true, date, true).stableId == L"season.winter");
}

TEST_METHOD(CyclingUsesCatalogOrderAndWraps)
{
	const auto& themes = ui::theme::GetThemes();
	for (size_t index = 0; index < themes.size(); ++index)
	{
		Assert::IsTrue(
			ui::theme::GetNextTheme(themes[index].stableId).stableId == themes[(index + 1) % themes.size()].stableId);
		Assert::IsTrue(
			ui::theme::GetNextTheme(themes[index].stableId, true).stableId ==
			themes[(index + themes.size() - 1) % themes.size()].stableId);
		Assert::IsTrue(
			ui::theme::GetNextTheme(ui::theme::GetNextTheme(themes[index].stableId).stableId, true).stableId ==
			themes[index].stableId);
	}

	Assert::IsTrue(
		ui::theme::GetNextTheme(L"not-a-theme").stableId ==
		ui::theme::GetNextTheme(ui::theme::GetDefaultTheme().stableId).stableId);
	Assert::IsTrue(ui::theme::GetNextTheme(L"not-a-theme", true).stableId == themes.back().stableId);
	Assert::IsTrue(ui::theme::GetNextTheme(L"", true).stableId == themes.back().stableId);
}

TEST_METHOD(StandardSurfaceColorsPreserveTheOriginalAppearance)
{
	using ui::uiColor;
	const std::pair<uiColor, uiColor> surfaces[] = {
		{uiColor::TitleBarBackground, uiColor::Background},
		{uiColor::TitleBarText, uiColor::Text},
		{uiColor::TitleBarButton, uiColor::FrameSelected},
		{uiColor::TitleBarButtonHover, uiColor::FrameSelected},
		{uiColor::ColumnHeaderBackground, uiColor::Background},
		{uiColor::ColumnHeaderText, uiColor::Text},
		{uiColor::ColumnHeaderArrow, uiColor::Arrow},
		{uiColor::MenuBackground, uiColor::Background},
		{uiColor::MenuText, uiColor::Text},
		{uiColor::MenuTextDisabled, uiColor::TextDisabled}};
	const auto& standard = ui::theme::GetDefaultTheme();
	Assert::IsTrue(standard.stableId == L"standard");
	for (const auto& palette : {standard.light, standard.dark})
	{
		for (const auto& [surface, original] : surfaces)
		{
			const auto& actual = palette.values[static_cast<size_t>(surface)];
			const auto& expected = palette.values[static_cast<size_t>(original)];
			Assert::AreEqual(expected.value, actual.value);
			Assert::AreEqual(expected.system, actual.system);
			Assert::AreEqual(expected.highContrast, actual.highContrast);
		}
	}
}

TEST_METHOD(UnknownIdsAreNotFound)
{
	Assert::IsNull(ui::theme::FindTheme(L"not-a-theme"));
	Assert::IsNull(ui::theme::FindTheme(L""));
}

TEST_METHOD(StatusRibbonIsOptInAndHiddenInHighContrast)
{
	auto theme = ui::theme::GetDefaultTheme();
	Assert::IsTrue(theme.statusRibbon.empty());
	Assert::IsTrue(ui::theme::GetStatusRibbon(theme, ui::uiTheme::Light).empty());
	theme.statusRibbon = {RGB(255, 0, 0), RGB(0, 255, 0), RGB(0, 0, 255)};
	for (const auto appearance : {ui::uiTheme::Light, ui::uiTheme::Dark})
	{
		const auto colors = ui::theme::GetStatusRibbon(theme, appearance);
		Assert::AreEqual(size_t{3}, colors.size());
		Assert::AreEqual(theme.statusRibbon[0], colors[0]);
		Assert::AreEqual(theme.statusRibbon[2], colors[2]);
	}

	Assert::IsTrue(ui::theme::GetStatusRibbon(theme, ui::uiTheme::HighContrast).empty());
}

TEST_METHOD(InvalidColorRolesAndAppearancesAreRejected)
{
	auto color = COLORREF{};
	Assert::IsFalse(ui::theme::TryGetColor(
		ui::theme::GetDefaultTheme(), ui::theme::Appearance::Light, static_cast<ui::theme::ColorRole>(99), color));
	Assert::IsFalse(ui::theme::TryGetColor(
		ui::theme::GetDefaultTheme(), static_cast<ui::theme::Appearance>(99), ui::theme::ColorRole::Accent, color));
}
}
;
} // namespace themetest
