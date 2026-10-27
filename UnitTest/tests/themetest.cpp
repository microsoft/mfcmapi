#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <UI/Themes/Theme.h>

namespace themetest
{
	TEST_CLASS(ThemeTest)
	{
	public:
		TEST_METHOD(CatalogIdsAreUniqueAndResolveToTheirDefinitions)
		{
			const auto& themes = ui::theme::GetThemes();
			Assert::IsFalse(themes.empty());

			for (const auto& theme : themes)
			{
				Assert::IsFalse(theme.stableId.empty());
				Assert::IsTrue(ui::theme::FindTheme(theme.stableId) == &theme);
			}
		}

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
				Assert::AreEqual(GetSysColor(COLOR_HIGHLIGHT),
								 ui::theme::GetColor(theme, ui::uiTheme::HighContrast, role));
			}
		}

		TEST_METHOD(CyclingUsesCatalogOrderAndWraps)
		{
			const auto& themes = ui::theme::GetThemes();
			for (size_t index = 0; index < themes.size(); ++index)
			{
				Assert::IsTrue(ui::theme::GetNextTheme(themes[index].stableId).stableId ==
							   themes[(index + 1) % themes.size()].stableId);
			}

			Assert::IsTrue(ui::theme::GetNextTheme(L"not-a-theme").stableId ==
						   ui::theme::GetNextTheme(ui::theme::GetDefaultTheme().stableId).stableId);
		}

		TEST_METHOD(UnknownIdsAreNotFound)
		{
			Assert::IsNull(ui::theme::FindTheme(L"not-a-theme"));
			Assert::IsNull(ui::theme::FindTheme(L""));
		}

		TEST_METHOD(InvalidColorRolesAndAppearancesAreRejected)
		{
			auto color = COLORREF{};
			Assert::IsFalse(ui::theme::TryGetColor(ui::theme::GetDefaultTheme(),
												   ui::theme::Appearance::Light,
												   static_cast<ui::theme::ColorRole>(99),
												   color));
			Assert::IsFalse(ui::theme::TryGetColor(ui::theme::GetDefaultTheme(),
												   static_cast<ui::theme::Appearance>(99),
												   ui::theme::ColorRole::Accent,
												   color));
		}
	};
} // namespace themetest
