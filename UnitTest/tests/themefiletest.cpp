#include <UnitTest/stdafx.h>
#include <UnitTest/UnitTest.h>
#include <UI/Themes/ThemeFile.h>
#include <filesystem>
#include <fstream>

namespace themefiletest
{
	static std::filesystem::path ExampleDirectory()
	{
		return std::filesystem::path{__FILE__}.parent_path().parent_path().parent_path() / L"docs" / L"Themes";
	}

	static std::wstring ReadExample()
	{
		std::ifstream stream(ExampleDirectory() / L"Halloween.json", std::ios::binary);
		Assert::IsTrue(stream.good());

		const auto bytes = std::string{std::istreambuf_iterator<char>{stream}, std::istreambuf_iterator<char>{}};

		return {bytes.begin(), bytes.end()};
	}

	static void Replace(std::wstring& text, const std::wstring_view before, const std::wstring_view after)
	{
		const auto position = text.find(before);
		Assert::IsTrue(position != std::wstring::npos);

		text.replace(position, before.size(), after);
	}

	static double RelativeLuminance(const COLORREF color)
	{
		const auto linearize = [](const BYTE channel) {
			const auto value = static_cast<double>(channel) / 255.0;
			return value <= 0.04045 ? value / 12.92 : std::pow((value + 0.055) / 1.055, 2.4);
		};

		return 0.2126 * linearize(GetRValue(color)) + 0.7152 * linearize(GetGValue(color)) +
			   0.0722 * linearize(GetBValue(color));
	}

	static double ContrastRatio(const COLORREF first, const COLORREF second)
	{
		const auto firstLuminance = RelativeLuminance(first);
		const auto secondLuminance = RelativeLuminance(second);

		const auto brighter = firstLuminance > secondLuminance ? firstLuminance : secondLuminance;
		const auto darker = firstLuminance < secondLuminance ? firstLuminance : secondLuminance;

		return (brighter + 0.05) / (darker + 0.05);
	}

	static void WriteFile(const std::filesystem::path& path, const std::wstring& text, const bool bom = false)
	{
		const auto length = WideCharToMultiByte(
			CP_UTF8, WC_ERR_INVALID_CHARS, text.data(), static_cast<int>(text.size()), nullptr, 0, nullptr, nullptr);
		Assert::IsTrue(length > 0);

		auto bytes = std::string(static_cast<size_t>(length), '\0');
		Assert::AreEqual(
			length,
			WideCharToMultiByte(
				CP_UTF8,
				WC_ERR_INVALID_CHARS,
				text.data(),
				static_cast<int>(text.size()),
				bytes.data(),
				length,
				nullptr,
				nullptr));

		std::ofstream stream(path, std::ios::binary);
		Assert::IsTrue(stream.good());
		if (bom) stream.write("\xEF\xBB\xBF", 3);

		stream.write(bytes.data(), static_cast<std::streamsize>(bytes.size()));
		Assert::IsTrue(stream.good());
	}

	TEST_CLASS(ThemeFileTest)
	{
	public:
		static const bool dummy_var = true;

		TEST_METHOD_INITIALIZE(CreateDirectory)
		{
			wchar_t temporary[MAX_PATH]{};
			Assert::IsTrue(GetTempPathW(_countof(temporary), temporary) != 0);

			auto guid = GUID{};
			Assert::AreEqual(S_OK, CoCreateGuid(&guid));

			wchar_t id[40]{};
			Assert::IsTrue(StringFromGUID2(guid, id, _countof(id)) != 0);
			m_directory = std::filesystem::path{temporary} / (std::wstring(L"MFCMAPI-theme-test-") + id);
			Assert::IsTrue(std::filesystem::create_directory(m_directory));
		}

		TEST_METHOD_CLEANUP(RemoveDirectory) { std::filesystem::remove_all(m_directory); }

		TEST_METHOD(EveryThemeFileLoadsAndMeetsContrastTargets)
		{
			size_t fileCount = 0;
			for (const auto& entry : std::filesystem::directory_iterator(ExampleDirectory()))
			{
				if (entry.is_regular_file() && entry.path().extension() == L".json") ++fileCount;
			}

			Assert::IsTrue(fileCount > 0);

			const auto catalog = ui::theme::LoadThemeDirectory(ExampleDirectory().wstring());
			auto errors = std::wstring{};
			for (const auto& error : catalog.errors)
			{
				errors += error.path + L": " + error.message + L"\n";
			}
			Assert::IsTrue(catalog.errors.empty(), errors.c_str());
			Assert::AreEqual(fileCount, catalog.themes.size() - ui::theme::GetThemes().size());

			struct ContrastPair
			{
				ui::uiColor foreground;
				ui::uiColor background;
				std::wstring_view description;
				double minimum;
			};
			const ContrastPair textPairs[] = {
				{ui::uiColor::Text, ui::uiColor::Background, L"text/background", 4.5},
				{ui::uiColor::TitleBarText, ui::uiColor::TitleBarBackground, L"title bar text/background", 4.5},
				{ui::uiColor::ColumnHeaderText,
				 ui::uiColor::ColumnHeaderBackground,
				 L"column header text/background",
				 4.5},
				{ui::uiColor::MenuText, ui::uiColor::MenuBackground, L"menu text/background", 4.5},
				{ui::uiColor::MenuTextDisabled,
				 ui::uiColor::MenuBackground,
				 L"disabled menu text/menu background",
				 3.0},
				{ui::uiColor::TextDisabled, ui::uiColor::Background, L"disabled menu text/background", 3.0},
				{ui::uiColor::TextDisabledSelected,
				 ui::uiColor::DisabledSelectedBackground,
				 L"disabled selected menu text/background",
				 4.5},
				{ui::uiColor::TextReadOnly, ui::uiColor::BackgroundReadOnly, L"read-only text/background", 4.5},
				{ui::uiColor::GlowText, ui::uiColor::SelectedBackground, L"selected text/background", 4.5},
				{ui::uiColor::GlowText, ui::uiColor::GlowBackground, L"hover text/background", 4.5},
				{ui::uiColor::StatusText, ui::uiColor::Status, L"status text/background", 4.5},
				{ui::uiColor::PaneHeaderText, ui::uiColor::PaneHeaderBackground, L"pane header text/background", 4.5},
				{ui::uiColor::TextHighlight, ui::uiColor::TextHighlightBackground, L"highlight text/background", 4.5}};
			const ContrastPair nonTextPairs[] = {
				{ui::uiColor::TitleBarButton, ui::uiColor::TitleBarBackground, L"title bar button/background", 3.0},
				{ui::uiColor::TitleBarButtonHover,
				 ui::uiColor::GlowBackground,
				 L"title bar button/hover background",
				 3.0},
				{ui::uiColor::ColumnHeaderArrow,
				 ui::uiColor::ColumnHeaderBackground,
				 L"column header arrow/background",
				 3.0},
				{ui::uiColor::FrameSelected, ui::uiColor::Background, L"focus indicator/background", 3.0},
				{ui::uiColor::GlowBackground, ui::uiColor::Background, L"hover/background", 3.0},
				{ui::uiColor::SelectedBackground, ui::uiColor::Background, L"selection/background", 3.0},
				{ui::uiColor::TextHighlightBackground, ui::uiColor::Background, L"text highlight/background", 3.0}};

			auto violations = std::wostringstream{};
			for (auto index = ui::theme::GetThemes().size(); index < catalog.themes.size(); ++index)
			{
				const auto& theme = catalog.themes[index];
				for (const auto appearance : {ui::uiTheme::Light, ui::uiTheme::Dark})
				{
					const auto& colors = appearance == ui::uiTheme::Light ? theme.light : theme.dark;
					const auto appearanceName = appearance == ui::uiTheme::Light ? L"light" : L"dark";
					const auto checkPairs = [&](const auto& pairs) {
						for (const auto& pair : pairs)
						{
							const auto ratio = ContrastRatio(
								colors.values[static_cast<size_t>(pair.foreground)].value,
								colors.values[static_cast<size_t>(pair.background)].value);
							if (ratio < pair.minimum)
							{
								violations << theme.name << L" (" << appearanceName << L"): " << pair.description
										   << L" = " << ratio << L":1; requires " << pair.minimum << L":1\n";
							}
						}
					};

					checkPairs(textPairs);
					checkPairs(nonTextPairs);

					const auto normalTextDifference = ContrastRatio(
						colors.values[static_cast<size_t>(ui::uiColor::MenuTextDisabled)].value,
						colors.values[static_cast<size_t>(ui::uiColor::MenuText)].value);
					if (normalTextDifference < 3.0)
					{
						violations << theme.name << L" (" << appearanceName
								   << L"): enabled and disabled menu text are too similar (" << normalTextDifference
								   << L":1; requires 3:1)\n";
					}

					const auto selectedTextDifference = ContrastRatio(
						colors.values[static_cast<size_t>(ui::uiColor::TextDisabledSelected)].value,
						colors.values[static_cast<size_t>(ui::uiColor::GlowText)].value);
					if (selectedTextDifference < 3.0)
					{
						violations << theme.name << L" (" << appearanceName
								   << L"): enabled and disabled selected menu text are too similar ("
								   << selectedTextDifference << L":1; requires 3:1)\n";
					}

					if (colors.values[static_cast<size_t>(ui::uiColor::MenuTextDisabled)].highContrast ==
							colors.values[static_cast<size_t>(ui::uiColor::MenuText)].highContrast ||
						colors.values[static_cast<size_t>(ui::uiColor::TextDisabledSelected)].highContrast ==
							colors.values[static_cast<size_t>(ui::uiColor::GlowText)].highContrast)
					{
						violations << theme.name << L" (" << appearanceName
								   << L"): enabled and disabled menu text use the same High Contrast color\n";
					}
				}
			}

			Assert::IsTrue(violations.str().empty(), violations.str().c_str());
		}

		TEST_METHOD(BuiltInSelectedDisabledTextMeetsContrastTarget)
		{
			auto violations = std::wostringstream{};
			for (const auto& theme : ui::theme::GetThemes())
			{
				for (const auto appearance : {ui::uiTheme::Light, ui::uiTheme::Dark})
				{
					const auto& colors = appearance == ui::uiTheme::Light ? theme.light : theme.dark;
					const auto ratio = ContrastRatio(
						colors.values[static_cast<size_t>(ui::uiColor::TextDisabledSelected)].value,
						colors.values[static_cast<size_t>(ui::uiColor::DisabledSelectedBackground)].value);
					const auto borderContrast = ContrastRatio(
						colors.values[static_cast<size_t>(ui::uiColor::DisabledSelectedBorder)].value,
						colors.values[static_cast<size_t>(ui::uiColor::DisabledSelectedBackground)].value);
					const auto selectedDifference = ContrastRatio(
						colors.values[static_cast<size_t>(ui::uiColor::TextDisabledSelected)].value,
						colors.values[static_cast<size_t>(ui::uiColor::GlowText)].value);
					const auto normalDifference = ContrastRatio(
						colors.values[static_cast<size_t>(ui::uiColor::TextDisabled)].value,
						colors.values[static_cast<size_t>(ui::uiColor::Text)].value);
					if (ratio < 4.5 || borderContrast < 3.0 || selectedDifference < 3.0 || normalDifference < 3.0 ||
						colors.values[static_cast<size_t>(ui::uiColor::TextDisabled)].highContrast ==
							colors.values[static_cast<size_t>(ui::uiColor::Text)].highContrast ||
						colors.values[static_cast<size_t>(ui::uiColor::TextDisabledSelected)].highContrast ==
							colors.values[static_cast<size_t>(ui::uiColor::GlowText)].highContrast)
					{
						violations << theme.name << L" (" << (appearance == ui::uiTheme::Light ? L"light" : L"dark")
								   << L"): disabled text/fill " << ratio << L":1, border/fill " << borderContrast
								   << L":1, selected text difference " << selectedDifference
								   << L":1, normal text difference " << normalDifference << L":1\n";
					}
				}
			}

			Assert::IsTrue(violations.str().empty(), violations.str().c_str());
		}

		TEST_METHOD(StatusRibbonParsesLiteralColorsAndCanBeDisabled)
		{
			for (const auto value : {L"[]", L"[\"red\"]", L"[\"red\", \"#00FF00\", \"BLUE\"]"})
			{
				const auto json = L"{\"schemaVersion\":1,\"id\":\"test.ribbon\",\"name\":\"Ribbon test\","
								  L"\"colors\":{\"light\":{},\"dark\":{}},\"statusRibbon\":" +
								  std::wstring(value) + L"}";
				auto theme = ui::theme::ThemeDefinition{};
				auto error = std::wstring{};
				Assert::IsTrue(ui::theme::ParseThemeFile(json, theme, error), error.c_str());
				const auto count = std::wstring_view(value) == L"[]"		  ? 0u
								   : std::wstring_view(value) == L"[\"red\"]" ? 1u
																			  : 3u;
				Assert::AreEqual(static_cast<size_t>(count), theme.statusRibbon.size());
				if (count)
				{
					Assert::AreEqual(RGB(255, 0, 0), theme.statusRibbon.front());
					if (count == 3)
					{
						Assert::AreEqual(RGB(0, 255, 0), theme.statusRibbon[1]);
						Assert::AreEqual(RGB(0, 0, 255), theme.statusRibbon[2]);
					}
				}
			}
		}

		TEST_METHOD(StatusBarStylesParseAndEmptyColorsDisableDecoration)
		{
			for (const auto& [name, style] :
				 {std::pair{L"ribbon", ui::theme::StatusBarStyle::Ribbon},
				  std::pair{L"bands", ui::theme::StatusBarStyle::Bands},
				  std::pair{L"gradient", ui::theme::StatusBarStyle::Gradient}})
			{
				for (const auto colors : {L"[]", L"[\"red\"]", L"[\"red\", \"#00FF00\", \"BLUE\"]"})
				{
					const auto json = L"{\"schemaVersion\":1,\"id\":\"test.status\",\"name\":\"Status test\","
									  L"\"colors\":{\"light\":{},\"dark\":{}},\"statusBar\":{\"style\":\"" +
									  std::wstring(name) + L"\",\"colors\":" + colors + L"}}";
					auto definition = ui::theme::GetDefaultTheme();
					auto error = std::wstring{};
					Assert::IsTrue(ui::theme::ParseThemeFile(json, definition, error), error.c_str());
					Assert::IsTrue(definition.statusBarStyle == style);
					const auto count = std::wstring_view(colors) == L"[]"		   ? 0u
									   : std::wstring_view(colors) == L"[\"red\"]" ? 1u
																				   : 3u;
					Assert::AreEqual(static_cast<size_t>(count), definition.statusRibbon.size());
					if (count)
					{
						Assert::AreEqual(RGB(255, 0, 0), definition.statusRibbon[0]);
						if (count == 3)
						{
							Assert::AreEqual(RGB(0, 255, 0), definition.statusRibbon[1]);
							Assert::AreEqual(RGB(0, 0, 255), definition.statusRibbon[2]);
						}
					}
					Assert::IsTrue(ui::theme::GetStatusRibbon(definition, ui::uiTheme::HighContrast).empty());
				}
			}
		}

		TEST_METHOD(BuiltInSurfaceAccentsMeetContrastTargets)
		{
			for (const auto& theme : ui::theme::GetThemes())
			{
				for (const auto& palette : {theme.light, theme.dark})
				{
					const auto check =
						[&](const ui::uiColor foreground, const ui::uiColor background, const double minimum) {
							Assert::IsTrue(
								ContrastRatio(
									palette.values[static_cast<size_t>(foreground)].value,
									palette.values[static_cast<size_t>(background)].value) >= minimum,
								theme.name.c_str());
						};
					check(ui::uiColor::TitleBarText, ui::uiColor::TitleBarBackground, 4.5);
					check(ui::uiColor::TitleBarButton, ui::uiColor::TitleBarBackground, 3.0);
					check(ui::uiColor::TitleBarButtonHover, ui::uiColor::GlowBackground, 3.0);
					check(ui::uiColor::ColumnHeaderText, ui::uiColor::ColumnHeaderBackground, 4.5);
					check(ui::uiColor::ColumnHeaderArrow, ui::uiColor::ColumnHeaderBackground, 3.0);
					check(ui::uiColor::MenuText, ui::uiColor::MenuBackground, 4.5);
					check(ui::uiColor::MenuTextDisabled, ui::uiColor::MenuBackground, 3.0);
				}
			}
		}

		TEST_METHOD(ExamplesLoadAndEventsOverrideSeasonsWithoutShifting)
		{
			using namespace std::chrono;

			for (const auto filename : {L"Halloween.json", L"Christmas.json"})
			{
				std::filesystem::copy_file(ExampleDirectory() / filename, m_directory / filename);
			}

			const auto catalog = ui::theme::LoadThemeDirectory(m_directory.wstring());
			Assert::IsTrue(catalog.errors.empty());

			for (const auto southern : {false, true})
			{
				const auto halloween =
					ui::theme::FindScheduledTheme(catalog.themes, year{2024} / October / 31, southern);
				const auto christmas =
					ui::theme::FindScheduledTheme(catalog.themes, year{2024} / December / 25, southern);
				Assert::IsNotNull(halloween);
				Assert::IsNotNull(christmas);
				Assert::IsTrue(halloween->stableId == L"example.halloween");
				Assert::IsTrue(christmas->stableId == L"example.christmas");

				const auto next = ui::theme::FindScheduledTheme(catalog.themes, year{2024} / November / 1, southern);
				Assert::IsNotNull(next);
				Assert::IsTrue(next->stableId == (southern ? L"season.spring" : L"season.autumn"));
			}
		}

		TEST_METHOD(ThemeFilesAreSortedByDisplayName)
		{
			auto alphabeticallyLater = ReadExample();
			Replace(alphabeticallyLater, L"example.halloween", L"theme.zulu");
			Replace(alphabeticallyLater, L"Halloween", L"Zulu");
			WriteFile(m_directory / L"a.json", alphabeticallyLater);

			auto alphabeticallyEarlier = ReadExample();
			Replace(alphabeticallyEarlier, L"example.halloween", L"theme.alpha");
			Replace(alphabeticallyEarlier, L"Halloween", L"Alpha");
			WriteFile(m_directory / L"z.json", alphabeticallyEarlier);

			const auto catalog = ui::theme::LoadThemeDirectory(m_directory.wstring());
			Assert::IsTrue(catalog.errors.empty());
			Assert::AreEqual(std::wstring{L"Alpha"}, catalog.themes[catalog.themes.size() - 2].name);
			Assert::AreEqual(std::wstring{L"Zulu"}, catalog.themes.back().name);
		}

		TEST_METHOD(ColorsParseHexChannelsCssAliasesAndSystemMappings)
		{
			auto json = ReadExample();
			Replace(json, L"OldLace", L"#1234ab");
			Replace(json, L"Wheat", L"rEbEcCaPuRpLe");
			Replace(json, L"DarkOrange", L"darkGREY");
			Replace(
				json,
				L"\"Text\": { \"value\": \"Black\", \"highContrast\": 8 }",
				L"\"Text\": { \"value\": \"Black\", \"system\": 0, \"highContrast\": 13 }");

			auto theme = ui::theme::ThemeDefinition{};
			auto error = std::wstring{};
			Assert::IsTrue(ui::theme::ParseThemeFile(json, theme, error), error.c_str());
			Assert::AreEqual(
				RGB(0x12, 0x34, 0xAB), ui::theme::GetColor(theme, ui::uiTheme::Light, ui::uiColor::Background));
			Assert::AreEqual(
				RGB(0x66, 0x33, 0x99), ui::theme::GetColor(theme, ui::uiTheme::Light, ui::uiColor::BackgroundReadOnly));
			Assert::AreEqual(RGB(0xA9, 0xA9, 0xA9), ui::theme::GetColor(theme, ui::uiTheme::Light, ui::uiColor::Glow));
			Assert::AreEqual(GetSysColor(0), ui::theme::GetColor(theme, ui::uiTheme::Light, ui::uiColor::Text));
			Assert::AreEqual(GetSysColor(13), ui::theme::GetColor(theme, ui::uiTheme::HighContrast, ui::uiColor::Text));
		}

		TEST_METHOD(OmittedColorsInheritTheMatchingStandardPalette)
		{
			const auto json = LR"({
				"schemaVersion": 1,
				"id": "my.partial",
				"name": "Partial",
				"colors": {
					"light": {
						"Text": { "value": "#123456", "highContrast": 8 },
						"SelectedBackground": { "value": "#111111", "highContrast": 13 }
					},
					"dark": {
						"SelectedBackground": { "value": "#FFFFFF", "highContrast": 13 }
					}
				}
			})";
			auto theme = ui::theme::ThemeDefinition{};
			auto error = std::wstring{};
			Assert::IsTrue(ui::theme::ParseThemeFile(json, theme, error), error.c_str());

			const auto& standard = ui::theme::GetDefaultTheme();
			for (size_t index = 0; index < static_cast<size_t>(ui::uiColor::UIEnd); ++index)
			{
				const auto& dark = theme.dark.values[index];
				const auto& expectedDark = standard.dark.values[index];
				const auto& light = theme.light.values[index];
				if (index == static_cast<size_t>(ui::uiColor::SelectedBackground))
				{
					Assert::AreEqual(RGB(0xFF, 0xFF, 0xFF), dark.value);
					Assert::AreEqual(-1, dark.system);
					Assert::AreEqual(13, dark.highContrast);
					Assert::AreEqual(RGB(0x11, 0x11, 0x11), light.value);
					Assert::AreEqual(-1, light.system);
					Assert::AreEqual(13, light.highContrast);
				}
				else
				{
					Assert::AreEqual(expectedDark.value, dark.value);
					Assert::AreEqual(expectedDark.system, dark.system);
					Assert::AreEqual(expectedDark.highContrast, dark.highContrast);

					if (index == static_cast<size_t>(ui::uiColor::Text))
					{
						Assert::AreEqual(RGB(0x12, 0x34, 0x56), light.value);
						Assert::AreEqual(-1, light.system);
						Assert::AreEqual(8, light.highContrast);
					}
					else
					{
						const auto& expectedLight = standard.light.values[index];
						Assert::AreEqual(expectedLight.value, light.value);
						Assert::AreEqual(expectedLight.system, light.system);
						Assert::AreEqual(expectedLight.highContrast, light.highContrast);
					}
				}
			}
		}

		TEST_METHOD(SurfaceColorsCanBeSharedOrDifferent)
		{
			for (const auto title : {L"#123456", L"#ABCDEF"})
			{
				const auto palette = L"{\"Status\":{\"value\":\"#123456\",\"highContrast\":13},"
									 L"\"TitleBarBackground\":{\"value\":\"" +
									 std::wstring(title) + L"\",\"highContrast\":5}}";
				const auto json = L"{\"schemaVersion\":1,\"id\":\"test.surfaces\",\"name\":\"Surface test\","
								  L"\"colors\":{\"light\":" +
								  palette + L",\"dark\":" + palette + L"}}";
				auto theme = ui::theme::ThemeDefinition{};
				auto error = std::wstring{};
				Assert::IsTrue(ui::theme::ParseThemeFile(json, theme, error), error.c_str());
				const auto expectedTitle =
					std::wstring_view(title) == L"#123456" ? RGB(0x12, 0x34, 0x56) : RGB(0xAB, 0xCD, 0xEF);
				for (const auto appearance : {ui::uiTheme::Light, ui::uiTheme::Dark})
				{
					Assert::AreEqual(
						expectedTitle, ui::theme::GetColor(theme, appearance, ui::uiColor::TitleBarBackground));
					Assert::AreEqual(
						RGB(0x12, 0x34, 0x56), ui::theme::GetColor(theme, appearance, ui::uiColor::Status));
				}
			}
		}

		TEST_METHOD(SurfaceRolesNeverInheritFromOtherCustomizedRoles)
		{
			const auto json = LR"({
				"schemaVersion": 1,
				"id": "my.surfaces",
				"name": "Independent surfaces",
				"colors": {
					"light": {
						"Background": { "value": "#123456", "system": 0, "highContrast": 13 },
						"Text": { "value": "#654321", "highContrast": 14 },
						"FrameSelected": { "value": "#ABCDEF", "highContrast": 14 },
						"Arrow": { "value": "#987654", "highContrast": 14 },
						"TextDisabled": { "value": "#123123", "highContrast": 14 },
						"TitleBarBackground": { "value": "#224466", "highContrast": 15 }
					},
					"dark": {
						"Background": { "value": "#654321", "highContrast": 13 },
						"Text": { "value": "#123456", "highContrast": 14 },
						"FrameSelected": { "value": "#FEDCBA", "highContrast": 14 },
						"Arrow": { "value": "#456789", "highContrast": 14 },
						"TextDisabled": { "value": "#321321", "highContrast": 14 },
						"ColumnHeaderBackground": { "value": "#446688", "highContrast": 15 }
					}
				}
			})";
			auto theme = ui::theme::ThemeDefinition{};
			auto error = std::wstring{};
			Assert::IsTrue(ui::theme::ParseThemeFile(json, theme, error), error.c_str());
			const auto& standard = ui::theme::GetDefaultTheme();
			for (const auto dark : {false, true})
			{
				const auto& palette = dark ? theme.dark : theme.light;
				const auto& defaults = dark ? standard.dark : standard.light;
				for (auto index = static_cast<size_t>(ui::uiColor::TitleBarBackground);
					 index < static_cast<size_t>(ui::uiColor::UIEnd);
					 ++index)
				{
					const auto& color = palette.values[index];
					if (index == static_cast<size_t>(
									 dark ? ui::uiColor::ColumnHeaderBackground : ui::uiColor::TitleBarBackground))
					{
						Assert::AreEqual(dark ? RGB(0x44, 0x66, 0x88) : RGB(0x22, 0x44, 0x66), color.value);
						Assert::AreEqual(-1, color.system);
						Assert::AreEqual(15, color.highContrast);
					}
					else
					{
						Assert::AreEqual(defaults.values[index].value, color.value);
						Assert::AreEqual(defaults.values[index].system, color.system);
						Assert::AreEqual(defaults.values[index].highContrast, color.highContrast);
					}
				}
			}
		}

		TEST_METHOD(DisabledSelectedColorsCanBeCustomized)
		{
			const auto json = LR"({
				"schemaVersion": 1,
				"id": "my.disabled-selected",
				"name": "Disabled Selected",
				"colors": {
					"light": {
						"GlowText": { "value": "#123456", "highContrast": 14 },
						"TextDisabledSelected": { "value": "#ABCDEF", "highContrast": 14 },
						"DisabledSelectedBackground": { "value": "#F0F0F0", "highContrast": 15 },
						"DisabledSelectedBorder": { "value": "#334455", "highContrast": 14 }
					},
					"dark": {
						"GlowText": { "value": "#654321", "highContrast": 14 },
						"TextDisabledSelected": { "value": "#FEDCBA", "highContrast": 14 },
						"DisabledSelectedBackground": { "value": "#101010", "highContrast": 15 },
						"DisabledSelectedBorder": { "value": "#CCDDEE", "highContrast": 14 }
					}
				}
			})";
			auto theme = ui::theme::ThemeDefinition{};
			auto error = std::wstring{};
			Assert::IsTrue(ui::theme::ParseThemeFile(json, theme, error), error.c_str());

			Assert::AreEqual(
				RGB(0xAB, 0xCD, 0xEF),
				theme.light.values[static_cast<size_t>(ui::uiColor::TextDisabledSelected)].value);
			Assert::AreEqual(
				14, theme.light.values[static_cast<size_t>(ui::uiColor::TextDisabledSelected)].highContrast);
			Assert::AreEqual(
				RGB(0xFE, 0xDC, 0xBA), theme.dark.values[static_cast<size_t>(ui::uiColor::TextDisabledSelected)].value);
			Assert::AreEqual(
				RGB(0xF0, 0xF0, 0xF0),
				theme.light.values[static_cast<size_t>(ui::uiColor::DisabledSelectedBackground)].value);
			Assert::AreEqual(
				RGB(0xCC, 0xDD, 0xEE),
				theme.dark.values[static_cast<size_t>(ui::uiColor::DisabledSelectedBorder)].value);
		}

		TEST_METHOD(EmptyPalettesInheritAllStandardColors)
		{
			const auto json = LR"({
				"schemaVersion": 1,
				"id": "my.missing-disabled-selected",
				"name": "Missing Disabled Selected",
				"colors": {
					"light": {},
					"dark": {}
				}
			})";
			auto theme = ui::theme::ThemeDefinition{};
			auto error = std::wstring{};
			Assert::IsTrue(ui::theme::ParseThemeFile(json, theme, error), error.c_str());

			const auto& standard = ui::theme::GetDefaultTheme();
			for (const auto appearance : {ui::uiTheme::Light, ui::uiTheme::Dark})
			{
				const auto& actual = appearance == ui::uiTheme::Light ? theme.light : theme.dark;
				const auto& expected = appearance == ui::uiTheme::Light ? standard.light : standard.dark;
				for (size_t index = 0; index < static_cast<size_t>(ui::uiColor::UIEnd); ++index)
				{
					Assert::AreEqual(expected.values[index].value, actual.values[index].value);
					Assert::AreEqual(expected.values[index].system, actual.values[index].system);
					Assert::AreEqual(expected.values[index].highContrast, actual.values[index].highContrast);
				}
			}
		}

		TEST_METHOD(InvalidDefinitionsReportErrorsWithoutChangingTheOutput)
		{
			const auto valid = ReadExample();
			const std::pair<std::wstring_view, std::wstring_view> invalid[] = {
				{L"\"schemaVersion\": 1",
				 L"\"schemaVersion\": 1, \"statusBar\": {\"style\":\"unknown\",\"colors\":[\"red\"]}"},
				{L"\"schemaVersion\": 1", L"\"schemaVersion\": 1, \"statusBar\": {\"colors\":[\"red\"]}"},
				{L"\"schemaVersion\": 1", L"\"schemaVersion\": 1, \"statusBar\": {\"style\":\"gradient\"}"},
				{L"\"schemaVersion\": 1",
				 L"\"schemaVersion\": 1, \"statusBar\": {\"style\":\"gradient\",\"colors\":[\"Text\"]}"},
				{L"\"schemaVersion\": 1",
				 L"\"schemaVersion\": 1, \"statusBar\": {\"style\":\"gradient\",\"colors\":[7]}"},
				{L"\"schemaVersion\": 1",
				 L"\"schemaVersion\": 1, \"statusBar\": {\"style\":\"gradient\",\"colors\":{}}"},
				{L"\"schemaVersion\": 1",
				 L"\"schemaVersion\": 1, \"statusBar\": {\"style\":\"gradient\",\"colors\":[],\"unknown\":true}"},
				{L"\"schemaVersion\": 1",
				 L"\"schemaVersion\": 1, \"statusBar\": {\"style\":\"gradient\",\"colors\":[]},\"statusRibbon\":[]"},
				{L"\"schemaVersion\": 1",
				 L"\"schemaVersion\": 1, \"statusBar\": "
				 L"{\"style\":\"gradient\",\"colors\":[\"red\",\"red\",\"red\",\"red\",\"red\",\"red\",\"red\",\"red\","
				 L"\"red\",\"red\",\"red\",\"red\",\"red\",\"red\",\"red\",\"red\",\"red\"]}"},
				{L"\"schemaVersion\": 1", L"\"schemaVersion\": 1, \"statusRibbon\": [\"not-a-color\"]"},
				{L"\"schemaVersion\": 1", L"\"schemaVersion\": 1, \"statusRibbon\": [\"#12345Z\"]"},
				{L"\"schemaVersion\": 1", L"\"schemaVersion\": 1, \"statusRibbon\": [7]"},
				{L"\"schemaVersion\": 1", L"\"schemaVersion\": 1, \"statusRibbon\": {}"},
				{L"\"schemaVersion\": 1",
				 L"\"schemaVersion\": 1, \"statusRibbon\": "
				 L"[\"red\",\"red\",\"red\",\"red\",\"red\",\"red\",\"red\",\"red\",\"red\",\"red\",\"red\",\"red\","
				 L"\"red\",\"red\",\"red\",\"red\",\"red\"]"},
				{L"\"schemaVersion\": 1", L"\"schemaVersion\": 2"},
				{L"\"schemaVersion\": 1", L"\"schemaVersion\": 1.5"},
				{L"\"name\": \"Halloween\"", L"\"name\": \"\""},
				{L"OldLace", L"not-a-color"},
				{L"OldLace", L"#12345Z"},
				{L"OldLace", L"#1234"},
				{L"\"highContrast\": 5", L"\"highContrast\": -1"},
				{L"\"highContrast\": 5", L"\"highContrast\": 25"},
				{L"\"highContrast\": 5", L"\"highContrast\": 32"},
				{L"\"highContrast\": 5", L"\"highContrast\": 5.5"},
				{L"\"Background\"", L"\"UnknownRole\""},
				{L"\"start\": \"10-25\"", L"\"start\": \"02-30\""},
				{L"\"kind\": \"dateRange\"", L"\"kind\": \"unknown\""},
				{L"\"priority\": 10", L"\"priority\": 0.5"},
				{L"\"seasonal\": false", L"\"seasonal\": \"false\""},
				{L"\"name\": \"Halloween\"", L"\"name\": \"Halloween\", \"category\": \"Holiday\""},
				{L"\"name\": \"Halloween\"", L"\"name\": \"Halloween\", \"variant\": \"Pumpkin\""},
				{L"\"dark\": {", L"\"missing\": {"}};

			auto theme = ui::theme::GetDefaultTheme();
			const auto originalId = theme.stableId;
			for (const auto& change : invalid)
			{
				auto json = valid;
				Replace(json, change.first, change.second);

				auto error = std::wstring{};
				Assert::IsFalse(ui::theme::ParseThemeFile(json, theme, error));
				Assert::IsFalse(error.empty());
				Assert::IsTrue(theme.stableId == originalId);
			}

			for (const auto text : {L"{broken", L"random characters", L"", L"[]", L"null"})
			{
				auto error = std::wstring{};
				Assert::IsFalse(ui::theme::ParseThemeFile(text, theme, error));
				Assert::IsTrue(error == L"Invalid JSON. Expected a theme object.");
				Assert::IsTrue(theme.stableId == originalId);
			}
		}

		TEST_METHOD(DiscoveryIsExplicitNonrecursiveAndIsolatesBrokenFilesAndDuplicateIds)
		{
			auto valid = ReadExample();
			WriteFile(m_directory / L"01-valid.JSON", valid, true);
			WriteFile(m_directory / L"02-duplicate.json", valid);

			Replace(valid, L"example.halloween", L"standard");
			WriteFile(m_directory / L"03-reserved.json", valid);
			WriteFile(m_directory / L"04-broken.json", L"{broken");
			WriteFile(m_directory / L"ignored.txt", L"{broken");

			{
				std::ofstream stream(m_directory / L"05-invalid-utf8.json", std::ios::binary);
				stream.write("\xFF\xFE\xFF", 3);
				Assert::IsTrue(stream.good());
			}

			Assert::IsTrue(std::filesystem::create_directory(m_directory / L"nested"));
			WriteFile(m_directory / L"nested" / L"ignored.json", ReadExample());

			const auto baseline = ui::theme::LoadThemeDirectory(L"");
			const auto loaded = ui::theme::LoadThemeDirectory(m_directory.wstring());
			Assert::AreEqual(baseline.themes.size() + 1, loaded.themes.size());
			Assert::AreEqual(size_t{4}, loaded.errors.size());
			Assert::IsTrue(loaded.themes.back().stableId == L"example.halloween");
			Assert::IsTrue(baseline.errors.empty());
			Assert::IsTrue(ui::theme::LoadThemeDirectory(L"relative").errors.size() == 1);
		}

		TEST_METHOD(ReloadReadsEditsRemovalsAndOwnsUnicodeMetadata)
		{
			auto json = ReadExample();
			Replace(json, L"Halloween", L"\u00C9v\u00E9nement");
			const auto path = m_directory / L"theme.json";
			WriteFile(path, json);

			auto first = ui::theme::LoadThemeDirectory(m_directory.wstring());
			Assert::IsTrue(first.errors.empty());
			Assert::IsTrue(first.themes.back().name == L"\u00C9v\u00E9nement");

			auto copy = first.themes.back();
			first = {};
			Assert::IsTrue(copy.name == L"\u00C9v\u00E9nement");

			WriteFile(path, L"{broken");
			const auto broken = ui::theme::LoadThemeDirectory(m_directory.wstring());
			Assert::AreEqual(size_t{1}, broken.errors.size());
			Assert::IsTrue(copy.stableId == L"example.halloween");

			Assert::IsTrue(std::filesystem::remove(path));
			const auto removed = ui::theme::LoadThemeDirectory(m_directory.wstring());
			Assert::IsTrue(removed.errors.empty());
			Assert::AreEqual(broken.themes.size(), removed.themes.size());
		}

		TEST_METHOD(OmittedScheduleProducesAManualOnlyTheme)
		{
			auto json = ReadExample();
			const auto start = json.find(L"  \"schedule\":");
			const auto end = json.find(L"  \"colors\":");
			Assert::IsTrue(start != std::wstring::npos && end > start);
			json.erase(start, end - start);

			auto theme = ui::theme::ThemeDefinition{};
			auto error = std::wstring{};
			Assert::IsTrue(ui::theme::ParseThemeFile(json, theme, error), error.c_str());

			const std::vector<ui::theme::ThemeDefinition> themes{theme};
			Assert::IsNull(
				ui::theme::FindScheduledTheme(themes, std::chrono::year{2024} / std::chrono::October / 31, false));
		}

	private:
		std::filesystem::path m_directory;
	};
} // namespace themefiletest
