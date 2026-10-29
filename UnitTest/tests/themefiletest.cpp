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

		TEST_METHOD(ColorsParseHexChannelsCssAliasesAndSystemMappings)
		{
			auto json = ReadExample();
			Replace(json, L"OldLace", L"#1234ab");
			Replace(json, L"Wheat", L"rEbEcCaPuRpLe");
			Replace(json, L"DarkOrange", L"darkGREY");
			Replace(
				json,
				L"\"value\": \"Black\", \"highContrast\": 8",
				L"\"value\": \"Black\", \"system\": 0, \"highContrast\": 13");

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
					"light": { "Text": { "value": "#123456", "highContrast": 8 } },
					"dark": {}
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
				Assert::AreEqual(expectedDark.value, dark.value);
				Assert::AreEqual(expectedDark.system, dark.system);
				Assert::AreEqual(expectedDark.highContrast, dark.highContrast);

				const auto& light = theme.light.values[index];
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

		TEST_METHOD(InvalidDefinitionsReportErrorsWithoutChangingTheOutput)
		{
			const auto valid = ReadExample();
			const std::pair<std::wstring_view, std::wstring_view> invalid[] = {
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
