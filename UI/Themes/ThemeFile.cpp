#include <StdAfx.h>
#include <UI/Themes/ThemeFile.h>
#include <UI/Themes/BuiltinThemes.h>
#include <UI/Themes/CssNamedColors.h>
#include <winrt/Windows.Data.Json.h>
#include <winrt/Windows.Foundation.Collections.h>
#include <roapi.h>
#include <filesystem>
#include <fstream>
#include <cmath>

#pragma comment(lib, "runtimeobject.lib")

namespace ui::theme
{
	static constexpr std::wstring_view roleNames[] = {
		L"Background",
		L"BackgroundReadOnly",
		L"Glow",
		L"GlowBackground",
		L"GlowText",
		L"FrameSelected",
		L"FrameUnselected",
		L"SelectedBackground",
		L"Arrow",
		L"Text",
		L"TextDisabled",
		L"TextDisabledSelected",
		L"TextReadOnly",
		L"BitmapTransBack",
		L"BitmapTransFore",
		L"Status",
		L"StatusText",
		L"PaneHeaderBackground",
		L"PaneHeaderText",
		L"TextHighlightBackground",
		L"TextHighlight",
		L"MenuGutter",
		L"TestPink",
		L"TestLavender",
		L"TestRed",
		L"TestGreen",
		L"TestOrange"};

	static_assert(std::size(roleNames) == static_cast<size_t>(uiColor::UIEnd));

	static void Require(const bool condition, const std::wstring& message)
	{
		if (!condition) throw winrt::hresult_invalid_argument(message);
	}

	static COLORREF ParseColor(const std::wstring_view text)
	{
		DWORD rgb = 0;
		if (text.size() == 7 && text[0] == L'#')
		{
			for (const auto digit : text.substr(1))
			{
				const auto lower = towlower(digit);
				Require(
					(lower >= L'0' && lower <= L'9') || (lower >= L'a' && lower <= L'f'),
					L"Invalid hex color: " + std::wstring(text));
				rgb = (rgb << 4) | (lower <= L'9' ? lower - L'0' : lower - L'a' + 10);
			}
		}
		else
		{
			const auto found = std::find_if(std::begin(namedColors), std::end(namedColors), [text](const auto& color) {
				return _wcsnicmp(text.data(), color.first.data(), text.size()) == 0 &&
					   text.size() == color.first.size();
			});
			Require(found != std::end(namedColors), L"Unknown color: " + std::wstring(text));
			rgb = found->second;
		}

		return RGB((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);
	}

	static void CheckKeys(
		const winrt::Windows::Data::Json::JsonObject& object,
		const std::initializer_list<std::wstring_view> allowed)
	{
		for (const auto& property : object)
		{
			const auto name = property.Key();
			const std::wstring_view key = name;
			Require(
				std::find(allowed.begin(), allowed.end(), key) != allowed.end(),
				L"Unknown field: " + std::wstring(key));
		}
	}

	static int ReadInteger(
		const winrt::Windows::Data::Json::JsonObject& object,
		const std::wstring_view key,
		const int minimum,
		const int maximum)
	{
		const auto number = object.GetNamedNumber(key);
		Require(
			std::isfinite(number) && number >= minimum && number <= maximum && std::floor(number) == number,
			L"Invalid integer: " + std::wstring(key));

		return static_cast<int>(number);
	}

	static int ReadSystemColor(const winrt::Windows::Data::Json::JsonObject& object, const std::wstring_view key)
	{
		if (!object.HasKey(key)) return -1;

		const auto index = ReadInteger(object, key, -1, COLOR_MENUBAR);
		Require(index != 25, L"Invalid system color index: 25");

		return index;
	}

	static ThemeColors ReadPalette(const winrt::Windows::Data::Json::JsonObject& object, const ThemeColors& defaults)
	{
		Require(object.HasKey(L"TextDisabledSelected"), L"Missing required color role: TextDisabledSelected");

		for (const auto& property : object)
		{
			const auto name = property.Key();
			const std::wstring_view role = name;
			Require(
				std::find(std::begin(roleNames), std::end(roleNames), role) != std::end(roleNames),
				L"Unknown color role: " + std::wstring(role));
		}

		auto colors = defaults;
		for (size_t index = 0; index < std::size(roleNames); ++index)
		{
			const auto role = roleNames[index];
			if (!object.HasKey(role)) continue;

			const auto entry = object.GetNamedObject(role);
			CheckKeys(entry, {L"value", L"system", L"highContrast"});
			colors.values[index] = {
				ParseColor(entry.GetNamedString(L"value")),
				ReadSystemColor(entry, L"system"),
				ReadSystemColor(entry, L"highContrast")};

			Require(
				builtin::Standard.light.values[index].highContrast < 0 || colors.values[index].highContrast >= 0,
				L"Missing High Contrast system color for role: " + std::wstring(role));
		}

		return colors;
	}

	static std::pair<unsigned int, unsigned int> ReadDate(const std::wstring_view text)
	{
		Require(text.size() == 5 && text[2] == L'-', L"Schedule dates must use MM-DD.");
		for (const auto index : {0, 1, 3, 4})
		{
			Require(text[index] >= L'0' && text[index] <= L'9', L"Schedule dates must use MM-DD.");
		}

		const auto month = static_cast<unsigned int>((text[0] - L'0') * 10 + text[1] - L'0');
		const auto day = static_cast<unsigned int>((text[3] - L'0') * 10 + text[4] - L'0');
		Require(
			(std::chrono::year{2000} / std::chrono::month{month} / std::chrono::day{day}).ok(),
			L"Invalid schedule date: " + std::wstring(text));

		return {month, day};
	}

	bool ParseThemeFile(const std::wstring_view json, ThemeDefinition& definition, std::wstring& error)
	{
		const auto initialized = RoInitialize(RO_INIT_SINGLETHREADED);
		auto success = false;
		error.clear();

		try
		{
			winrt::check_hresult(initialized == RPC_E_CHANGED_MODE ? S_OK : initialized);

			auto object = winrt::Windows::Data::Json::JsonObject{nullptr};
			if (!winrt::Windows::Data::Json::JsonObject::TryParse(json, object))
			{
				error = L"Invalid JSON. Expected a theme object.";
			}
			else
			{
				CheckKeys(object, {L"schemaVersion", L"id", L"name", L"schedule", L"colors"});
				Require(object.GetNamedNumber(L"schemaVersion") == 1, L"Unsupported schema version.");

				auto theme = ThemeDefinition{};
				theme.stableId = object.GetNamedString(L"id");
				theme.name = object.GetNamedString(L"name");
				Require(!theme.stableId.empty() && !theme.name.empty(), L"Theme id and name must not be empty.");

				const auto palettes = object.GetNamedObject(L"colors");
				CheckKeys(palettes, {L"light", L"dark"});
				theme.light = ReadPalette(palettes.GetNamedObject(L"light"), builtin::Standard.light);
				theme.dark = ReadPalette(palettes.GetNamedObject(L"dark"), builtin::Standard.dark);

				if (object.HasKey(L"schedule"))
				{
					const auto schedule = object.GetNamedObject(L"schedule");
					CheckKeys(schedule, {L"kind", L"start", L"end", L"priority", L"seasonal"});
					Require(schedule.GetNamedString(L"kind") == L"dateRange", L"Unknown schedule kind.");

					const auto start = ReadDate(schedule.GetNamedString(L"start"));
					const auto end = ReadDate(schedule.GetNamedString(L"end"));
					theme.schedule = {
						ScheduleKind::AnnualDateRange,
						start.first,
						start.second,
						end.first,
						end.second,
						schedule.HasKey(L"priority") ? ReadInteger(schedule, L"priority", INT_MIN, INT_MAX) : 0,
						schedule.HasKey(L"seasonal") ? schedule.GetNamedBoolean(L"seasonal") : false};
				}

				definition = std::move(theme);
				success = true;
			}
		} catch (const winrt::hresult_error& exception)
		{
			error = exception.message();
		}

		if (SUCCEEDED(initialized))
		{
			// Cached factories must be released before WinRT can unload their implementation.
			winrt::clear_factory_cache();
			RoUninitialize();
		}

		return success;
	}

	ThemeCatalog LoadThemeDirectory(const std::wstring& directory)
	{
		auto catalog = ThemeCatalog{};
		catalog.themes.assign(builtin::Definitions.begin(), builtin::Definitions.end());
		const auto firstExternalTheme = catalog.themes.size();
		if (directory.empty()) return catalog;

		const auto path = std::filesystem::path{directory};
		std::error_code error;
		if (!path.is_absolute() || !std::filesystem::is_directory(path, error))
		{
			catalog.errors.push_back({directory, L"Theme directory must be an existing absolute directory."});
			return catalog;
		}

		std::vector<std::filesystem::path> paths;
		auto iterator = std::filesystem::directory_iterator(path, error);
		const auto end = std::filesystem::directory_iterator{};
		while (!error && iterator != end)
		{
			const auto file = iterator->path();
			if (_wcsicmp(file.extension().c_str(), L".json") == 0 && iterator->is_regular_file(error))
			{
				paths.push_back(file);
			}

			if (error) break;

			iterator.increment(error);
		}

		if (error)
		{
			catalog.errors.push_back(
				{directory, winrt::hresult_error{HRESULT_FROM_WIN32(error.value())}.message().c_str()});
		}

		std::sort(paths.begin(), paths.end());
		for (const auto& file : paths)
		{
			auto message = std::wstring{};
			auto theme = ThemeDefinition{};

			try
			{
				std::ifstream stream(file, std::ios::binary);
				if (!stream)
				{
					catalog.errors.push_back({file.wstring(), L"Could not open theme file."});
					continue;
				}

				auto bytes = std::string{std::istreambuf_iterator<char>{stream}, std::istreambuf_iterator<char>{}};
				if (stream.bad())
				{
					catalog.errors.push_back({file.wstring(), L"Could not read theme file."});
					continue;
				}

				if (bytes.starts_with("\xEF\xBB\xBF")) bytes.erase(0, 3);
				if (!ParseThemeFile(winrt::to_hstring(bytes), theme, message))
				{
					catalog.errors.push_back({file.wstring(), message});
					continue;
				}
			} catch (const winrt::hresult_error& exception)
			{
				catalog.errors.push_back({file.wstring(), exception.message().c_str()});
				continue;
			}

			const auto duplicate =
				std::find_if(catalog.themes.begin(), catalog.themes.end(), [&theme](const auto& candidate) {
					return candidate.stableId == theme.stableId;
				});
			if (duplicate != catalog.themes.end())
			{
				catalog.errors.push_back({file.wstring(), L"Duplicate theme id: " + theme.stableId});
				continue;
			}

			catalog.themes.push_back(std::move(theme));
		}

		std::sort(
			catalog.themes.begin() + firstExternalTheme,
			catalog.themes.end(),
			[](const auto& left, const auto& right) {
				const auto nameOrder = _wcsicmp(left.name.c_str(), right.name.c_str());
				return nameOrder == 0 ? left.stableId < right.stableId : nameOrder < 0;
			});

		return catalog;
	}
} // namespace ui::theme
