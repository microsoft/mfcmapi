#include <StdAfx.h>
#include <UI/Themes/ThemeFile.h>
#include <UI/Themes/BuiltinThemes.h>
#include <UI/Themes/CssNamedColors.h>
#include <core/utility/error.h>
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
		L"DisabledSelectedBackground",
		L"DisabledSelectedBorder",
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
		L"TestOrange",
		L"TitleBarBackground",
		L"TitleBarText",
		L"TitleBarButton",
		L"TitleBarButtonHover",
		L"ColumnHeaderBackground",
		L"ColumnHeaderText",
		L"ColumnHeaderArrow",
		L"MenuBackground",
		L"MenuText",
		L"MenuTextDisabled",
		L"StatusTextOutline"};

	static_assert(std::size(roleNames) == static_cast<size_t>(uiColor::UIEnd));

	using JsonObject = winrt::Windows::Data::Json::JsonObject;
	using JsonArray = winrt::Windows::Data::Json::JsonArray;

	template <typename Call> static bool ReadJsonApi(Call&& call, const std::wstring_view context, std::wstring& error)
	{
		try
		{
			call();
		} catch (const winrt::hresult_error& exception)
		{
			error = std::wstring(context) + L": " + exception.message().c_str();
			return false;
		}

		return true;
	}

	static bool
	ReadObject(const JsonObject& object, const std::wstring_view key, JsonObject& value, std::wstring& error)
	{
		return ReadJsonApi([&] { value = object.GetNamedObject(key); }, key, error);
	}

	static bool
	ReadString(const JsonObject& object, const std::wstring_view key, std::wstring& value, std::wstring& error)
	{
		return ReadJsonApi([&] { value = object.GetNamedString(key); }, key, error);
	}

	static bool HasKey(const JsonObject& object, const std::wstring_view key, bool& present, std::wstring& error)
	{
		return ReadJsonApi([&] { present = object.HasKey(key); }, key, error);
	}

	static bool ReadKeys(const JsonObject& object, std::vector<std::wstring>& keys, std::wstring& error)
	{
		return ReadJsonApi(
			[&] {
				for (const auto& property : object)
				{
					keys.emplace_back(property.Key());
				}
			},
			L"Object fields",
			error);
	}

	static bool ParseColor(const std::wstring_view text, COLORREF& value, std::wstring& error)
	{
		DWORD rgb = 0;
		if (text.size() == 7 && text[0] == L'#')
		{
			for (const auto digit : text.substr(1))
			{
				const auto lower = towlower(digit);
				if (!((lower >= L'0' && lower <= L'9') || (lower >= L'a' && lower <= L'f')))
				{
					error = L"Invalid hex color: " + std::wstring(text);
					return false;
				}

				rgb = (rgb << 4) | (lower <= L'9' ? lower - L'0' : lower - L'a' + 10);
			}
		}
		else
		{
			const auto found = std::find_if(std::begin(namedColors), std::end(namedColors), [text](const auto& color) {
				return _wcsnicmp(text.data(), color.first.data(), text.size()) == 0 &&
					   text.size() == color.first.size();
			});
			if (found == std::end(namedColors))
			{
				error = L"Unknown color: " + std::wstring(text);
				return false;
			}

			rgb = found->second;
		}

		value = RGB((rgb >> 16) & 0xFF, (rgb >> 8) & 0xFF, rgb & 0xFF);
		return true;
	}

	static bool
	CheckKeys(const JsonObject& object, const std::initializer_list<std::wstring_view> allowed, std::wstring& error)
	{
		auto keys = std::vector<std::wstring>{};
		if (!ReadKeys(object, keys, error)) return false;

		for (const auto& key : keys)
		{
			if (std::find(allowed.begin(), allowed.end(), key) == allowed.end())
			{
				error = L"Unknown field: " + key;
				return false;
			}
		}

		return true;
	}

	static bool ReadInteger(
		const JsonObject& object,
		const std::wstring_view key,
		const int minimum,
		const int maximum,
		int& value,
		std::wstring& error)
	{
		double number{};
		if (!ReadJsonApi([&] { number = object.GetNamedNumber(key); }, key, error)) return false;

		if (!std::isfinite(number) || number < minimum || number > maximum || std::floor(number) != number)
		{
			error = L"Invalid integer: " + std::wstring(key);
			return false;
		}

		value = static_cast<int>(number);
		return true;
	}

	static bool ReadSystemColor(const JsonObject& object, const std::wstring_view key, int& value, std::wstring& error)
	{
		bool present{};
		if (!HasKey(object, key, present, error)) return false;

		value = -1;
		if (!present) return true;

		if (!ReadInteger(object, key, -1, COLOR_MENUBAR, value, error)) return false;

		if (value == 25)
		{
			error = L"Invalid system color index: 25";
			return false;
		}

		return true;
	}

	static bool
	ReadPalette(const JsonObject& object, const ThemeColors& defaults, ThemeColors& colors, std::wstring& error)
	{
		auto keys = std::vector<std::wstring>{};
		if (!ReadKeys(object, keys, error)) return false;

		for (const auto& role : keys)
		{
			if (std::find(std::begin(roleNames), std::end(roleNames), role) == std::end(roleNames))
			{
				error = L"Unknown color role: " + role;
				return false;
			}
		}

		colors = defaults;
		for (size_t index = 0; index < std::size(roleNames); ++index)
		{
			const auto role = roleNames[index];
			if (std::find(keys.begin(), keys.end(), role) == keys.end()) continue;

			auto entry = JsonObject{nullptr};
			auto text = std::wstring{};
			auto& color = colors.values[index];
			if (!ReadObject(object, role, entry, error) ||
				!CheckKeys(entry, {L"value", L"system", L"highContrast"}, error) ||
				!ReadString(entry, L"value", text, error) || !ParseColor(text, color.value, error) ||
				!ReadSystemColor(entry, L"system", color.system, error) ||
				!ReadSystemColor(entry, L"highContrast", color.highContrast, error))
			{
				error = std::wstring(role) + L": " + error;
				return false;
			}

			if (builtin::Standard.light.values[index].highContrast >= 0 && color.highContrast < 0)
			{
				error = L"Missing High Contrast system color for role: " + std::wstring(role);
				return false;
			}
		}

		return true;
	}

	static bool
	ReadDate(const std::wstring_view text, std::pair<unsigned int, unsigned int>& value, std::wstring& error)
	{
		if (text.size() != 5 || text[2] != L'-')
		{
			error = L"Schedule dates must use MM-DD.";
			return false;
		}

		for (const auto index : {0, 1, 3, 4})
		{
			if (text[index] < L'0' || text[index] > L'9')
			{
				error = L"Schedule dates must use MM-DD.";
				return false;
			}
		}

		const auto month = static_cast<unsigned int>((text[0] - L'0') * 10 + text[1] - L'0');
		const auto day = static_cast<unsigned int>((text[3] - L'0') * 10 + text[4] - L'0');
		if (!(std::chrono::year{2000} / std::chrono::month{month} / std::chrono::day{day}).ok())
		{
			error = L"Invalid schedule date: " + std::wstring(text);
			return false;
		}

		value = {month, day};
		return true;
	}

	static bool ReadPaletteObject(
		const JsonObject& palettes,
		const std::wstring_view key,
		const ThemeColors& defaults,
		ThemeColors& colors,
		std::wstring& error)
	{
		auto object = JsonObject{nullptr};
		if (!ReadObject(palettes, key, object, error) || !ReadPalette(object, defaults, colors, error))
		{
			error = L"colors." + std::wstring(key) + L": " + error;
			return false;
		}

		return true;
	}

	static bool ReadColorArray(
		const JsonObject& object,
		const std::wstring_view key,
		std::vector<COLORREF>& colors,
		std::wstring& error)
	{
		auto array = JsonArray{nullptr};
		if (!ReadJsonApi([&] { array = object.GetNamedArray(key); }, key, error)) return false;

		uint32_t size{};
		if (!ReadJsonApi([&] { size = array.Size(); }, key, error)) return false;

		if (size > 16)
		{
			error = std::wstring(key) + L" supports at most 16 colors.";
			return false;
		}

		colors.clear();
		for (uint32_t index = 0; index < size; ++index)
		{
			const auto context = std::wstring(key) + L"[" + std::to_wstring(index) + L"]";
			auto text = std::wstring{};
			if (!ReadJsonApi([&] { text = array.GetStringAt(index); }, context, error)) return false;

			COLORREF color{};
			if (!ParseColor(text, color, error))
			{
				error = context + L": " + error;
				return false;
			}

			colors.push_back(color);
		}

		return true;
	}

	static bool ReadStatusBar(const JsonObject& object, ThemeDefinition& theme, std::wstring& error)
	{
		auto statusBar = JsonObject{nullptr};
		auto style = std::wstring{};
		if (!ReadObject(object, L"statusBar", statusBar, error) ||
			!CheckKeys(statusBar, {L"style", L"colors"}, error) || !ReadString(statusBar, L"style", style, error))
		{
			return false;
		}

		if (style != L"ribbon" && style != L"bands" && style != L"gradient")
		{
			error = L"Status bar style must be ribbon, bands, or gradient.";
			return false;
		}

		theme.statusBarStyle = style == L"gradient" ? StatusBarStyle::Gradient
							   : style == L"bands"	? StatusBarStyle::Bands
													: StatusBarStyle::Ribbon;
		return ReadColorArray(statusBar, L"colors", theme.statusRibbon, error);
	}

	static bool ReadSchedule(const JsonObject& object, ThemeDefinition& theme, std::wstring& error)
	{
		auto schedule = JsonObject{nullptr};
		auto kind = std::wstring{};
		if (!ReadObject(object, L"schedule", schedule, error) ||
			!CheckKeys(schedule, {L"kind", L"start", L"end", L"priority", L"seasonal"}, error) ||
			!ReadString(schedule, L"kind", kind, error))
		{
			return false;
		}

		if (kind != L"dateRange")
		{
			error = L"Unknown schedule kind.";
			return false;
		}

		auto start = std::pair<unsigned int, unsigned int>{};
		auto end = std::pair<unsigned int, unsigned int>{};
		for (const auto& field : {std::pair{L"start", &start}, std::pair{L"end", &end}})
		{
			auto text = std::wstring{};
			if (!ReadString(schedule, field.first, text, error) || !ReadDate(text, *field.second, error))
			{
				error = std::wstring(field.first) + L": " + error;
				return false;
			}
		}

		int priority{};
		bool seasonal{};
		bool present{};
		if (!HasKey(schedule, L"priority", present, error)) return false;

		if (present && !ReadInteger(schedule, L"priority", INT_MIN, INT_MAX, priority, error)) return false;

		if (!HasKey(schedule, L"seasonal", present, error)) return false;

		if (present && !ReadJsonApi([&] { seasonal = schedule.GetNamedBoolean(L"seasonal"); }, L"seasonal", error))
		{
			return false;
		}

		theme.schedule = {
			ScheduleKind::AnnualDateRange, start.first, start.second, end.first, end.second, priority, seasonal};
		return true;
	}

	static bool ReadDefinition(const JsonObject& object, ThemeDefinition& theme, std::wstring& error)
	{
		if (!CheckKeys(
				object,
				{L"schemaVersion", L"id", L"name", L"schedule", L"colors", L"statusRibbon", L"statusBar"},
				error))
		{
			return false;
		}

		double version{};
		if (!ReadJsonApi([&] { version = object.GetNamedNumber(L"schemaVersion"); }, L"schemaVersion", error))
			return false;

		if (version != 1)
		{
			error = L"Unsupported schema version.";
			return false;
		}

		if (!ReadString(object, L"id", theme.stableId, error) || !ReadString(object, L"name", theme.name, error))
			return false;

		if (theme.stableId.empty() || theme.name.empty())
		{
			error = L"Theme id and name must not be empty.";
			return false;
		}

		auto palettes = JsonObject{nullptr};
		if (!ReadObject(object, L"colors", palettes, error) || !CheckKeys(palettes, {L"light", L"dark"}, error) ||
			!ReadPaletteObject(palettes, L"light", builtin::Standard.light, theme.light, error) ||
			!ReadPaletteObject(palettes, L"dark", builtin::Standard.dark, theme.dark, error))
		{
			return false;
		}

		theme.statusRibbon = builtin::Standard.statusRibbon;
		theme.statusBarStyle = builtin::Standard.statusBarStyle;
		bool hasStatusBar{};
		bool hasRibbon{};
		bool hasSchedule{};
		if (!HasKey(object, L"statusBar", hasStatusBar, error) || !HasKey(object, L"statusRibbon", hasRibbon, error) ||
			!HasKey(object, L"schedule", hasSchedule, error))
		{
			return false;
		}

		if (hasStatusBar && hasRibbon)
		{
			error = L"Use either statusBar or statusRibbon, not both.";
			return false;
		}

		if (hasStatusBar && !ReadStatusBar(object, theme, error))
		{
			error = L"statusBar: " + error;
			return false;
		}

		if (hasRibbon && !ReadColorArray(object, L"statusRibbon", theme.statusRibbon, error)) return false;

		if (hasSchedule && !ReadSchedule(object, theme, error))
		{
			error = L"schedule: " + error;
			return false;
		}

		return true;
	}

	bool ParseThemeFile(const std::wstring_view json, ThemeDefinition& definition, std::wstring& error)
	{
		error.clear();
		const auto initialized = RoInitialize(RO_INIT_SINGLETHREADED);
		if (FAILED(initialized) && initialized != RPC_E_CHANGED_MODE)
		{
			WC_H_S(initialized);
			error = L"Could not initialize the theme JSON runtime: " +
					std::wstring(winrt::hresult_error{initialized}.message());
			return false;
		}

		struct RuntimeCleanup
		{
			bool uninitialize;
			~RuntimeCleanup()
			{
				if (uninitialize)
				{
					// Cached factories must be released before WinRT can unload their implementation.
					winrt::clear_factory_cache();
					RoUninitialize();
				}
			}
		} cleanup{SUCCEEDED(initialized)};

		auto object = JsonObject{nullptr};
		bool parsed{};
		if (!ReadJsonApi([&] { parsed = JsonObject::TryParse(json, object); }, L"JSON", error)) return false;

		if (!parsed)
		{
			error = L"Invalid JSON. Expected a theme object.";
			return false;
		}

		auto theme = ThemeDefinition{};
		if (!ReadDefinition(object, theme, error)) return false;

		definition = std::move(theme);
		return true;
	}

	static bool ReadUtf8(const std::string& bytes, std::wstring& text, std::wstring& error)
	{
		text.clear();
		if (bytes.empty()) return true;

		if (bytes.size() > INT_MAX)
		{
			error = L"Theme file is too large to decode as UTF-8.";
			return false;
		}

		const auto size = static_cast<int>(bytes.size());
		const auto length =
			WC_D(int, MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, bytes.data(), size, nullptr, 0));
		if (!length)
		{
			error = L"Could not decode theme file as UTF-8.";
			return false;
		}

		text.resize(length);
		if (!WC_D(int, MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS, bytes.data(), size, text.data(), length)))
		{
			error = L"Could not decode theme file as UTF-8.";
			return false;
		}

		return true;
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

			auto json = std::wstring{};
			if (!ReadUtf8(bytes, json, message) || !ParseThemeFile(json, theme, message))
			{
				catalog.errors.push_back({file.wstring(), message});
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
			catalog.themes.begin() + firstExternalTheme, catalog.themes.end(), [](const auto& left, const auto& right) {
				const auto nameOrder = _wcsicmp(left.name.c_str(), right.name.c_str());
				return nameOrder == 0 ? left.stableId < right.stableId : nameOrder < 0;
			});

		return catalog;
	}
} // namespace ui::theme
