#include <StdAfx.h>
#include <UI/Themes/Theme.h>
#include <UI/Themes/ThemeFile.h>
#include <UI/UIFunctions.h>
#include <UI/Dialogs/BaseDialog.h>
#include <core/utility/registry.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <algorithm>
#include <optional>

namespace ui::theme
{
	static UINT_PTR scheduleTimer = 0;

	static void CALLBACK ThemeScheduleTimerProc(HWND, UINT, UINT_PTR, DWORD) { RefreshScheduledTheme(); }

	const ThemeDefinition& ResolveActiveTheme()
	{
		if (!registry::themesEnabled || !registry::automaticThemes)
		{
			return ResolveTheme(registry::activeTheme.szCurSTRING, registry::themesEnabled);
		}

		auto localTime = SYSTEMTIME{};
		GetLocalTime(&localTime);

		const auto date = std::chrono::year{localTime.wYear} / std::chrono::month{localTime.wMonth} /
						  std::chrono::day{localTime.wDay};

		return ResolveTheme(registry::activeTheme.szCurSTRING, true, true, date, UsesSouthernSeasons());
	}

	void StopThemeSchedule() noexcept
	{
		if (!scheduleTimer) return;

		WC_B_S(KillTimer(nullptr, scheduleTimer));
		scheduleTimer = 0;
	}

	void StartThemeSchedule() noexcept
	{
		StopThemeSchedule();

		auto localTime = SYSTEMTIME{};
		GetLocalTime(&localTime);

		const auto today = std::chrono::year{localTime.wYear} / std::chrono::month{localTime.wMonth} /
						   std::chrono::day{localTime.wDay};
		const auto tomorrow = std::chrono::year_month_day{std::chrono::sys_days{today} + std::chrono::days{1}};

		auto midnight = SYSTEMTIME{};
		midnight.wYear = static_cast<WORD>(static_cast<int>(tomorrow.year()));
		midnight.wMonth = static_cast<WORD>(static_cast<unsigned int>(tomorrow.month()));
		midnight.wDay = static_cast<WORD>(static_cast<unsigned int>(tomorrow.day()));

		auto utcMidnight = SYSTEMTIME{};
		if (FAILED(WC_B(TzSpecificLocalTimeToSystemTime(nullptr, &midnight, &utcMidnight)))) return;

		auto deadline = FILETIME{};
		if (FAILED(WC_B(SystemTimeToFileTime(&utcMidnight, &deadline)))) return;

		auto now = FILETIME{};
		GetSystemTimeAsFileTime(&now);

		const auto getTicks = [](const FILETIME& time) {
			return (static_cast<ULONGLONG>(time.dwHighDateTime) << 32) | time.dwLowDateTime;
		};
		const auto delay = getTicks(deadline) > getTicks(now) ? (getTicks(deadline) - getTicks(now)) / 10000 + 1 : 1;

		scheduleTimer = WC_D(UINT_PTR, SetTimer(nullptr, 0, static_cast<UINT>(delay), ThemeScheduleTimerProc));
	}

	void RefreshScheduledTheme() noexcept
	{
		if (&ui::GetActiveTheme() != &ResolveActiveTheme())
		{
			ApplyThemeToAllWindows();
		}
		else
		{
			EnumThreadWindows(
				GetCurrentThreadId(),
				[](HWND hWnd, LPARAM) -> BOOL {
					::SendMessage(hWnd, WM_MFCMAPI_THEMECHANGED, 0, 0);
					return true;
				},
				0);
		}

		StartThemeSchedule();
	}

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

		if (registry::activeTheme.szCurSTRING == stableId && !registry::automaticThemes) return;

		registry::automaticThemes = false;
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

	void CycleTheme(const bool reverse)
	{
		if (!registry::themesEnabled) return;

		const auto& next = GetNextTheme(ui::GetActiveTheme().stableId, reverse);
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

	void SetAutomaticThemes(const bool enabled)
	{
		if (static_cast<bool>(registry::automaticThemes) == enabled) return;

		registry::automaticThemes = enabled;
		registry::WriteToRegistry();
		ApplyThemeToAllWindows();
	}

	static ThemeCatalog& GetCatalog()
	{
		static auto catalog = LoadThemeDirectory(registry::themeDirectory.szCurSTRING);

		return catalog;
	}

	const std::vector<ThemeDefinition>& GetThemes() { return GetCatalog().themes; }

	const std::vector<ThemeFileError>& GetThemeFileErrors() { return GetCatalog().errors; }

	void ReloadThemes(const std::wstring& directory)
	{
		auto replacement = LoadThemeDirectory(directory);
		auto previous = std::move(GetCatalog());
		GetCatalog() = std::move(replacement);

		registry::themeDirectory = directory;
		registry::WriteToRegistry();

		// Keep the old applied definition alive until every window has refreshed.
		ApplyThemeToAllWindows();
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

	const ThemeDefinition& ResolveTheme(
		const std::wstring_view selectedTheme,
		const bool enabled,
		const bool automatic,
		const std::chrono::year_month_day date,
		const bool southernHemisphere)
	{
		if (enabled && automatic)
		{
			if (const auto scheduled = FindScheduledTheme(GetThemes(), date, southernHemisphere)) return *scheduled;
		}

		return ResolveTheme(selectedTheme, enabled);
	}

	bool UsesSouthernSeasons()
	{
		const auto region = GetUserGeoID(GEOCLASS_NATION);
		if (region != GEOID_NOT_AVAILABLE)
		{
			wchar_t latitude[32]{};

			if (WC_D(int, GetGeoInfoW(region, GEO_LATITUDE, latitude, _countof(latitude), 0)))
			{
				return latitude[0] == L'-';
			}
		}

		output::DebugPrint(
			output::dbgLevel::Generic,
			L"Theme scheduling: Windows region latitude unavailable; using Northern Hemisphere seasons.\n");

		return false;
	}

	static std::chrono::year_month_day
	ScheduleBoundary(const std::chrono::year year, const unsigned int month, const unsigned int day, const bool shift)
	{
		using namespace std::chrono;

		const auto sourceMonth = std::chrono::month{month};
		const auto targetMonth = shift ? sourceMonth + months{6} : sourceMonth;
		const auto lastDay = year_month_day{year / targetMonth / last}.day();

		// Shift full-month seasonal boundaries as full months, including February.
		const auto sourceLastDay = year_month_day{std::chrono::year{2000} / sourceMonth / last}.day();
		const auto targetDay =
			shift && day == static_cast<unsigned int>(sourceLastDay) ? lastDay : min(std::chrono::day{day}, lastDay);

		return year / targetMonth / targetDay;
	}

	static std::optional<std::chrono::days>
	MatchingDuration(const Schedule& schedule, const std::chrono::year_month_day date, const bool southernHemisphere)
	{
		using namespace std::chrono;

		if (schedule.kind == ScheduleKind::None) return std::nullopt;

		const auto startDate = year{2000} / month{schedule.startMonth} / day{schedule.startDay};
		const auto endDate = year{2000} / month{schedule.endMonth} / day{schedule.endDay};

		if (schedule.kind != ScheduleKind::AnnualDateRange || !startDate.ok() || !endDate.ok())
		{
			WC_H_S(E_INVALIDARG);
			return std::nullopt;
		}

		const auto shift = southernHemisphere && schedule.seasonal;
		auto start = ScheduleBoundary(date.year(), schedule.startMonth, schedule.startDay, shift);
		auto end = ScheduleBoundary(date.year(), schedule.endMonth, schedule.endDay, shift);

		if (sys_days{end} < sys_days{start})
		{
			if (sys_days{date} < sys_days{start})
			{
				start = ScheduleBoundary(date.year() - years{1}, schedule.startMonth, schedule.startDay, shift);
			}
			else
			{
				end = ScheduleBoundary(date.year() + years{1}, schedule.endMonth, schedule.endDay, shift);
			}
		}

		if (sys_days{date} < sys_days{start} || sys_days{date} > sys_days{end}) return std::nullopt;

		return sys_days{end} - sys_days{start} + days{1};
	}

	const ThemeDefinition* FindScheduledTheme(
		const std::vector<ThemeDefinition>& themes,
		const std::chrono::year_month_day date,
		const bool southernHemisphere)
	{
		if (!date.ok())
		{
			WC_H_S(E_INVALIDARG);
			return nullptr;
		}

		const ThemeDefinition* selected = nullptr;
		auto selectedDuration = std::chrono::days{};

		for (const auto& theme : themes)
		{
			const auto duration = MatchingDuration(theme.schedule, date, southernHemisphere);
			if (!duration) continue;

			if (!selected || theme.schedule.priority > selected->schedule.priority ||
				(theme.schedule.priority == selected->schedule.priority &&
				 (*duration < selectedDuration ||
				  (*duration == selectedDuration && theme.stableId < selected->stableId))))
			{
				selected = &theme;
				selectedDuration = *duration;
			}
		}

		return selected;
	}

	COLORREF GetColor(const ThemeDefinition& definition, const uiTheme appearance, const uiColor role) noexcept
	{
		const auto& colors = appearance == uiTheme::Dark ? definition.dark : definition.light;
		const auto& color = colors.values.at(static_cast<size_t>(role));
		if (appearance == uiTheme::HighContrast && color.highContrast >= 0) return GetSysColor(color.highContrast);

		if (color.system >= 0) return GetSysColor(color.system);

		return color.value;
	}

	std::span<const COLORREF> GetStatusRibbon(const ThemeDefinition& definition, const uiTheme appearance) noexcept
	{
		if (appearance == uiTheme::HighContrast) return {};
		return definition.statusRibbon;
	}

	std::vector<StatusRibbonBand>
	LayoutStatusRibbon(const std::span<const COLORREF> colors, const RECT& rect, const int height, const bool gradient)
	{
		std::vector<StatusRibbonBand> bands;
		if (colors.empty() || height <= 0 || rect.right <= rect.left || rect.bottom <= rect.top) return bands;

		const auto width = static_cast<long long>(rect.right) - rect.left;
		const auto segments = gradient && colors.size() > 1 ? colors.size() - 1 : colors.size();
		const auto count = static_cast<long long>(segments);
		const auto bottom = rect.top + min(height, rect.bottom - rect.top);
		for (size_t index = 0; index < segments; ++index)
		{
			const auto left = rect.left + static_cast<LONG>(width * static_cast<long long>(index) / count);
			const auto right = rect.left + static_cast<LONG>(width * static_cast<long long>(index + 1) / count);
			const auto endColor = gradient && colors.size() > 1 ? colors[index + 1] : colors[index];
			if (right > left) bands.push_back({{left, rect.top, right, bottom}, colors[index], endColor});
		}

		return bands;
	}

	const ThemeDefinition& GetNextTheme(const std::wstring_view stableId, const bool reverse)
	{
		const auto& themes = GetThemes();
		const auto current = FindTheme(stableId);
		const auto index = current ? static_cast<size_t>(current - themes.data()) : 0;

		return themes[(index + (reverse ? themes.size() - 1 : 1)) % themes.size()];
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
