#pragma once
#include <UI/Themes/Theme.h>
#include <UI/Themes/Colors.h>

namespace ui::theme::builtin
{
	// Complete light/dark palettes follow uiColor order.
	inline constexpr ThemeDefinition Standard = {
		L"standard",
		L"Standard",
		L"Standard",
		L"",
		{ScheduleKind::None, 0, 0, 0, 0},
		{{{
			{color::White, COLOR_WINDOW, COLOR_WINDOW}, // Background
			{color::LightGray, -1, COLOR_BTNFACE}, // BackgroundReadOnly
			{color::custom::HeroBlue, -1, COLOR_HIGHLIGHT}, // Glow
			{color::custom::HeroPaleBlue, -1, COLOR_HIGHLIGHT}, // GlowBackground
			{color::Black, -1, COLOR_HIGHLIGHTTEXT}, // GlowText
			{color::Black, COLOR_WINDOWTEXT, COLOR_WINDOWTEXT}, // FrameSelected
			{color::Gray, COLOR_3DLIGHT, COLOR_GRAYTEXT}, // FrameUnselected
			{color::custom::HeroSelectionBlue, -1, COLOR_HIGHLIGHT}, // SelectedBackground
			{color::Gray, COLOR_GRAYTEXT, COLOR_WINDOWTEXT}, // Arrow
			{color::Black, COLOR_WINDOWTEXT, COLOR_WINDOWTEXT}, // Text
			{color::Gray, COLOR_GRAYTEXT, COLOR_GRAYTEXT}, // TextDisabled
			{color::Black, -1, COLOR_WINDOWTEXT}, // TextReadOnly
			{color::Magenta, -1, -1}, // BitmapTransBack
			{color::Cyan, -1, -1}, // BitmapTransFore
			{color::custom::HeroBlue, -1, COLOR_HIGHLIGHT}, // Status
			{color::White, -1, COLOR_HIGHLIGHTTEXT}, // StatusText
			{color::custom::HeroPaleBlue, -1, COLOR_BTNFACE}, // PaneHeaderBackground
			{color::Black, -1, COLOR_BTNTEXT}, // PaneHeaderText
			{color::custom::HeroSelectionBlue, -1, COLOR_HIGHLIGHT}, // TextHighlightBackground
			{color::Black, -1, COLOR_HIGHLIGHTTEXT}, // TextHighlight
			{color::LightGray, -1, COLOR_BTNFACE}, // MenuGutter
			{color::Pink, -1, -1}, // TestPink
			{color::Lavender, -1, -1}, // TestLavender
			{color::IndianRed, -1, -1}, // TestRed
			{color::Lime, -1, -1}, // TestGreen
			{color::Coral, -1, -1}, // TestOrange
		}}},
		{{{
			{RGB(0x20, 0x20, 0x20), -1, COLOR_WINDOW}, // Background
			{color::custom::CharcoalGray, -1, COLOR_BTNFACE}, // BackgroundReadOnly
			{color::custom::HeroBlue, -1, COLOR_HIGHLIGHT}, // Glow
			{color::custom::HeroDarkBlue, -1, COLOR_HIGHLIGHT}, // GlowBackground
			{color::White, -1, COLOR_HIGHLIGHTTEXT}, // GlowText
			{color::LightGray, -1, COLOR_WINDOWTEXT}, // FrameSelected
			{RGB(0x5A, 0x5A, 0x5A), -1, COLOR_GRAYTEXT}, // FrameUnselected
			{color::custom::HeroSelectionDarkBlue, -1, COLOR_HIGHLIGHT}, // SelectedBackground
			{RGB(0xC8, 0xC8, 0xC8), -1, COLOR_WINDOWTEXT}, // Arrow
			{RGB(0xF1, 0xF1, 0xF1), -1, COLOR_WINDOWTEXT}, // Text
			{RGB(0x8A, 0x8A, 0x8A), -1, COLOR_GRAYTEXT}, // TextDisabled
			{RGB(0xD6, 0xD6, 0xD6), -1, COLOR_WINDOWTEXT}, // TextReadOnly
			{color::Magenta, -1, -1}, // BitmapTransBack
			{color::Cyan, -1, -1}, // BitmapTransFore
			{color::custom::HeroBlue, -1, COLOR_HIGHLIGHT}, // Status
			{color::White, -1, COLOR_HIGHLIGHTTEXT}, // StatusText
			{color::custom::HeroDarkBlue, -1, COLOR_BTNFACE}, // PaneHeaderBackground
			{color::White, -1, COLOR_BTNTEXT}, // PaneHeaderText
			{color::custom::HeroSelectionDarkBlue, -1, COLOR_HIGHLIGHT}, // TextHighlightBackground
			{color::White, -1, COLOR_HIGHLIGHTTEXT}, // TextHighlight
			{color::custom::CharcoalGray, -1, COLOR_BTNFACE}, // MenuGutter
			{color::Pink, -1, -1}, // TestPink
			{color::Lavender, -1, -1}, // TestLavender
			{color::IndianRed, -1, -1}, // TestRed
			{color::Lime, -1, -1}, // TestGreen
			{color::Coral, -1, -1}, // TestOrange
		}}}};

	inline constexpr std::array<ThemeDefinition, 5> Definitions = {{
		Standard,
		[] {
			auto theme = Standard;
			theme.stableId = L"season.spring";
			theme.name = L"Spring";
			theme.category = L"Seasonal";
			theme.schedule = {ScheduleKind::AnnualDateRange, 3, 1, 5, 31, 0, true};

			theme.light.values[static_cast<size_t>(uiColor::Glow)] = {color::DarkGreen, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::GlowBackground)] = {color::Honeydew, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::GlowText)] = {color::DarkGreen, -1, COLOR_HIGHLIGHTTEXT};
			theme.light.values[static_cast<size_t>(uiColor::FrameSelected)] = {color::SteelBlue, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::SelectedBackground)] = {
				color::PaleGreen, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::Status)] = {color::DarkGreen, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::PaneHeaderBackground)] = {
				color::Honeydew, -1, COLOR_BTNFACE};
			theme.light.values[static_cast<size_t>(uiColor::PaneHeaderText)] = {color::DarkGreen, -1, COLOR_BTNTEXT};
			theme.light.values[static_cast<size_t>(uiColor::TextHighlightBackground)] = {
				color::PaleGreen, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::TextHighlight)] = {
				color::DarkGreen, -1, COLOR_HIGHLIGHTTEXT};
			theme.light.values[static_cast<size_t>(uiColor::MenuGutter)] = {color::Honeydew, -1, COLOR_BTNFACE};

			theme.dark.values[static_cast<size_t>(uiColor::Glow)] = {color::LightGreen, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::GlowBackground)] = {color::DarkGreen, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::FrameSelected)] = {
				color::LightSkyBlue, -1, COLOR_WINDOWTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::SelectedBackground)] = {
				color::DarkOliveGreen, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::Status)] = {color::DarkGreen, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::PaneHeaderBackground)] = {
				color::DarkGreen, -1, COLOR_BTNFACE};
			theme.dark.values[static_cast<size_t>(uiColor::TextHighlightBackground)] = {
				color::DarkOliveGreen, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::MenuGutter)] = {color::DarkGreen, -1, COLOR_BTNFACE};

			return theme;
		}(),
		[] {
			auto theme = Standard;
			theme.stableId = L"season.summer";
			theme.name = L"Summer";
			theme.category = L"Seasonal";
			theme.schedule = {ScheduleKind::AnnualDateRange, 6, 1, 8, 31, 0, true};

			theme.light.values[static_cast<size_t>(uiColor::Glow)] = {color::Teal, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::GlowBackground)] = {color::Azure, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::GlowText)] = {color::MidnightBlue, -1, COLOR_HIGHLIGHTTEXT};
			theme.light.values[static_cast<size_t>(uiColor::FrameSelected)] = {color::Chocolate, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::SelectedBackground)] = {
				color::PowderBlue, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::Status)] = {color::Teal, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::PaneHeaderBackground)] = {color::Azure, -1, COLOR_BTNFACE};
			theme.light.values[static_cast<size_t>(uiColor::PaneHeaderText)] = {color::MidnightBlue, -1, COLOR_BTNTEXT};
			theme.light.values[static_cast<size_t>(uiColor::TextHighlightBackground)] = {
				color::PowderBlue, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::TextHighlight)] = {
				color::MidnightBlue, -1, COLOR_HIGHLIGHTTEXT};
			theme.light.values[static_cast<size_t>(uiColor::MenuGutter)] = {color::Azure, -1, COLOR_BTNFACE};

			theme.dark.values[static_cast<size_t>(uiColor::Glow)] = {color::Turquoise, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::GlowBackground)] = {
				color::DarkSlateGray, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::FrameSelected)] = {color::PeachPuff, -1, COLOR_WINDOWTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::SelectedBackground)] = {color::Teal, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::Status)] = {color::DarkSlateGray, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::PaneHeaderBackground)] = {
				color::DarkSlateGray, -1, COLOR_BTNFACE};
			theme.dark.values[static_cast<size_t>(uiColor::TextHighlightBackground)] = {
				color::Teal, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::MenuGutter)] = {color::DarkSlateGray, -1, COLOR_BTNFACE};

			return theme;
		}(),
		[] {
			auto theme = Standard;
			theme.stableId = L"season.autumn";
			theme.name = L"Autumn";
			theme.category = L"Seasonal";
			theme.schedule = {ScheduleKind::AnnualDateRange, 9, 1, 11, 30, 0, true};

			theme.light.values[static_cast<size_t>(uiColor::Glow)] = {color::SaddleBrown, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::GlowBackground)] = {color::OldLace, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::GlowText)] = {color::Maroon, -1, COLOR_HIGHLIGHTTEXT};
			theme.light.values[static_cast<size_t>(uiColor::FrameSelected)] = {color::SteelBlue, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::SelectedBackground)] = {color::Wheat, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::Status)] = {color::SaddleBrown, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::PaneHeaderBackground)] = {
				color::OldLace, -1, COLOR_BTNFACE};
			theme.light.values[static_cast<size_t>(uiColor::PaneHeaderText)] = {color::Maroon, -1, COLOR_BTNTEXT};
			theme.light.values[static_cast<size_t>(uiColor::TextHighlightBackground)] = {
				color::Wheat, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::TextHighlight)] = {color::Maroon, -1, COLOR_HIGHLIGHTTEXT};
			theme.light.values[static_cast<size_t>(uiColor::MenuGutter)] = {color::OldLace, -1, COLOR_BTNFACE};

			theme.dark.values[static_cast<size_t>(uiColor::Glow)] = {color::SandyBrown, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::GlowBackground)] = {color::Maroon, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::FrameSelected)] = {
				color::LightSteelBlue, -1, COLOR_WINDOWTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::SelectedBackground)] = {
				color::SaddleBrown, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::Status)] = {color::Maroon, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::PaneHeaderBackground)] = {color::Maroon, -1, COLOR_BTNFACE};
			theme.dark.values[static_cast<size_t>(uiColor::TextHighlightBackground)] = {
				color::SaddleBrown, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::MenuGutter)] = {color::Maroon, -1, COLOR_BTNFACE};

			return theme;
		}(),
		[] {
			auto theme = Standard;
			theme.stableId = L"season.winter";
			theme.name = L"Winter";
			theme.category = L"Seasonal";
			theme.schedule = {ScheduleKind::AnnualDateRange, 12, 1, 2, 29, 0, true};

			theme.light.values[static_cast<size_t>(uiColor::Glow)] = {color::RoyalBlue, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::GlowBackground)] = {color::AliceBlue, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::GlowText)] = {color::MidnightBlue, -1, COLOR_HIGHLIGHTTEXT};
			theme.light.values[static_cast<size_t>(uiColor::FrameSelected)] = {
				color::RebeccaPurple, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::SelectedBackground)] = {
				color::LightSteelBlue, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::Status)] = {color::RoyalBlue, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::PaneHeaderBackground)] = {
				color::AliceBlue, -1, COLOR_BTNFACE};
			theme.light.values[static_cast<size_t>(uiColor::PaneHeaderText)] = {color::MidnightBlue, -1, COLOR_BTNTEXT};
			theme.light.values[static_cast<size_t>(uiColor::TextHighlightBackground)] = {
				color::LightSteelBlue, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::TextHighlight)] = {
				color::MidnightBlue, -1, COLOR_HIGHLIGHTTEXT};
			theme.light.values[static_cast<size_t>(uiColor::MenuGutter)] = {color::AliceBlue, -1, COLOR_BTNFACE};

			theme.dark.values[static_cast<size_t>(uiColor::Glow)] = {color::LightSkyBlue, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::GlowBackground)] = {
				color::MidnightBlue, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::FrameSelected)] = {color::Plum, -1, COLOR_WINDOWTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::SelectedBackground)] = {
				color::DarkSlateBlue, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::Status)] = {color::MidnightBlue, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::PaneHeaderBackground)] = {
				color::MidnightBlue, -1, COLOR_BTNFACE};
			theme.dark.values[static_cast<size_t>(uiColor::TextHighlightBackground)] = {
				color::DarkSlateBlue, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::MenuGutter)] = {color::MidnightBlue, -1, COLOR_BTNFACE};

			return theme;
		}(),
	}};
} // namespace ui::theme::builtin
