#pragma once
#include <UI/Themes/Theme.h>
#include <UI/Themes/Colors.h>

namespace ui::theme::builtin
{
	// Complete light/dark palettes follow uiColor order.
	inline const ThemeDefinition Standard = {
		L"standard",
		L"Standard",
		{ScheduleKind::None, 0, 0, 0, 0},
		{{{
			{color::White, COLOR_WINDOW, COLOR_WINDOW}, // Background
			{color::LightGray, -1, COLOR_BTNFACE}, // BackgroundReadOnly
			{color::custom::HeroBlue, -1, COLOR_HIGHLIGHT}, // Glow
			{color::custom::HeroMediumBlue, -1, COLOR_HIGHLIGHT}, // GlowBackground
			{color::White, -1, COLOR_HIGHLIGHTTEXT}, // GlowText
			{color::Black, COLOR_WINDOWTEXT, COLOR_WINDOWTEXT}, // FrameSelected
			{color::Gray, COLOR_3DLIGHT, COLOR_GRAYTEXT}, // FrameUnselected
			{color::custom::HeroDeepBlue, -1, COLOR_HIGHLIGHT}, // SelectedBackground
			{color::LightGray, -1, COLOR_BTNFACE}, // DisabledSelectedBackground
			{color::custom::HeroBlue, -1, COLOR_HIGHLIGHT}, // DisabledSelectedBorder
			{color::Gray, COLOR_GRAYTEXT, COLOR_WINDOWTEXT}, // Arrow
			{color::Black, COLOR_WINDOWTEXT, COLOR_WINDOWTEXT}, // Text
			{color::Gray, COLOR_GRAYTEXT, COLOR_GRAYTEXT}, // TextDisabled
			{color::custom::CarbonGray, -1, COLOR_GRAYTEXT}, // TextDisabledSelected
			{color::Black, -1, COLOR_WINDOWTEXT}, // TextReadOnly
			{color::Magenta, -1, -1}, // BitmapTransBack
			{color::Cyan, -1, -1}, // BitmapTransFore
			{color::custom::HeroBlue, -1, COLOR_HIGHLIGHT}, // Status
			{color::White, -1, COLOR_HIGHLIGHTTEXT}, // StatusText
			{color::custom::HeroPaleBlue, -1, COLOR_BTNFACE}, // PaneHeaderBackground
			{color::Black, -1, COLOR_BTNTEXT}, // PaneHeaderText
			{color::custom::HeroDeepBlue, -1, COLOR_HIGHLIGHT}, // TextHighlightBackground
			{color::White, -1, COLOR_HIGHLIGHTTEXT}, // TextHighlight
			{color::LightGray, -1, COLOR_BTNFACE}, // MenuGutter
			{color::Pink, -1, -1}, // TestPink
			{color::Lavender, -1, -1}, // TestLavender
			{color::IndianRed, -1, -1}, // TestRed
			{color::Lime, -1, -1}, // TestGreen
			{color::Coral, -1, -1}, // TestOrange
			{color::White, COLOR_WINDOW, COLOR_WINDOW}, // TitleBarBackground
			{color::Black, COLOR_WINDOWTEXT, COLOR_WINDOWTEXT}, // TitleBarText
			{color::Black, COLOR_WINDOWTEXT, COLOR_WINDOWTEXT}, // TitleBarButton
			{color::White, -1, COLOR_WINDOWTEXT}, // TitleBarButtonHover
			{color::White, COLOR_WINDOW, COLOR_WINDOW}, // ColumnHeaderBackground
			{color::Black, COLOR_WINDOWTEXT, COLOR_WINDOWTEXT}, // ColumnHeaderText
			{color::Gray, COLOR_GRAYTEXT, COLOR_WINDOWTEXT}, // ColumnHeaderArrow
			{color::White, COLOR_WINDOW, COLOR_WINDOW}, // MenuBackground
			{color::Black, COLOR_WINDOWTEXT, COLOR_WINDOWTEXT}, // MenuText
			{color::Gray, COLOR_GRAYTEXT, COLOR_GRAYTEXT}, // MenuTextDisabled
			{color::Black, -1, COLOR_WINDOWTEXT}, // StatusTextOutline
		}}},
		{{{
			{color::custom::GraphiteGray, -1, COLOR_WINDOW}, // Background
			{color::custom::CharcoalGray, -1, COLOR_BTNFACE}, // BackgroundReadOnly
			{color::custom::HeroBlue, -1, COLOR_HIGHLIGHT}, // Glow
			{color::custom::HeroSkyBlue, -1, COLOR_HIGHLIGHT}, // GlowBackground
			{color::Black, -1, COLOR_HIGHLIGHTTEXT}, // GlowText
			{color::LightGray, -1, COLOR_WINDOWTEXT}, // FrameSelected
			{color::custom::CarbonGray, -1, COLOR_GRAYTEXT}, // FrameUnselected
			{color::custom::HeroIceBlue, -1, COLOR_HIGHLIGHT}, // SelectedBackground
			{color::custom::CharcoalGray, -1, COLOR_BTNFACE}, // DisabledSelectedBackground
			{color::LightSkyBlue, -1, COLOR_HIGHLIGHT}, // DisabledSelectedBorder
			{color::custom::SilverGray, -1, COLOR_WINDOWTEXT}, // Arrow
			{color::custom::MistWhite, -1, COLOR_WINDOWTEXT}, // Text
			{color::custom::StoneGray, -1, COLOR_GRAYTEXT}, // TextDisabled
			{color::custom::PewterGray, -1, COLOR_GRAYTEXT}, // TextDisabledSelected
			{color::custom::PaleGray, -1, COLOR_WINDOWTEXT}, // TextReadOnly
			{color::Magenta, -1, -1}, // BitmapTransBack
			{color::Cyan, -1, -1}, // BitmapTransFore
			{color::custom::HeroBlue, -1, COLOR_HIGHLIGHT}, // Status
			{color::White, -1, COLOR_HIGHLIGHTTEXT}, // StatusText
			{color::custom::HeroDarkBlue, -1, COLOR_BTNFACE}, // PaneHeaderBackground
			{color::White, -1, COLOR_BTNTEXT}, // PaneHeaderText
			{color::custom::HeroIceBlue, -1, COLOR_HIGHLIGHT}, // TextHighlightBackground
			{color::Black, -1, COLOR_HIGHLIGHTTEXT}, // TextHighlight
			{color::custom::CharcoalGray, -1, COLOR_BTNFACE}, // MenuGutter
			{color::Pink, -1, -1}, // TestPink
			{color::Lavender, -1, -1}, // TestLavender
			{color::IndianRed, -1, -1}, // TestRed
			{color::Lime, -1, -1}, // TestGreen
			{color::Coral, -1, -1}, // TestOrange
			{color::custom::GraphiteGray, -1, COLOR_WINDOW}, // TitleBarBackground
			{color::custom::MistWhite, -1, COLOR_WINDOWTEXT}, // TitleBarText
			{color::LightGray, -1, COLOR_WINDOWTEXT}, // TitleBarButton
			{color::Black, -1, COLOR_WINDOWTEXT}, // TitleBarButtonHover
			{color::custom::GraphiteGray, -1, COLOR_WINDOW}, // ColumnHeaderBackground
			{color::custom::MistWhite, -1, COLOR_WINDOWTEXT}, // ColumnHeaderText
			{color::custom::SilverGray, -1, COLOR_WINDOWTEXT}, // ColumnHeaderArrow
			{color::custom::GraphiteGray, -1, COLOR_WINDOW}, // MenuBackground
			{color::custom::MistWhite, -1, COLOR_WINDOWTEXT}, // MenuText
			{color::custom::StoneGray, -1, COLOR_GRAYTEXT}, // MenuTextDisabled
			{color::Black, -1, COLOR_WINDOWTEXT}, // StatusTextOutline
		}}}};

	inline const std::array<ThemeDefinition, 5> Definitions = {{
		Standard,
		[] {
			auto theme = Standard;
			theme.stableId = L"season.spring";
			theme.name = L"Spring";
			theme.schedule = {ScheduleKind::AnnualDateRange, 3, 1, 5, 31, 0, true};

			theme.light.values[static_cast<size_t>(uiColor::Background)] = {color::Honeydew, -1, COLOR_WINDOW};
			theme.light.values[static_cast<size_t>(uiColor::Glow)] = {color::DarkGreen, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::GlowBackground)] = {
				RGB(0x24, 0x71, 0x3C), -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::GlowText)] = {color::White, -1, COLOR_HIGHLIGHTTEXT};
			theme.light.values[static_cast<size_t>(uiColor::FrameSelected)] = {color::SteelBlue, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::SelectedBackground)] = {
				color::custom::PineGreen, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::Status)] = {color::DarkGreen, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::PaneHeaderBackground)] = {
				color::Honeydew, -1, COLOR_BTNFACE};
			theme.light.values[static_cast<size_t>(uiColor::PaneHeaderText)] = {color::DarkGreen, -1, COLOR_BTNTEXT};
			theme.light.values[static_cast<size_t>(uiColor::TextHighlightBackground)] = {
				color::custom::PineGreen, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::TextHighlight)] = {color::White, -1, COLOR_HIGHLIGHTTEXT};
			theme.light.values[static_cast<size_t>(uiColor::MenuGutter)] = {color::Honeydew, -1, COLOR_BTNFACE};

			theme.dark.values[static_cast<size_t>(uiColor::Glow)] = {color::LightGreen, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::GlowBackground)] = {color::LightGreen, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::GlowText)] = {color::Black, -1, COLOR_HIGHLIGHTTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::FrameSelected)] = {
				color::LightSkyBlue, -1, COLOR_WINDOWTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::SelectedBackground)] = {
				color::PaleGreen, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::Status)] = {color::DarkGreen, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::PaneHeaderBackground)] = {
				color::DarkGreen, -1, COLOR_BTNFACE};
			theme.dark.values[static_cast<size_t>(uiColor::TextHighlightBackground)] = {
				color::PaleGreen, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::TextHighlight)] = {color::Black, -1, COLOR_HIGHLIGHTTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::MenuGutter)] = {color::DarkGreen, -1, COLOR_BTNFACE};

			theme.light.values[static_cast<size_t>(uiColor::TextDisabledSelected)] = {
				RGB(0x5A, 0x5A, 0x5A), -1, COLOR_GRAYTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::TextDisabledSelected)] = {
				RGB(0x94, 0x94, 0x94), -1, COLOR_GRAYTEXT};
			theme.light.values[static_cast<size_t>(uiColor::DisabledSelectedBorder)] = {
				color::DarkGreen, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::DisabledSelectedBorder)] = {
				color::LightGreen, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::TitleBarBackground)] = {
				RGB(0x31, 0x5D, 0x85), -1, COLOR_WINDOW};
			theme.light.values[static_cast<size_t>(uiColor::TitleBarText)] = {color::White, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::TitleBarButton)] = {color::White, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::TitleBarButtonHover)] = {
				color::White, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::ColumnHeaderBackground)] = {
				color::PaleGreen, -1, COLOR_WINDOW};
			theme.light.values[static_cast<size_t>(uiColor::ColumnHeaderText)] = {
				color::DarkGreen, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::ColumnHeaderArrow)] = {
				color::DarkGreen, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::MenuBackground)] = {color::Honeydew, -1, COLOR_WINDOW};
			theme.dark.values[static_cast<size_t>(uiColor::TitleBarBackground)] = {
				RGB(0x3D, 0x4F, 0x87), -1, COLOR_WINDOW};
			theme.dark.values[static_cast<size_t>(uiColor::MenuBackground)] = {RGB(0x14, 0x23, 0x14), -1, COLOR_WINDOW};
			theme.dark.values[static_cast<size_t>(uiColor::TitleBarButtonHover)] = {color::Black, -1, COLOR_WINDOWTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::ColumnHeaderBackground)] = {
				color::DarkGreen, -1, COLOR_WINDOW};
			theme.dark.values[static_cast<size_t>(uiColor::ColumnHeaderText)] = {color::White, -1, COLOR_WINDOWTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::ColumnHeaderArrow)] = {
				color::LightGreen, -1, COLOR_WINDOWTEXT};
			return theme;
		}(),
		[] {
			auto theme = Standard;
			theme.stableId = L"season.summer";
			theme.name = L"Summer";
			theme.schedule = {ScheduleKind::AnnualDateRange, 6, 1, 8, 31, 0, true};

			theme.light.values[static_cast<size_t>(uiColor::Background)] = {color::Azure, -1, COLOR_WINDOW};
			theme.light.values[static_cast<size_t>(uiColor::Glow)] = {color::Teal, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::GlowBackground)] = {
				RGB(0x15, 0x6E, 0x7A), -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::GlowText)] = {color::White, -1, COLOR_HIGHLIGHTTEXT};
			theme.light.values[static_cast<size_t>(uiColor::FrameSelected)] = {color::Chocolate, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::SelectedBackground)] = {
				color::custom::DeepTeal, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::Status)] = {color::Teal, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::PaneHeaderBackground)] = {color::Azure, -1, COLOR_BTNFACE};
			theme.light.values[static_cast<size_t>(uiColor::PaneHeaderText)] = {color::MidnightBlue, -1, COLOR_BTNTEXT};
			theme.light.values[static_cast<size_t>(uiColor::TextHighlightBackground)] = {
				color::custom::DeepTeal, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::TextHighlight)] = {color::White, -1, COLOR_HIGHLIGHTTEXT};
			theme.light.values[static_cast<size_t>(uiColor::MenuGutter)] = {color::Azure, -1, COLOR_BTNFACE};

			theme.dark.values[static_cast<size_t>(uiColor::Glow)] = {color::Turquoise, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::GlowBackground)] = {
				color::custom::SoftTurquoise, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::GlowText)] = {color::Black, -1, COLOR_HIGHLIGHTTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::FrameSelected)] = {color::PeachPuff, -1, COLOR_WINDOWTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::SelectedBackground)] = {
				color::LightCyan, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::Status)] = {color::DarkSlateGray, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::PaneHeaderBackground)] = {
				color::DarkSlateGray, -1, COLOR_BTNFACE};
			theme.dark.values[static_cast<size_t>(uiColor::TextHighlightBackground)] = {
				color::custom::SoftTurquoise, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::TextHighlight)] = {color::Black, -1, COLOR_HIGHLIGHTTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::MenuGutter)] = {color::DarkSlateGray, -1, COLOR_BTNFACE};

			theme.light.values[static_cast<size_t>(uiColor::TextDisabledSelected)] = {
				RGB(0x5A, 0x5A, 0x5A), -1, COLOR_GRAYTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::TextDisabledSelected)] = {
				RGB(0x94, 0x94, 0x94), -1, COLOR_GRAYTEXT};
			theme.light.values[static_cast<size_t>(uiColor::DisabledSelectedBorder)] = {
				color::Teal, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::DisabledSelectedBorder)] = {
				color::Turquoise, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::TitleBarBackground)] = {
				color::SaddleBrown, -1, COLOR_WINDOW};
			theme.light.values[static_cast<size_t>(uiColor::TitleBarText)] = {color::White, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::TitleBarButton)] = {color::White, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::TitleBarButtonHover)] = {
				color::White, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::ColumnHeaderBackground)] = {
				color::PowderBlue, -1, COLOR_WINDOW};
			theme.light.values[static_cast<size_t>(uiColor::ColumnHeaderText)] = {
				color::MidnightBlue, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::ColumnHeaderArrow)] = {color::Teal, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::MenuBackground)] = {color::Azure, -1, COLOR_WINDOW};
			theme.dark.values[static_cast<size_t>(uiColor::TitleBarBackground)] = {
				RGB(0x75, 0x45, 0x2B), -1, COLOR_WINDOW};
			theme.dark.values[static_cast<size_t>(uiColor::MenuBackground)] = {RGB(0x14, 0x26, 0x28), -1, COLOR_WINDOW};
			theme.dark.values[static_cast<size_t>(uiColor::TitleBarButtonHover)] = {color::Black, -1, COLOR_WINDOWTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::ColumnHeaderBackground)] = {
				color::DarkSlateGray, -1, COLOR_WINDOW};
			theme.dark.values[static_cast<size_t>(uiColor::ColumnHeaderText)] = {color::White, -1, COLOR_WINDOWTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::ColumnHeaderArrow)] = {
				color::Turquoise, -1, COLOR_WINDOWTEXT};
			return theme;
		}(),
		[] {
			auto theme = Standard;
			theme.stableId = L"season.autumn";
			theme.name = L"Autumn";
			theme.schedule = {ScheduleKind::AnnualDateRange, 9, 1, 11, 30, 0, true};

			theme.light.values[static_cast<size_t>(uiColor::Background)] = {color::OldLace, -1, COLOR_WINDOW};
			theme.light.values[static_cast<size_t>(uiColor::Glow)] = {color::SaddleBrown, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::GlowBackground)] = {
				RGB(0x98, 0x5A, 0x11), -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::GlowText)] = {color::White, -1, COLOR_HIGHLIGHTTEXT};
			theme.light.values[static_cast<size_t>(uiColor::FrameSelected)] = {color::SteelBlue, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::SelectedBackground)] = {
				color::custom::ChestnutBrown, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::Status)] = {color::SaddleBrown, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::PaneHeaderBackground)] = {
				color::OldLace, -1, COLOR_BTNFACE};
			theme.light.values[static_cast<size_t>(uiColor::PaneHeaderText)] = {color::Maroon, -1, COLOR_BTNTEXT};
			theme.light.values[static_cast<size_t>(uiColor::TextHighlightBackground)] = {
				color::custom::ChestnutBrown, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::TextHighlight)] = {color::White, -1, COLOR_HIGHLIGHTTEXT};
			theme.light.values[static_cast<size_t>(uiColor::MenuGutter)] = {color::OldLace, -1, COLOR_BTNFACE};

			theme.dark.values[static_cast<size_t>(uiColor::Glow)] = {color::SandyBrown, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::GlowBackground)] = {
				RGB(0xE0, 0xA4, 0x6D), -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::GlowText)] = {color::Black, -1, COLOR_HIGHLIGHTTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::FrameSelected)] = {
				color::LightSteelBlue, -1, COLOR_WINDOWTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::SelectedBackground)] = {
				color::custom::PaleApricot, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::Status)] = {color::Maroon, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::PaneHeaderBackground)] = {color::Maroon, -1, COLOR_BTNFACE};
			theme.dark.values[static_cast<size_t>(uiColor::TextHighlightBackground)] = {
				color::custom::PaleApricot, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::TextHighlight)] = {color::Black, -1, COLOR_HIGHLIGHTTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::MenuGutter)] = {color::Maroon, -1, COLOR_BTNFACE};

			theme.light.values[static_cast<size_t>(uiColor::TextDisabledSelected)] = {
				RGB(0x5A, 0x5A, 0x5A), -1, COLOR_GRAYTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::TextDisabledSelected)] = {
				RGB(0x94, 0x94, 0x94), -1, COLOR_GRAYTEXT};
			theme.light.values[static_cast<size_t>(uiColor::DisabledSelectedBorder)] = {
				color::SaddleBrown, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::DisabledSelectedBorder)] = {
				color::SandyBrown, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::TitleBarBackground)] = {
				RGB(0x31, 0x5D, 0x85), -1, COLOR_WINDOW};
			theme.light.values[static_cast<size_t>(uiColor::TitleBarText)] = {color::White, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::TitleBarButton)] = {color::White, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::TitleBarButtonHover)] = {
				color::White, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::ColumnHeaderBackground)] = {color::Wheat, -1, COLOR_WINDOW};
			theme.light.values[static_cast<size_t>(uiColor::ColumnHeaderText)] = {color::Maroon, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::ColumnHeaderArrow)] = {
				color::SaddleBrown, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::MenuBackground)] = {color::OldLace, -1, COLOR_WINDOW};
			theme.dark.values[static_cast<size_t>(uiColor::TitleBarBackground)] = {
				RGB(0x3F, 0x52, 0x78), -1, COLOR_WINDOW};
			theme.dark.values[static_cast<size_t>(uiColor::TitleBarButtonHover)] = {color::Black, -1, COLOR_WINDOWTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::MenuBackground)] = {RGB(0x28, 0x1B, 0x18), -1, COLOR_WINDOW};
			theme.dark.values[static_cast<size_t>(uiColor::ColumnHeaderBackground)] = {color::Maroon, -1, COLOR_WINDOW};
			theme.dark.values[static_cast<size_t>(uiColor::ColumnHeaderText)] = {color::White, -1, COLOR_WINDOWTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::ColumnHeaderArrow)] = {
				color::SandyBrown, -1, COLOR_WINDOWTEXT};
			return theme;
		}(),
		[] {
			auto theme = Standard;
			theme.stableId = L"season.winter";
			theme.name = L"Winter";
			theme.schedule = {ScheduleKind::AnnualDateRange, 12, 1, 2, 29, 0, true};

			theme.light.values[static_cast<size_t>(uiColor::Background)] = {color::AliceBlue, -1, COLOR_WINDOW};
			theme.light.values[static_cast<size_t>(uiColor::Glow)] = {color::RoyalBlue, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::GlowBackground)] = {
				RGB(0x31, 0x5F, 0x96), -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::GlowText)] = {color::White, -1, COLOR_HIGHLIGHTTEXT};
			theme.light.values[static_cast<size_t>(uiColor::FrameSelected)] = {
				color::RebeccaPurple, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::SelectedBackground)] = {
				color::custom::DeepViolet, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::Status)] = {color::RoyalBlue, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::PaneHeaderBackground)] = {
				color::AliceBlue, -1, COLOR_BTNFACE};
			theme.light.values[static_cast<size_t>(uiColor::PaneHeaderText)] = {color::MidnightBlue, -1, COLOR_BTNTEXT};
			theme.light.values[static_cast<size_t>(uiColor::TextHighlightBackground)] = {
				color::custom::DeepViolet, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::TextHighlight)] = {color::White, -1, COLOR_HIGHLIGHTTEXT};
			theme.light.values[static_cast<size_t>(uiColor::MenuGutter)] = {color::AliceBlue, -1, COLOR_BTNFACE};

			theme.dark.values[static_cast<size_t>(uiColor::Glow)] = {color::LightSkyBlue, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::GlowBackground)] = {
				RGB(0x82, 0xAA, 0xD4), -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::GlowText)] = {color::Black, -1, COLOR_HIGHLIGHTTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::FrameSelected)] = {color::Plum, -1, COLOR_WINDOWTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::SelectedBackground)] = {
				color::custom::FrostedLilac, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::Status)] = {color::MidnightBlue, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::PaneHeaderBackground)] = {
				color::MidnightBlue, -1, COLOR_BTNFACE};
			theme.dark.values[static_cast<size_t>(uiColor::TextHighlightBackground)] = {
				color::custom::FrostedLilac, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::TextHighlight)] = {color::Black, -1, COLOR_HIGHLIGHTTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::MenuGutter)] = {color::MidnightBlue, -1, COLOR_BTNFACE};

			theme.light.values[static_cast<size_t>(uiColor::TextDisabledSelected)] = {
				RGB(0x5A, 0x5A, 0x5A), -1, COLOR_GRAYTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::TextDisabledSelected)] = {
				RGB(0x94, 0x94, 0x94), -1, COLOR_GRAYTEXT};
			theme.light.values[static_cast<size_t>(uiColor::DisabledSelectedBorder)] = {
				color::RoyalBlue, -1, COLOR_HIGHLIGHT};
			theme.dark.values[static_cast<size_t>(uiColor::DisabledSelectedBorder)] = {
				color::LightSkyBlue, -1, COLOR_HIGHLIGHT};
			theme.light.values[static_cast<size_t>(uiColor::TitleBarBackground)] = {
				color::RebeccaPurple, -1, COLOR_WINDOW};
			theme.light.values[static_cast<size_t>(uiColor::TitleBarText)] = {color::White, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::TitleBarButton)] = {color::White, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::TitleBarButtonHover)] = {
				color::White, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::ColumnHeaderBackground)] = {
				color::LightSteelBlue, -1, COLOR_WINDOW};
			theme.light.values[static_cast<size_t>(uiColor::ColumnHeaderText)] = {
				color::MidnightBlue, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::ColumnHeaderArrow)] = {
				color::MidnightBlue, -1, COLOR_WINDOWTEXT};
			theme.light.values[static_cast<size_t>(uiColor::MenuBackground)] = {color::AliceBlue, -1, COLOR_WINDOW};
			theme.dark.values[static_cast<size_t>(uiColor::TitleBarBackground)] = {
				RGB(0x4F, 0x36, 0x70), -1, COLOR_WINDOW};
			theme.dark.values[static_cast<size_t>(uiColor::TitleBarButtonHover)] = {color::Black, -1, COLOR_WINDOWTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::MenuBackground)] = {RGB(0x18, 0x1C, 0x30), -1, COLOR_WINDOW};
			theme.dark.values[static_cast<size_t>(uiColor::ColumnHeaderBackground)] = {
				color::MidnightBlue, -1, COLOR_WINDOW};
			theme.dark.values[static_cast<size_t>(uiColor::ColumnHeaderText)] = {color::White, -1, COLOR_WINDOWTEXT};
			theme.dark.values[static_cast<size_t>(uiColor::ColumnHeaderArrow)] = {
				color::LightSkyBlue, -1, COLOR_WINDOWTEXT};
			return theme;
		}(),
	}};
} // namespace ui::theme::builtin
