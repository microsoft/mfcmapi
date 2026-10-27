#pragma once
#include <Windows.h>

namespace ui::color
{
	// Standard CSS names always retain their exact RGB values.
	inline constexpr COLORREF Black = RGB(0x00, 0x00, 0x00);
	inline constexpr COLORREF White = RGB(0xFF, 0xFF, 0xFF);
	inline constexpr COLORREF LightGray = RGB(0xD3, 0xD3, 0xD3);
	inline constexpr COLORREF Gray = RGB(0x80, 0x80, 0x80);
	inline constexpr COLORREF DimGray = RGB(0x69, 0x69, 0x69);
	inline constexpr COLORREF Cyan = RGB(0x00, 0xFF, 0xFF);
	inline constexpr COLORREF Magenta = RGB(0xFF, 0x00, 0xFF);
	inline constexpr COLORREF Pink = RGB(0xFF, 0xC0, 0xCB);
	inline constexpr COLORREF Lavender = RGB(0xE6, 0xE6, 0xFA);
	inline constexpr COLORREF IndianRed = RGB(0xCD, 0x5C, 0x5C);
	inline constexpr COLORREF Lime = RGB(0x00, 0xFF, 0x00);
	inline constexpr COLORREF Coral = RGB(0xFF, 0x7F, 0x50);
	inline constexpr COLORREF DarkGreen = RGB(0x00, 0x64, 0x00);
	inline constexpr COLORREF Honeydew = RGB(0xF0, 0xFF, 0xF0);
	inline constexpr COLORREF PaleGreen = RGB(0x98, 0xFB, 0x98);
	inline constexpr COLORREF LightGreen = RGB(0x90, 0xEE, 0x90);
	inline constexpr COLORREF DarkOliveGreen = RGB(0x55, 0x6B, 0x2F);
	inline constexpr COLORREF LightSkyBlue = RGB(0x87, 0xCE, 0xFA);
	inline constexpr COLORREF Teal = RGB(0x00, 0x80, 0x80);
	inline constexpr COLORREF Azure = RGB(0xF0, 0xFF, 0xFF);
	inline constexpr COLORREF PowderBlue = RGB(0xB0, 0xE0, 0xE6);
	inline constexpr COLORREF MidnightBlue = RGB(0x19, 0x19, 0x70);
	inline constexpr COLORREF Chocolate = RGB(0xD2, 0x69, 0x1E);
	inline constexpr COLORREF Turquoise = RGB(0x40, 0xE0, 0xD0);
	inline constexpr COLORREF DarkSlateGray = RGB(0x2F, 0x4F, 0x4F);
	inline constexpr COLORREF PeachPuff = RGB(0xFF, 0xDA, 0xB9);
	inline constexpr COLORREF SaddleBrown = RGB(0x8B, 0x45, 0x13);
	inline constexpr COLORREF OldLace = RGB(0xFD, 0xF5, 0xE6);
	inline constexpr COLORREF Wheat = RGB(0xF5, 0xDE, 0xB3);
	inline constexpr COLORREF Maroon = RGB(0x80, 0x00, 0x00);
	inline constexpr COLORREF SteelBlue = RGB(0x46, 0x82, 0xB4);
	inline constexpr COLORREF SandyBrown = RGB(0xF4, 0xA4, 0x60);
	inline constexpr COLORREF LightSteelBlue = RGB(0xB0, 0xC4, 0xDE);
	inline constexpr COLORREF RoyalBlue = RGB(0x41, 0x69, 0xE1);
	inline constexpr COLORREF AliceBlue = RGB(0xF0, 0xF8, 0xFF);
	inline constexpr COLORREF RebeccaPurple = RGB(0x66, 0x33, 0x99);
	inline constexpr COLORREF DarkSlateBlue = RGB(0x48, 0x3D, 0x8B);
	inline constexpr COLORREF Plum = RGB(0xDD, 0xA0, 0xDD);

	// Application-specific shades are not standard named colors.
	namespace custom
	{
		inline constexpr COLORREF HeroBlue = RGB(0x00, 0x72, 0xC6);
		inline constexpr COLORREF HeroPaleBlue = RGB(0xE6, 0xF2, 0xFA);
		inline constexpr COLORREF HeroSelectionBlue = RGB(0xCD, 0xE6, 0xF7);
		inline constexpr COLORREF HeroDarkBlue = RGB(0x09, 0x47, 0x71);
		inline constexpr COLORREF HeroSelectionDarkBlue = RGB(0x26, 0x4F, 0x78);
		inline constexpr COLORREF CharcoalGray = RGB(0x2D, 0x2D, 0x30);
	} // namespace custom
} // namespace ui::color
