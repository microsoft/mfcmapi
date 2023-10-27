// Common UI functions for MFCMAPI
#include <StdAfx.h>
#include <UI/UIFunctions.h>
#include <windowsx.h>
#include <UI/RichEditOleCallback.h>
#include <UI/ViewPane/CheckPane.h>
#include <UI/DoubleBuffer.h>
#include <UI/addinui.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/addin/mfcmapi.h>
#include <core/utility/registry.h>

namespace ui
{
	HFONT g_hFontSegoe = nullptr;
	HFONT g_hFontSegoeBold = nullptr;

#define SEGOE _T("Segoe UI") // STRING_OK
#define SEGOEW L"Segoe UI" // STRING_OK
#define SEGOEBOLD L"Segoe UI Bold" // STRING_OK
#define BUTTON_STYLE _T("ButtonStyle") // STRING_OK
#define LABEL_STYLE _T("LabelStyle") // STRING_OK

#define BORDER_VISIBLEWIDTH 2

	enum class myColor
	{
		White,
		LightGrey,
		Grey,
		DarkGrey,
		Black,
		Cyan,
		Magenta,
		Blue,
		MedBlue,
		PaleBlue,
		Pink,
		Lavender,
		Red,
		Green,
		Orange,
		ColorEnd
	};

	// Keep in sync with enum myColor
	COLORREF g_Colors[static_cast<int>(myColor::ColorEnd)] = {
		RGB(0xFF, 0xFF, 0xFF), // cWhite
		RGB(0xD3, 0xD3, 0xD3), // cLightGrey
		RGB(0xAD, 0xAC, 0xAE), // cGrey
		RGB(0x64, 0x64, 0x64), // cDarkGrey
		RGB(0x00, 0x00, 0x00), // cBlack
		RGB(0x00, 0xFF, 0xFF), // cCyan
		RGB(0xFF, 0x00, 0xFF), // cMagenta
		RGB(0x00, 0x72, 0xC6), // cBlue
		RGB(0xCD, 0xE6, 0xF7), // cMedBlue
		RGB(0xE6, 0xF2, 0xFA), // cPaleBlue
		RGB(0xFF, 0xC0, 0xCB), // cPink
		RGB(0xE6, 0xE6, 0xFA), // cLavender
		RGB(0xCD, 0x5C, 0x5C), // cRed
		RGB(0x00, 0xFF, 0x00), // cGreen
		RGB(0xFF, 0x7F, 0x50), // cOrange
	};

	// Fixed mapping of UI elements to colors
	// Will be overridden by system colors when specified in g_SysColors
	// Keep in sync with enum uiColor
	// We can swap in cCyan for various entries to test coverage
	myColor g_FixedColors[static_cast<int>(uiColor::UIEnd)] = {
		myColor::White, // cBackground
		myColor::LightGrey, // cBackgroundReadOnly
		myColor::Blue, // cGlow
		myColor::PaleBlue, // cGlowBackground
		myColor::Black, // cGlowText
		myColor::Black, // cFrameSelected
		myColor::Grey, // cFrameUnselected
		myColor::MedBlue, // cSelectedBackground
		myColor::Grey, // cArrow
		myColor::Black, // cText
		myColor::Grey, // cTextDisabled
		myColor::Black, // cTextReadOnly
		myColor::Magenta, // cBitmapTransBack
		myColor::Cyan, // cBitmapTransFore
		myColor::Blue, // cStatus
		myColor::White, // cStatusText
		myColor::PaleBlue, // cPaneHeaderBackground,
		myColor::Black, // cPaneHeaderText,
		myColor::MedBlue, // cTextHighlightBackground,
		myColor::Black, // cTextHighlight,
		myColor::Pink, // cTestPink,
		myColor::Lavender, // cTestLavender,
		myColor::Red, // cTestRed,
		myColor::Green, // cTestGreen,
		myColor::Orange, // cTestOrange,
	};

	// Mapping of UI elements to system colors
	// NULL entries will get the fixed mapping from g_FixedColors
	int g_SysColors[static_cast<int>(uiColor::UIEnd)] = {
		COLOR_WINDOW, // cBackground
		NULL, // cBackgroundReadOnly
		NULL, // cGlow
		NULL, // cGlowBackground
		NULL, // cGlowText
		COLOR_WINDOWTEXT, // cFrameSelected
		COLOR_3DLIGHT, // cFrameUnselected
		NULL,
		COLOR_GRAYTEXT, // cArrow
		COLOR_WINDOWTEXT, // cText
		COLOR_GRAYTEXT, // cTextDisabled
		NULL, // cTextReadOnly
		NULL, // cBitmapTransBack
		NULL, // cBitmapTransFore
		NULL, // cStatus
		NULL, // cStatusText
	};

	// Mapping of bitmap resources to constants
	// NULL entries will get the fixed mapping from g_FixedColors
	int g_BitmapResources[static_cast<int>(uiBitmap::BitmapEnd)] = {
		IDB_ADVISE, // cNotify,
		IDB_CLOSE, // cClose,
		IDB_MINIMIZE, // cMinimize,
		IDB_MAXIMIZE, // cMaximize,
		IDB_RESTORE, // cRestore,
	};

	HBRUSH g_FixedBrushes[static_cast<int>(myColor::ColorEnd)] = {};
	HBRUSH g_SysBrushes[static_cast<int>(uiColor::UIEnd)] = {};
	HPEN g_Pens[static_cast<int>(uiPen::PenEnd)] = {};
	HBITMAP g_Bitmaps[static_cast<int>(uiBitmap::BitmapEnd)] = {};

	void InitializeGDI() noexcept {}

	void UninitializeGDI() noexcept
	{
		if (g_hFontSegoe) DeleteObject(g_hFontSegoe);
		g_hFontSegoe = nullptr;
		if (g_hFontSegoeBold) DeleteObject(g_hFontSegoeBold);
		g_hFontSegoeBold = nullptr;

		for (auto& fixedBrush : g_FixedBrushes)
		{
			if (fixedBrush)
			{
				DeleteObject(fixedBrush);
				fixedBrush = nullptr;
			}
		}

		for (auto& sysBrush : g_SysBrushes)
		{
			if (sysBrush)
			{
				DeleteObject(sysBrush);
				sysBrush = nullptr;
			}
		}

		for (auto& g_Pen : g_Pens)
		{
			if (g_Pen)
			{
				DeleteObject(g_Pen);
				g_Pen = nullptr;
			}
		}

		for (auto& bitmap : g_Bitmaps)
		{
			if (bitmap)
			{
				DeleteObject(bitmap);
				bitmap = nullptr;
			}
		}
	}

	_Check_return_ LPMENUENTRY CreateMenuEntry(_In_ const std::wstring& szMenu)
	{
		auto lpMenu = new (std::nothrow) MenuEntry;
		if (lpMenu)
		{
			lpMenu->m_MSAA.dwMSAASignature = MSAA_MENU_SIG;
			lpMenu->m_bOnMenuBar = false;
			lpMenu->m_MSAA.cchWText = static_cast<DWORD>(szMenu.length() + 1);
			lpMenu->m_MSAA.pszWText = const_cast<LPWSTR>(strings::wstringToLPCWSTR(szMenu));
			lpMenu->m_pName = lpMenu->m_MSAA.pszWText;
			return lpMenu;
		}

		return nullptr;
	}

	_Check_return_ LPMENUENTRY CreateMenuEntry(const UINT iudMenu)
	{
		return CreateMenuEntry(strings::loadstring(iudMenu));
	}

	void DeleteMenuEntry(_In_ LPMENUENTRY lpMenu) noexcept
	{
		if (!lpMenu) return;
		delete[] lpMenu->m_MSAA.pszWText;
		delete lpMenu;
	}

	void DeleteMenuEntries(_In_ HMENU hMenu)
	{
		const UINT nCount = GetMenuItemCount(hMenu);
		if (nCount == static_cast<UINT>(-1)) return;

		for (UINT nPosition = 0; nPosition < nCount; nPosition++)
		{
			auto menuiteminfo = MENUITEMINFOW{sizeof(MENUITEMINFOW), MIIM_DATA | MIIM_SUBMENU};

			GetMenuItemInfoW(hMenu, nPosition, true, &menuiteminfo);
			if (menuiteminfo.dwItemData)
			{
				DeleteMenuEntry(reinterpret_cast<LPMENUENTRY>(menuiteminfo.dwItemData));
			}

			if (menuiteminfo.hSubMenu)
			{
				DeleteMenuEntries(menuiteminfo.hSubMenu);
			}
		}
	}

	// Walk through the menu structure and convert any string entries to owner draw using MENUENTRY
	void ConvertMenuOwnerDraw(_In_ HMENU hMenu, const bool bRoot)
	{
		const UINT nCount = GetMenuItemCount(hMenu);
		if (nCount == static_cast<UINT>(-1)) return;

		for (UINT nPosition = 0; nPosition < nCount; nPosition++)
		{
			auto menuiteminfo = MENUITEMINFOW{};
			WCHAR szMenu[128] = {0};
			menuiteminfo.cbSize = sizeof menuiteminfo;
			menuiteminfo.fMask = MIIM_STRING | MIIM_SUBMENU | MIIM_FTYPE;
			menuiteminfo.cch = _countof(szMenu);
			menuiteminfo.dwTypeData = szMenu;

			EC_B_S(::GetMenuItemInfoW(hMenu, nPosition, true, &menuiteminfo));
			const auto bOwnerDrawn = (menuiteminfo.fType & MF_OWNERDRAW) != 0;
			if (!bOwnerDrawn)
			{
				auto lpMenuEntry = CreateMenuEntry(szMenu);
				if (lpMenuEntry)
				{
					lpMenuEntry->m_bOnMenuBar = bRoot;
					menuiteminfo.fMask = MIIM_DATA | MIIM_FTYPE;
					menuiteminfo.fType |= MF_OWNERDRAW;
					menuiteminfo.dwItemData = reinterpret_cast<ULONG_PTR>(lpMenuEntry);
					EC_B_S(::SetMenuItemInfoW(hMenu, nPosition, true, &menuiteminfo));
				}
			}

			if (menuiteminfo.hSubMenu)
			{
				ConvertMenuOwnerDraw(menuiteminfo.hSubMenu, false);
			}
		}
	}

	void UpdateMenuString(_In_ HWND hWnd, UINT uiMenuTag, const UINT uidNewString)
	{
		auto szNewString = strings::loadstring(uidNewString);

		output::DebugPrint(
			output::dbgLevel::Menu,
			L"UpdateMenuString: Changing menu item 0x%X on window %p to \"%ws\"\n",
			uiMenuTag,
			hWnd,
			szNewString.c_str());
		const auto hMenu = GetMenu(hWnd);
		if (!hMenu) return;

		auto menuiteminfo = MENUITEMINFOW{sizeof(MENUITEMINFOW), MIIM_DATA | MIIM_FTYPE};

		EC_B_S(::GetMenuItemInfoW(hMenu, uiMenuTag, false, &menuiteminfo));
		auto lpMenuEntry = CreateMenuEntry(szNewString);
		if (lpMenuEntry)
		{
			// Clean up old owner draw string data if it exists
			DeleteMenuEntry(reinterpret_cast<LPMENUENTRY>(menuiteminfo.dwItemData));

			menuiteminfo.fMask = MIIM_DATA | MIIM_FTYPE;
			menuiteminfo.fType |= MF_OWNERDRAW;
			menuiteminfo.dwItemData = reinterpret_cast<ULONG_PTR>(lpMenuEntry);
			EC_B_S(::SetMenuItemInfoW(hMenu, uiMenuTag, false, &menuiteminfo));
		}
	}

	void MergeMenu(_In_ HMENU hMenuDestination, _In_ HMENU hMenuAdd)
	{
		GetMenuItemCount(hMenuDestination);
		auto iMenuDestItemCount = GetMenuItemCount(hMenuDestination);
		const auto iMenuAddItemCount = GetMenuItemCount(hMenuAdd);

		if (iMenuAddItemCount == 0) return;

		if (iMenuDestItemCount > 0) ::AppendMenu(hMenuDestination, MF_SEPARATOR, NULL, nullptr);

		WCHAR szMenu[128] = {0};
		for (auto iLoop = 0; iLoop < iMenuAddItemCount; iLoop++)
		{
			GetMenuStringW(hMenuAdd, iLoop, szMenu, _countof(szMenu), MF_BYPOSITION);

			const auto hSubMenu = GetSubMenu(hMenuAdd, iLoop);
			if (!hSubMenu)
			{
				const auto nState = GetMenuState(hMenuAdd, iLoop, MF_BYPOSITION);
				const auto nItemID = GetMenuItemID(hMenuAdd, iLoop);

				AppendMenuW(hMenuDestination, nState, nItemID, szMenu);
				iMenuDestItemCount++;
			}
			else
			{
				const auto iInsertPosDefault = GetMenuItemCount(hMenuDestination);

				auto hNewPopupMenu = CreatePopupMenu();

				MergeMenu(hNewPopupMenu, hSubMenu);

				InsertMenuW(
					hMenuDestination,
					iInsertPosDefault,
					MF_BYPOSITION | MF_POPUP | MF_ENABLED,
					reinterpret_cast<UINT_PTR>(hNewPopupMenu),
					szMenu);
				iMenuDestItemCount++;
			}
		}
	}

	void DisplayContextMenu(UINT uiClassMenu, const UINT uiControlMenu, _In_ HWND hWnd, const int x, const int y)
	{
		if (!uiClassMenu) uiClassMenu = IDR_MENU_DEFAULT_POPUP;
		const auto hPopup = CreateMenu();
		auto hContext = ::LoadMenu(nullptr, MAKEINTRESOURCE(uiClassMenu));

		::InsertMenu(hPopup, 0, MF_BYPOSITION | MF_POPUP, reinterpret_cast<UINT_PTR>(hContext), _T(""));

		const auto hRealPopup = GetSubMenu(hPopup, 0);

		if (hRealPopup)
		{
			if (uiControlMenu)
			{
				const auto hAppended = ::LoadMenu(nullptr, MAKEINTRESOURCE(uiControlMenu));
				if (hAppended)
				{
					MergeMenu(hRealPopup, hAppended);
					DestroyMenu(hAppended);
				}
			}

			if (IDR_MENU_PROPERTY_POPUP == uiClassMenu)
			{
				static_cast<void>(addinui::ExtendAddInMenu(hRealPopup, MENU_CONTEXT_PROPERTY));
			}

			ConvertMenuOwnerDraw(hRealPopup, false);
			TrackPopupMenu(hRealPopup, TPM_LEFTALIGN | TPM_RIGHTBUTTON, x, y, NULL, hWnd, nullptr);
			DeleteMenuEntries(hRealPopup);
		}

		DestroyMenu(hContext);
		DestroyMenu(hPopup);
	}

	HMENU LocateSubmenu(_In_ HMENU hMenu, const UINT uid)
	{
		const UINT nCount = GetMenuItemCount(hMenu);
		if (nCount == static_cast<UINT>(-1)) return nullptr;

		for (UINT nPosition = 0; nPosition < nCount; nPosition++)
		{
			auto menuiteminfo = MENUITEMINFOW{sizeof(MENUITEMINFOW), MIIM_SUBMENU | MIIM_ID};
			WC_B_S(GetMenuItemInfoW(hMenu, nPosition, true, &menuiteminfo));

			if (menuiteminfo.wID == uid) return hMenu;

			if (menuiteminfo.hSubMenu)
			{
				const auto hSub = LocateSubmenu(menuiteminfo.hSubMenu, uid);
				if (hSub) return hSub;
			}
		}

		return nullptr;
	}

	// Returns true if this HMENU contains an item with this uid
	bool ContainsMenuItem(_In_ HMENU hMenu, const UINT uid)
	{
		const UINT nCount = GetMenuItemCount(hMenu);
		if (nCount == static_cast<UINT>(-1)) return false;

		for (UINT nPosition = 0; nPosition < nCount; nPosition++)
		{
			auto menuiteminfo = MENUITEMINFOW{sizeof(MENUITEMINFOW), MIIM_ID};
			WC_B_S(GetMenuItemInfoW(hMenu, nPosition, true, &menuiteminfo));

			if (menuiteminfo.wID == uid) return true;
		}

		return false;
	}

	bool DeleteMenu(_In_ HMENU hMenu, UINT uid)
	{
		const UINT nCount = GetMenuItemCount(hMenu);
		if (nCount == static_cast<UINT>(-1)) return false;

		for (UINT nPosition = 0; nPosition < nCount; nPosition++)
		{
			auto menuiteminfo = MENUITEMINFOW{sizeof(MENUITEMINFOW), MIIM_DATA | MIIM_ID | MIIM_SUBMENU};

			GetMenuItemInfoW(hMenu, nPosition, true, &menuiteminfo);
			if (menuiteminfo.wID == uid && menuiteminfo.dwItemData)
			{
				DeleteMenuEntry(reinterpret_cast<LPMENUENTRY>(menuiteminfo.dwItemData));
				WC_B_S(::RemoveMenu(hMenu, nPosition, MF_BYPOSITION));
				return true;
			}

			if (menuiteminfo.hSubMenu)
			{
				// This submenu may not contain the item, but it may contain an submenu that does, so check it.
				if (DeleteMenu(menuiteminfo.hSubMenu, uid)) return true;
			}
		}

		return false;
	}

	bool DeleteSubmenu(_In_ HMENU hMenu, UINT uid)
	{
		const UINT nCount = GetMenuItemCount(hMenu);
		if (nCount == static_cast<UINT>(-1)) return false;

		// Walk through child items looking for submenus
		for (UINT nPosition = 0; nPosition < nCount; nPosition++)
		{
			auto menuiteminfo = MENUITEMINFOW{sizeof(MENUITEMINFOW), MIIM_DATA | MIIM_SUBMENU};
			WC_B_S(GetMenuItemInfoW(hMenu, nPosition, true, &menuiteminfo));

			if (menuiteminfo.hSubMenu)
			{
				if (ContainsMenuItem(menuiteminfo.hSubMenu, uid))
				{
					// Delete children
					DeleteMenuEntries(menuiteminfo.hSubMenu);

					// Delete self
					if (menuiteminfo.dwItemData)
					{
						DeleteMenuEntry(reinterpret_cast<LPMENUENTRY>(menuiteminfo.dwItemData));
					}

					// Remove self from parent menu
					WC_B_S(::RemoveMenu(hMenu, nPosition, MF_BYPOSITION));
					return true;
				}

				// This submenu may not contain the item, but it may contain an submenu that does, so check it.
				if (DeleteSubmenu(menuiteminfo.hSubMenu, uid)) return true;
			}
		}

		return false;
	}

	_Check_return_ int GetEditHeight(_In_opt_ HWND hwndEdit)
	{
		// Calculate the new height for the edit control.
		const auto iHeight = GetTextHeight(hwndEdit) +
							 2 * GetSystemMetrics(SM_CYFIXEDFRAME) // Adjust for the edit border
							 + 2 * GetSystemMetrics(SM_CXEDGE); // Adjust for the edit border
		return iHeight;
	}

	_Check_return_ int GetTextHeight(_In_opt_ HWND hwndEdit)
	{
		auto tmFont = TEXTMETRIC{};

		// Get the DC for the edit control.
		const auto hdc = WC_D(HDC, GetDC(hwndEdit));
		if (hdc)
		{
			// Get the metrics for the Segoe font.
			const auto hOldFont = SelectObject(hdc, GetSegoeFont());
			WC_B_S(::GetTextMetrics(hdc, &tmFont));
			SelectObject(hdc, hOldFont);
			ReleaseDC(hwndEdit, hdc);
		}

		// Calculate the new height for the static control.
		return tmFont.tmHeight;
	}

	SIZE GetTextExtentPoint32(HDC hdc, const std::wstring& szText) noexcept
	{
		auto size = SIZE{};
		GetTextExtentPoint32W(hdc, szText.c_str(), static_cast<int>(szText.length()), &size);
		return size;
	}

	int CALLBACK EnumFontFamExProcW(
		_In_ LPLOGFONTW lplf,
		_In_ NEWTEXTMETRICEXW* /*lpntme*/,
		DWORD /*FontType*/,
		LPARAM lParam) noexcept
	{
		// Use a 9 point font
		lplf->lfHeight = -MulDiv(9, GetDeviceCaps(GetDC(nullptr), LOGPIXELSY), 72);
		lplf->lfWidth = 0;
		lplf->lfCharSet = DEFAULT_CHARSET;
		*reinterpret_cast<HFONT*>(lParam) = CreateFontIndirectW(lplf);
		return 0;
	}

	// This font is not cached and must be delete manually
	HFONT GetFont(_In_z_ const LPCWSTR szFont) noexcept
	{
		HFONT hFont = nullptr;
		LOGFONTW lfFont = {};
		wcscpy_s(lfFont.lfFaceName, _countof(lfFont.lfFaceName), szFont);

		EnumFontFamiliesExW(
			GetDC(nullptr),
			&lfFont,
			reinterpret_cast<FONTENUMPROCW>(EnumFontFamExProcW),
			reinterpret_cast<LPARAM>(&hFont),
			0);
		if (hFont) return hFont;

		// If we can't get our font, fallback to a system font
		auto ncm = NONCLIENTMETRICSW{};
		ncm.cbSize = sizeof ncm;
		SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, NULL, &ncm, NULL);
		static auto lf = ncm.lfMessageFont;

		// Create the font, and then return its handle.
		hFont = CreateFontW(
			lf.lfHeight,
			lf.lfWidth,
			lf.lfEscapement,
			lf.lfOrientation,
			lf.lfWeight,
			lf.lfItalic,
			lf.lfUnderline,
			lf.lfStrikeOut,
			lf.lfCharSet,
			lf.lfOutPrecision,
			lf.lfClipPrecision,
			lf.lfQuality,
			lf.lfPitchAndFamily,
			lf.lfFaceName);
		return hFont;
	}

	// Cached for deletion in UninitializeGDI
	HFONT GetSegoeFont() noexcept
	{
		if (g_hFontSegoe) return g_hFontSegoe;
		g_hFontSegoe = GetFont(SEGOEW);
		return g_hFontSegoe;
	}

	// Cached for deletion in UninitializeGDI
	HFONT GetSegoeFontBold() noexcept
	{
		if (g_hFontSegoeBold) return g_hFontSegoeBold;
		g_hFontSegoeBold = GetFont(SEGOEBOLD);
		return g_hFontSegoeBold;
	}

	_Check_return_ HBRUSH GetSysBrush(const uiColor uc) noexcept
	{
		// Return a cached brush if we have one
		if (g_SysBrushes[static_cast<int>(uc)]) return g_SysBrushes[static_cast<int>(uc)];
		// No cached brush found, cache and return a system brush if requested
		const auto iSysColor = g_SysColors[static_cast<int>(uc)];
		if (iSysColor)
		{
			g_SysBrushes[static_cast<int>(uc)] = GetSysColorBrush(iSysColor);
			return g_SysBrushes[static_cast<int>(uc)];
		}

		// No system brush for this color, cache and return a solid brush of the requested color
		const auto mc = static_cast<int>(g_FixedColors[static_cast<int>(uc)]);
		if (g_FixedBrushes[mc]) return g_FixedBrushes[mc];
		g_FixedBrushes[mc] = CreateSolidBrush(g_Colors[mc]);
		return g_FixedBrushes[mc];
	}

	_Check_return_ COLORREF MyGetSysColor(const uiColor uc) noexcept
	{
		// Return a system color if we have one in g_SysColors
		const auto iSysColor = g_SysColors[static_cast<int>(uc)];
		if (iSysColor) return GetSysColor(iSysColor);

		// No system color listed in g_SysColors, return a hard coded color
		const auto mc = static_cast<int>(g_FixedColors[static_cast<int>(uc)]);
		return g_Colors[mc];
	}

	_Check_return_ HPEN GetPen(const uiPen up) noexcept
	{
		if (g_Pens[static_cast<int>(up)]) return g_Pens[static_cast<int>(up)];
		auto lbr = LOGBRUSH{};
		lbr.lbStyle = BS_SOLID;

		switch (up)
		{
		case uiPen::SolidPen:
		{
			lbr.lbColor = MyGetSysColor(uiColor::FrameSelected);
			g_Pens[static_cast<int>(uiPen::SolidPen)] = ExtCreatePen(PS_COSMETIC | PS_SOLID, 1, &lbr, 0, nullptr);
			return g_Pens[static_cast<int>(uiPen::SolidPen)];
		}
		case uiPen::SolidGreyPen:
		{
			lbr.lbColor = MyGetSysColor(uiColor::FrameUnselected);
			g_Pens[static_cast<int>(uiPen::SolidGreyPen)] = ExtCreatePen(PS_COSMETIC | PS_SOLID, 1, &lbr, 0, nullptr);
			return g_Pens[static_cast<int>(uiPen::SolidGreyPen)];
		}
		case uiPen::DashedPen:
		{
			lbr.lbColor = MyGetSysColor(uiColor::FrameSelected);
			const DWORD rgStyle[2] = {1, 3};
			g_Pens[static_cast<int>(uiPen::DashedPen)] = ExtCreatePen(PS_GEOMETRIC | PS_USERSTYLE, 1, &lbr, 2, rgStyle);
			return g_Pens[static_cast<int>(uiPen::DashedPen)];
		}
		default:
			break;
		}
		return nullptr;
	}

	HBITMAP GetBitmap(const uiBitmap ub) noexcept
	{
		if (g_Bitmaps[static_cast<int>(ub)]) return g_Bitmaps[static_cast<int>(ub)];

		g_Bitmaps[static_cast<int>(ub)] =
			::LoadBitmap(GetModuleHandle(nullptr), MAKEINTRESOURCE(g_BitmapResources[static_cast<int>(ub)]));
		return g_Bitmaps[static_cast<int>(ub)];
	}

	SCALE GetDPIScale() noexcept
	{
		const auto hdcWin = GetWindowDC(nullptr);
		const auto dpiX = GetDeviceCaps(hdcWin, LOGPIXELSX);
		const auto dpiY = GetDeviceCaps(hdcWin, LOGPIXELSY);

		if (hdcWin) DeleteDC(hdcWin);
		return {dpiX, dpiY, 96};
	}

	HBITMAP ScaleBitmap(HBITMAP hBitmap, const SCALE& scale) noexcept
	{
		auto bm = BITMAP{};
		::GetObject(hBitmap, sizeof(BITMAP), &bm);

		const SIZE sizeSrc = {bm.bmWidth, bm.bmHeight};
		const SIZE sizeDst = {scale.x * sizeSrc.cx / scale.denominator, scale.y * sizeSrc.cy / scale.denominator};

		const auto hdcWin = GetWindowDC(nullptr);
		const auto hRet = CreateCompatibleBitmap(hdcWin, sizeDst.cx, sizeDst.cy);

		const auto hdcSrc = CreateCompatibleDC(hdcWin);
		const auto hdcDst = CreateCompatibleDC(hdcWin);

		const auto bmpSrc = SelectObject(hdcSrc, hBitmap);
		const auto bmpDst = SelectObject(hdcDst, hRet);

		static_cast<void>(
			StretchBlt(hdcDst, 0, 0, sizeDst.cx, sizeDst.cy, hdcSrc, 0, 0, sizeSrc.cx, sizeSrc.cy, SRCCOPY));

		static_cast<void>(SelectObject(hdcSrc, bmpSrc));
		static_cast<void>(SelectObject(hdcDst, bmpDst));

		if (bmpDst) DeleteObject(bmpDst);
		if (bmpSrc) DeleteObject(bmpSrc);
		if (hdcDst) DeleteDC(hdcDst);
		if (hdcSrc) DeleteDC(hdcSrc);
		if (hdcWin) DeleteDC(hdcWin);

		return hRet;
	}

	void DrawSegoeTextW(
		_In_ HDC hdc,
		_In_ const std::wstring& lpchText,
		_In_ const COLORREF color,
		_In_ const RECT& rc,
		const bool bBold,
		_In_ const UINT format)
	{
		output::DebugPrint(output::dbgLevel::Draw, L"Draw %d, \"%ws\"\n", rc.right - rc.left, lpchText.c_str());
		const auto hfontOld = SelectObject(hdc, bBold ? GetSegoeFontBold() : GetSegoeFont());
		const auto crText = SetTextColor(hdc, color);
		SetBkMode(hdc, TRANSPARENT);
		RECT drawRc = rc;
		DrawTextW(hdc, lpchText.c_str(), -1, &drawRc, format);

#ifdef SKIPBUFFER
		::FrameRect(hdc, &drawRc, GetSysBrush(bBold ? cBitmapTransFore : cBitmapTransBack));
#endif

		SelectObject(hdc, hfontOld);
		static_cast<void>(SetTextColor(hdc, crText));
	}

	// Clear/initialize formatting on the rich edit control.
	// We have to force load the system riched20 to ensure this doesn't break since
	// Office's riched20 apparently doesn't handle CFM_COLOR at all.
	// Sets our text color, script, and turns off bold, italic, etc formatting.
	void ClearEditFormatting(_In_ HWND hWnd, const bool bReadOnly) noexcept
	{
		CHARFORMAT2 cf{};
		ZeroMemory(&cf, sizeof cf);
		cf.cbSize = sizeof cf;
		cf.dwMask = CFM_COLOR | CFM_FACE | CFM_BOLD | CFM_ITALIC | CFM_UNDERLINE | CFM_STRIKEOUT;
		cf.crTextColor = MyGetSysColor(bReadOnly ? uiColor::TextReadOnly : uiColor::Text);
		_tcscpy_s(cf.szFaceName, _countof(cf.szFaceName), SEGOE);
		static_cast<void>(::SendMessage(hWnd, EM_SETCHARFORMAT, SCF_ALL, reinterpret_cast<LPARAM>(&cf)));
	}

	// Lighten the colors of the base, being careful not to overflow
	constexpr COLORREF LightColor(const COLORREF crBase)
	{
		constexpr auto f = .20;
		auto bRed = static_cast<BYTE>(GetRValue(crBase) + 255 * f);
		auto bGreen = static_cast<BYTE>(GetGValue(crBase) + 255 * f);
		auto bBlue = static_cast<BYTE>(GetBValue(crBase) + 255 * f);
		if (bRed < GetRValue(crBase)) bRed = 0xff;
		if (bGreen < GetGValue(crBase)) bGreen = 0xff;
		if (bBlue < GetBValue(crBase)) bBlue = 0xff;
		return RGB(bRed, bGreen, bBlue);
	}

	void GradientFillRect(_In_ HDC hdc, const RECT rc, const uiColor uc) noexcept
	{
		// Gradient fill the background
		const auto crGlow = MyGetSysColor(uc);
		const auto crLightGlow = LightColor(crGlow);

		TRIVERTEX vertex[2] = {};
		// Light at the top
		vertex[0].x = rc.left;
		vertex[0].y = rc.top;
		vertex[0].Red = GetRValue(crLightGlow) << 8;
		vertex[0].Green = GetGValue(crLightGlow) << 8;
		vertex[0].Blue = GetBValue(crLightGlow) << 8;
		vertex[0].Alpha = 0x0000;

		// Dark at the bottom
		vertex[1].x = rc.right;
		vertex[1].y = rc.bottom;
		vertex[1].Red = GetRValue(crGlow) << 8;
		vertex[1].Green = GetGValue(crGlow) << 8;
		vertex[1].Blue = GetBValue(crGlow) << 8;
		vertex[1].Alpha = 0x0000;

		// Create a GRADIENT_RECT structure that references the TRIVERTEX vertices.
		auto gRect = GRADIENT_RECT{};
		gRect.UpperLeft = 0;
		gRect.LowerRight = 1;
		GradientFill(hdc, vertex, 2, &gRect, 1, GRADIENT_FILL_RECT_V);
	}

	void DrawFilledPolygon(
		_In_ HDC hdc,
		_In_count_(cpt) const POINT* apt,
		_In_ const int cpt,
		const COLORREF cEdge,
		_In_ HBRUSH hFill) noexcept
	{
		const auto hPen = CreatePen(PS_SOLID, 0, cEdge);
		const auto hBrushOld = SelectObject(hdc, hFill);
		const auto hPenOld = SelectObject(hdc, hPen);
		Polygon(hdc, apt, cpt);
		SelectObject(hdc, hPenOld);
		SelectObject(hdc, hBrushOld);
		DeleteObject(hPen);
	}

	// Draw the frame of our edit controls
	LRESULT CALLBACK DrawEditProc(
		_In_ HWND hWnd,
		const UINT uMsg,
		const WPARAM wParam,
		const LPARAM lParam,
		const UINT_PTR uIdSubclass,
		DWORD_PTR /*dwRefData*/) noexcept
	{
		static auto borderWidth = 2;
		switch (uMsg)
		{
		case WM_NCCALCSIZE:
		{
			const auto ret = DefSubclassProc(hWnd, uMsg, wParam, lParam);
			InflateRect((LPRECT) lParam, -borderWidth, -borderWidth);
			return ret;
		}
		case WM_NCDESTROY:
			RemoveWindowSubclass(hWnd, DrawEditProc, uIdSubclass);
			return DefSubclassProc(hWnd, uMsg, wParam, lParam);
		case WM_SETFOCUS:
		case WM_KILLFOCUS:
			// Trigger WM_NCPAINT so we can get an updated frame reflecting focus status
			::RedrawWindow(hWnd, nullptr, nullptr, RDW_FRAME | RDW_INVALIDATE | RDW_ERASENOW);
			break;
		case WM_NCPAINT:
		{
			const auto hdc = GetWindowDC(hWnd);
			if (hdc)
			{
				auto rc = RECT{};
				GetWindowRect(hWnd, &rc);
				OffsetRect(&rc, -rc.left, -rc.top);
				if ((::GetFocus() == hWnd))
				{
					FrameRect(hdc, rc, borderWidth, uiColor::Glow);
				}
				else
				{
					// Clear out any previous border first, plus 1 for the built in border
					FrameRect(hdc, rc, borderWidth + 1, uiColor::Background);
					FrameRect(hdc, rc, 1, uiColor::FrameSelected);
				}

				ReleaseDC(hWnd, hdc);
			}

			// Let the system paint the scroll bar if we have one
			// Be sure to use window coordinates
			const auto ws = GetWindowLongPtr(hWnd, GWL_STYLE);
			if (ws & WS_VSCROLL)
			{
				auto rcScroll = RECT{};
				GetWindowRect(hWnd, &rcScroll);
				InflateRect(&rcScroll, -1, -1);
				rcScroll.left = rcScroll.right - GetSystemMetrics(SM_CXHSCROLL);
				auto hRgnCaption = CreateRectRgnIndirect(&rcScroll);
				::DefWindowProc(hWnd, WM_NCPAINT, reinterpret_cast<WPARAM>(hRgnCaption), NULL);
				DeleteObject(hRgnCaption);
			}

			return 0;
		}
		}

		return DefSubclassProc(hWnd, uMsg, wParam, lParam);
	}

	void SubclassEdit(_In_ HWND hWnd, _In_opt_ HWND hWndParent, const bool bReadOnly)
	{
		SetWindowSubclass(hWnd, DrawEditProc, 0, reinterpret_cast<DWORD_PTR>(hWndParent));

		auto lStyle = ::GetWindowLongPtr(hWnd, GWL_EXSTYLE);
		lStyle &= ~WS_EX_CLIENTEDGE;
		static_cast<void>(::SetWindowLongPtr(hWnd, GWL_EXSTYLE, lStyle));
		if (bReadOnly)
		{
			static_cast<void>(::SendMessage(
				hWnd,
				EM_SETBKGNDCOLOR,
				static_cast<WPARAM>(0),
				static_cast<LPARAM>(MyGetSysColor(uiColor::BackgroundReadOnly))));
			static_cast<void>(::SendMessage(hWnd, EM_SETREADONLY, true, 0L));
		}
		else
		{
			static_cast<void>(::SendMessage(
				hWnd,
				EM_SETBKGNDCOLOR,
				static_cast<WPARAM>(0),
				static_cast<LPARAM>(MyGetSysColor(uiColor::Background))));
		}

		ClearEditFormatting(hWnd, bReadOnly);

		// Set up callback to control paste and context menus
		auto reCallback = new (std::nothrow) CRichEditOleCallback(hWnd, hWndParent);
		if (reCallback)
		{
			static_cast<void>(
				::SendMessage(hWnd, EM_SETOLECALLBACK, static_cast<WPARAM>(0), reinterpret_cast<LPARAM>(reCallback)));
			reCallback->Release();
		}
	}

	void CustomDrawList(_In_ LPNMLVCUSTOMDRAW lvcd, _In_ LRESULT* pResult, const DWORD_PTR iItemCurHover) noexcept
	{
		static auto bSelected = false;
		if (!lvcd) return;
		const auto iItem = lvcd->nmcd.dwItemSpec;

		// If there's nothing to paint, this is a "fake paint" and we don't want to toggle selection highlight
		// Toggling selection highlight causes a repaint, so this logic prevents flicker
		const auto bFakePaint =
			lvcd->nmcd.rc.bottom == 0 && lvcd->nmcd.rc.top == 0 && lvcd->nmcd.rc.left == 0 && lvcd->nmcd.rc.right == 0;

		switch (lvcd->nmcd.dwDrawStage)
		{
		case CDDS_PREPAINT:
			*pResult = CDRF_NOTIFYITEMDRAW | CDRF_NOTIFYPOSTPAINT;
			break;

		case CDDS_ITEMPREPAINT:
			// Let the item draw itself without a selection highlight
			bSelected = ListView_GetItemState(lvcd->nmcd.hdr.hwndFrom, iItem, LVIS_SELECTED) != 0;
			if (bSelected && !bFakePaint)
			{
				// Turn off listview selection highlight
				ListView_SetItemState(lvcd->nmcd.hdr.hwndFrom, iItem, 0, LVIS_SELECTED)
			}

			*pResult = CDRF_DODEFAULT | CDRF_NOTIFYPOSTPAINT | CDRF_NOTIFYSUBITEMDRAW;
			break;

		case CDDS_SUBITEM | CDDS_ITEMPREPAINT:
		{
			// Big speedup - we skip painting if the target rectangle is outside the client area

			// On some systems, top/bottom will both be zero, but not left/right. We want to consider those
			// as if they had positive area, so bump bottom to 1 and remember to reset it after.
			auto bResetBottom = false;
			if (0 == lvcd->nmcd.rc.top && 0 == lvcd->nmcd.rc.bottom)
			{
				lvcd->nmcd.rc.bottom = 1;
				bResetBottom = true;
			}

			auto rc = RECT{};
			GetClientRect(lvcd->nmcd.hdr.hwndFrom, &rc);
			IntersectRect(&rc, &rc, &lvcd->nmcd.rc);

			if (bResetBottom)
			{
				lvcd->nmcd.rc.bottom = 0;
			}

			if (IsRectEmpty(&rc))
			{
				*pResult = CDRF_SKIPDEFAULT;
				break;
			}

			// Turn on listview hover highlight
			if (bSelected)
			{
				lvcd->clrText = MyGetSysColor(uiColor::GlowText);
				lvcd->clrTextBk = MyGetSysColor(uiColor::SelectedBackground);
			}
			else if (iItemCurHover == iItem)
			{
				lvcd->clrText = MyGetSysColor(uiColor::GlowText);
				lvcd->clrTextBk = MyGetSysColor(uiColor::GlowBackground);
			}

			*pResult = CDRF_NEWFONT;
			break;
		}

		case CDDS_ITEMPOSTPAINT:
			// And then we'll handle our frame
			if (bSelected && !bFakePaint)
			{
				// Turn on listview selection highlight
				ListView_SetItemState(lvcd->nmcd.hdr.hwndFrom, iItem, LVIS_SELECTED, LVIS_SELECTED);
			}

			break;

		default:
			*pResult = CDRF_DODEFAULT;
			break;
		}
	}

	// Handle highlight glow for list items
	void DrawListItemGlow(_In_ HWND hWnd, const UINT itemID) noexcept
	{
		if (itemID == static_cast<UINT>(-1)) return;

		// If the item already has the selection glow, we don't need to redraw
		const auto bSelected = ListView_GetItemState(hWnd, itemID, LVIS_SELECTED) != 0;
		if (bSelected) return;

		auto rcIcon = RECT{};
		auto rcLabels = RECT{};
		ListView_GetItemRect(hWnd, itemID, &rcLabels, LVIR_BOUNDS);
		ListView_GetItemRect(hWnd, itemID, &rcIcon, LVIR_ICON);
		rcLabels.left = rcIcon.right;
		if (rcLabels.left >= rcLabels.right) return;
		auto rcClient = RECT{};
		GetClientRect(hWnd, &rcClient); // Get our client size
		IntersectRect(&rcLabels, &rcLabels, &rcClient);
		RedrawWindow(hWnd, &rcLabels, nullptr, RDW_INVALIDATE | RDW_UPDATENOW);
	}

	void DrawTreeItemGlow(_In_ HWND hWnd, _In_ HTREEITEM hItem) noexcept
	{
		auto rect = RECT{};
		auto rectTree = RECT{};
		TreeView_GetItemRect(hWnd, hItem, &rect, false);
		GetClientRect(hWnd, &rectTree);
		rect.left = rectTree.left;
		rect.right = rectTree.right;
		InvalidateRect(hWnd, &rect, false);
	}

	// Copies iWidth x iHeight rect from hdcSource to hdcTarget, replacing colors
	// No scaling is performed
	void CopyBitmap(
		HDC hdcSource,
		HDC hdcTarget,
		const int iWidth,
		const int iHeight,
		const uiColor cSource,
		const uiColor cReplace) noexcept
	{
		RECT rcBM = {0, 0, iWidth, iHeight};

		const auto hbmTarget = CreateCompatibleBitmap(hdcSource, iWidth, iHeight);
		static_cast<void>(SelectObject(hdcTarget, hbmTarget));
		FillRect(hdcTarget, &rcBM, GetSysBrush(cReplace));

		static_cast<void>(
			TransparentBlt(hdcTarget, 0, 0, iWidth, iHeight, hdcSource, 0, 0, iWidth, iHeight, MyGetSysColor(cSource)));
		if (hbmTarget) DeleteObject(hbmTarget);
	}

	// Copies iWidth x iHeight rect from hdcSource to hdcTarget, shifted diagnoally by offset
	// Background filled with cBackground
	void ShiftBitmap(
		HDC hdcSource,
		HDC hdcTarget,
		const int iWidth,
		const int iHeight,
		const int offset,
		const uiColor cBackground) noexcept
	{
		RECT rcBM = {0, 0, iWidth, iHeight};

		const auto hbmTarget = CreateCompatibleBitmap(hdcSource, iWidth, iHeight);
		static_cast<void>(SelectObject(hdcTarget, hbmTarget));
		FillRect(hdcTarget, &rcBM, GetSysBrush(cBackground));

		static_cast<void>(BitBlt(hdcTarget, offset, offset, iWidth, iHeight, hdcSource, 0, 0, SRCCOPY));
		if (hbmTarget) DeleteObject(hbmTarget);
	}

	// Draws a bitmap on the screen, double buffered, with two color replacement
	// Fills rectangle with cBackground
	// Replaces cBitmapTransFore (cyan) with cFrameSelected
	// Replaces cBitmapTransBack (magenta) with the cBackground
	// Scales from size of bitmap to size of target rectangle
	void DrawBitmap(
		_In_ HDC hdc,
		_In_ const RECT& rcTarget,
		const uiBitmap iBitmap,
		const bool bHover,
		const int offset) noexcept
	{
		const int iWidth = rcTarget.right - rcTarget.left;
		const int iHeight = rcTarget.bottom - rcTarget.top;

		// hdcBitmap: Load the image
		const auto hdcBitmap = CreateCompatibleDC(hdc);
		// TODO: pass target dimensions here and load the most appropriate bitmap
		const auto hbmBitmap = GetBitmap(iBitmap);
		static_cast<void>(SelectObject(hdcBitmap, hbmBitmap));

		auto bm = BITMAP{};
		::GetObject(hbmBitmap, sizeof bm, &bm);

		// hdcForeReplace: Create a bitmap compatible with hdc, select it, fill with cFrameSelected, copy from hdcBitmap, with cBitmapTransFore transparent
		const auto hdcForeReplace = CreateCompatibleDC(hdc);
		CopyBitmap(
			hdcBitmap, hdcForeReplace, bm.bmWidth, bm.bmHeight, uiColor::BitmapTransFore, uiColor::FrameSelected);

		// hdcBackReplace: Create a bitmap compatible with hdc, select it, fill with cBackground, copy from hdcForeReplace, with cBitmapTransBack transparent
		const auto hdcBackReplace = CreateCompatibleDC(hdc);
		CopyBitmap(
			hdcForeReplace,
			hdcBackReplace,
			bm.bmWidth,
			bm.bmHeight,
			uiColor::BitmapTransBack,
			bHover ? uiColor::GlowBackground : uiColor::Background);

		const auto hdcShift = CreateCompatibleDC(hdc);
		ShiftBitmap(
			hdcBackReplace,
			hdcShift,
			bm.bmWidth,
			bm.bmHeight,
			offset,
			bHover ? uiColor::GlowBackground : uiColor::Background);

		// In case the original bitmap dimensions doesn't match our target dimension, we stretch it to fit
		// We can get better results if the original bitmap happens to match.
		static_cast<void>(StretchBlt(
			hdc, rcTarget.left, rcTarget.top, iWidth, iHeight, hdcShift, 0, 0, bm.bmWidth, bm.bmHeight, SRCCOPY));

		if (hdcShift) DeleteDC(hdcShift);
		if (hdcBackReplace) DeleteDC(hdcBackReplace);
		if (hdcForeReplace) DeleteDC(hdcForeReplace);
		if (hdcBitmap) DeleteDC(hdcBitmap);

#ifdef SKIPBUFFER
		::FrameRect(hdc, &rcTarget, GetSysBrush(bHover ? cBitmapTransFore : cBitmapTransBack));
#endif
	}

	void
	CustomDrawTree(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult, const bool bHover, _In_ HTREEITEM hItemCurHover) noexcept
	{
		if (!pNMHDR) return;
		const auto lvcd = reinterpret_cast<LPNMTVCUSTOMDRAW>(pNMHDR);

		switch (lvcd->nmcd.dwDrawStage)
		{
		case CDDS_PREPAINT:
			*pResult = CDRF_NOTIFYITEMDRAW;
			break;

		case CDDS_ITEMPREPAINT:
		{
			auto hItem = reinterpret_cast<HTREEITEM>(lvcd->nmcd.dwItemSpec);

			if (hItem)
			{
				lvcd->clrTextBk = MyGetSysColor(uiColor::Background);
				lvcd->clrText = MyGetSysColor(uiColor::Text);
				const int iState = TreeView_GetItemState(lvcd->nmcd.hdr.hwndFrom, hItem, TVIS_SELECTED);
				TreeView_SetItemState(
					lvcd->nmcd.hdr.hwndFrom, hItem, iState & TVIS_SELECTED ? TVIS_BOLD : NULL, TVIS_BOLD);
			}

			*pResult = CDRF_DODEFAULT | CDRF_NOTIFYPOSTPAINT;

			if (hItem == hItemCurHover)
			{
				lvcd->clrText = MyGetSysColor(uiColor::GlowText);
				lvcd->clrTextBk = MyGetSysColor(uiColor::GlowBackground);
				*pResult |= CDRF_NEWFONT;
			}

			break;
		}

		case CDDS_ITEMPOSTPAINT:
		{
			const auto hItem = reinterpret_cast<HTREEITEM>(lvcd->nmcd.dwItemSpec);

			if (hItem)
			{
				// Cover over the +/- and paint triangles instead
				DrawExpandTriangle(
					lvcd->nmcd.hdr.hwndFrom,
					lvcd->nmcd.hdc,
					hItem,
					bHover && hItem == hItemCurHover,
					hItem == hItemCurHover);
			}
			break;
		}

		default:
			*pResult = CDRF_DODEFAULT;
			break;
		}
	}

	// Paints the triangles indicating expansion state
	void
	DrawExpandTriangle(_In_ HWND hWnd, _In_ HDC hdc, _In_ HTREEITEM hItem, const bool bGlow, const bool bHover) noexcept
	{
		auto tvitem = TVITEM{};
		tvitem.hItem = hItem;
		tvitem.mask = TVIF_CHILDREN | TVIF_STATE;
		TreeView_GetItem(hWnd, &tvitem);
		const auto bHasChildren = tvitem.cChildren != 0;
		if (bHasChildren)
		{
			auto rcButton = RECT{};
			TreeView_GetItemRect(hWnd, hItem, &rcButton, true);
			const auto triangleSize = (rcButton.bottom - rcButton.top) / 4;

			// Erase the +/- icons
			// We erase everything to the left of the label
			rcButton.right = rcButton.left;
			rcButton.left = 0;
			FillRect(hdc, &rcButton, GetSysBrush(bHover ? uiColor::GlowBackground : uiColor::Background));

			// Now we focus on a box 15 pixels wide to the left of the label
			rcButton.left = rcButton.right - 15;

			// Boundary box for the actual triangles
			auto rcTriangle = RECT{};

			POINT tri[3] = {};
			const auto bExpanded = TVIS_EXPANDED == (tvitem.state & TVIS_EXPANDED);
			if (bExpanded)
			{
				rcTriangle.top = (rcButton.top + rcButton.bottom) / 2 - (triangleSize - 1);
				rcTriangle.bottom = rcTriangle.top + (triangleSize + 1);
				rcTriangle.left = rcButton.left;
				rcTriangle.right = rcTriangle.left + (triangleSize + 1);

				tri[0].x = rcTriangle.left;
				tri[0].y = rcTriangle.bottom;
				tri[1].x = rcTriangle.right;
				tri[1].y = rcTriangle.top;
				tri[2].x = rcTriangle.right;
				tri[2].y = rcTriangle.bottom;
			}
			else
			{
				rcTriangle.top = (rcButton.top + rcButton.bottom) / 2 - triangleSize;
				rcTriangle.bottom = rcTriangle.top + triangleSize * 2;
				rcTriangle.left = rcButton.left;
				rcTriangle.right = rcTriangle.left + triangleSize;

				tri[0].x = rcTriangle.left;
				tri[0].y = rcTriangle.top;
				tri[1].x = rcTriangle.left;
				tri[1].y = rcTriangle.bottom;
				tri[2].x = rcTriangle.right;
				tri[2].y = (rcTriangle.top + rcTriangle.bottom) / 2;
			}

			auto cEdge = bGlow ? uiColor::Glow : bExpanded ? uiColor::FrameSelected : uiColor::Arrow;
			auto cFill = bGlow ? uiColor::Glow : bExpanded ? uiColor::FrameSelected : uiColor::Background;

			DrawFilledPolygon(hdc, tri, _countof(tri), MyGetSysColor(cEdge), GetSysBrush(cFill));
		}
	}

	void DrawTriangle(_In_ HWND hWnd, _In_ HDC hdc, _In_ const RECT& rc, const bool bButton, const bool bUp)
	{
		if (!hWnd || !hdc) return;

		CDoubleBuffer db;
		db.Begin(hdc, rc);

		FillRect(hdc, &rc, GetSysBrush(bButton ? uiColor::PaneHeaderBackground : uiColor::Background));

		POINT tri[3] = {};
		auto lCenter = LONG{};
		auto lTop = LONG{};
		const auto triangleSize = (rc.bottom - rc.top) / 5;
		if (bButton)
		{
			lCenter = rc.left + (rc.bottom - rc.top) / 2;
			lTop = (rc.top + rc.bottom - triangleSize) / 2;
		}
		else
		{
			lCenter = (rc.left + rc.right) / 2;
			lTop = rc.top + GetSystemMetrics(SM_CYBORDER);
		}

		if (bUp)
		{
			tri[0].x = lCenter;
			tri[0].y = lTop;
			tri[1].x = lCenter - triangleSize;
			tri[1].y = lTop + triangleSize;
			tri[2].x = lCenter + triangleSize;
			tri[2].y = lTop + triangleSize;
		}
		else
		{
			tri[0].x = lCenter;
			tri[0].y = lTop + triangleSize;
			tri[1].x = lCenter - triangleSize;
			tri[1].y = lTop;
			tri[2].x = lCenter + triangleSize;
			tri[2].y = lTop;
		}

		auto uiArrow = uiColor::Arrow;
		if (bButton) uiArrow = uiColor::PaneHeaderText;
		DrawFilledPolygon(hdc, tri, _countof(tri), MyGetSysColor(uiArrow), GetSysBrush(uiArrow));

		db.End(hdc);
	}

	void DrawHeaderItem(_In_ HWND hWnd, _In_ HDC hdc, const UINT itemID, _In_ const RECT& rc)
	{
		auto bSorted = false;
		auto bSortUp = true;

		WCHAR szHeader[255] = {0};
		auto hdItem = HDITEMW{};
		hdItem.mask = HDI_TEXT | HDI_FORMAT;
		hdItem.pszText = szHeader;
		hdItem.cchTextMax = _countof(szHeader);
		::SendMessage(hWnd, HDM_GETITEMW, static_cast<WPARAM>(itemID), reinterpret_cast<LPARAM>(&hdItem));

		if (hdItem.fmt & (HDF_SORTUP | HDF_SORTDOWN)) bSorted = true;
		if (bSorted)
		{
			bSortUp = (hdItem.fmt & HDF_SORTUP) != 0;
		}

		CDoubleBuffer db;
		db.Begin(hdc, rc);

		RECT rcHeader = rc;
		FillRect(hdc, &rcHeader, GetSysBrush(uiColor::Background));

		if (bSorted)
		{
			DrawTriangle(hWnd, hdc, rcHeader, false, bSortUp);
		}

		auto rcText = rcHeader;
		rcText.left += GetSystemMetrics(SM_CXEDGE);
		DrawSegoeTextW(
			hdc,
			hdItem.pszText,
			MyGetSysColor(uiColor::Text),
			rcText,
			false,
			DT_END_ELLIPSIS | DT_SINGLELINE | DT_VCENTER);

		// Draw a line under for some visual separation
		const auto hpenOld = SelectObject(hdc, GetPen(uiPen::SolidGreyPen));
		MoveToEx(hdc, rcHeader.left, rcHeader.bottom - 1, nullptr);
		LineTo(hdc, rcHeader.right, rcHeader.bottom - 1);
		static_cast<void>(SelectObject(hdc, hpenOld));

		// Draw our divider
		// Since no one else uses rcHeader after here, we can modify it in place
		InflateRect(&rcHeader, 0, -1);
		rcHeader.left = rcHeader.right - 2;
		rcHeader.bottom -= 1;
		::FrameRect(hdc, &rcHeader, GetSysBrush(uiColor::FrameUnselected));

		db.End(hdc);
	}

	// Draw the unused portion of the header
	void CustomDrawHeader(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
	{
		if (!pNMHDR) return;
		const auto lvcd = reinterpret_cast<LPNMCUSTOMDRAW>(pNMHDR);

		switch (lvcd->dwDrawStage)
		{
		case CDDS_PREPAINT:
			*pResult = CDRF_NOTIFYPOSTPAINT | CDRF_NOTIFYITEMDRAW;
			break;

		case CDDS_ITEMPREPAINT:
			DrawHeaderItem(lvcd->hdr.hwndFrom, lvcd->hdc, static_cast<UINT>(lvcd->dwItemSpec), lvcd->rc);
			*pResult = CDRF_SKIPDEFAULT;
			break;

		case CDDS_POSTPAINT:
		{
			auto rc = RECT{};
			// Get the rc for the entire header
			GetClientRect(lvcd->hdr.hwndFrom, &rc);
			// Get the last item in the header
			const auto iCount = Header_GetItemCount(lvcd->hdr.hwndFrom);
			if (iCount)
			{
				// Find the index of the last item in the header
				const auto iIndex = Header_OrderToIndex(lvcd->hdr.hwndFrom, (UINT_PTR) iCount - 1);

				auto rcRight = RECT{};
				// Compute the right edge of the last item
				static_cast<void>(Header_GetItemRect(lvcd->hdr.hwndFrom, iIndex, &rcRight));
				rc.left = rcRight.right;
			}

			// If we have visible non occupied header, paint it
			if (rc.left < rc.right)
			{
				FillRect(lvcd->hdc, &rc, GetSysBrush(uiColor::Background));
			}

			*pResult = CDRF_DODEFAULT;
			break;
		}

		default:
			*pResult = CDRF_DODEFAULT;
			break;
		}
	}

	void DrawTrackingBar(
		_In_ HWND hWndHeader,
		_In_ HWND hWndList,
		const int x,
		const int iHeaderHeight,
		const bool bRedraw) noexcept
	{
		auto rcTracker = RECT{};
		GetClientRect(hWndList, &rcTracker);
		const auto hdc = GetDC(hWndList);
		rcTracker.top += iHeaderHeight;
		rcTracker.left = x - 1; // this lines us up under the splitter line we drew in the header
		rcTracker.right = x;
		if (::MapWindowPoints(hWndHeader, hWndList, reinterpret_cast<LPPOINT>(&rcTracker), 2) != 0)
		{
			if (bRedraw)
			{
				InvalidateRect(hWndList, &rcTracker, true);
			}
			else
			{
				FillRect(hdc, &rcTracker, GetSysBrush(uiColor::FrameSelected));
			}
		}

		ReleaseDC(hWndList, hdc);
	}

	void StyleButton(_In_ HWND hWnd, uiButtonStyle bsStyle)
	{
		EC_B_S(::SetProp(hWnd, BUTTON_STYLE, reinterpret_cast<HANDLE>(bsStyle)));
	}

	void DrawButton(_In_ HWND hWnd, _In_ HDC hDC, _In_ const RECT& rc, const UINT itemState)
	{
		const auto background = registry::uiDiag ? GetSysBrush(uiColor::TestOrange) : GetSysBrush(uiColor::Background);
		FillRect(hDC, &rc, background);

		const auto iState = static_cast<int>(::SendMessage(hWnd, BM_GETSTATE, NULL, NULL));
		const auto bGlow = (iState & BST_HOT) != 0;
		const auto bPushed = (iState & BST_PUSHED) != 0;
		const auto bFocused = (itemState & CDIS_FOCUS) != 0;

		const auto bsStyle = static_cast<uiButtonStyle>(reinterpret_cast<intptr_t>(::GetProp(hWnd, BUTTON_STYLE)));
		switch (bsStyle)
		{
		case uiButtonStyle::Unstyled:
		{
			WCHAR szButton[255] = {0};
			GetWindowTextW(hWnd, szButton, _countof(szButton));
			const auto bDisabled = (itemState & CDIS_DISABLED) != 0;

			DrawSegoeTextW(
				hDC,
				szButton,
				bPushed || bDisabled ? MyGetSysColor(uiColor::TextDisabled) : MyGetSysColor(uiColor::Text),
				rc,
				false,
				DT_SINGLELINE | DT_VCENTER | DT_CENTER);
		}
		break;
		case uiButtonStyle::UpArrow:
			DrawTriangle(hWnd, hDC, rc, true, true);
			break;
		case uiButtonStyle::DownArrow:
			DrawTriangle(hWnd, hDC, rc, true, false);
			break;
		}

		::FrameRect(
			hDC,
			&rc,
			bFocused || bGlow || bPushed ? GetSysBrush(uiColor::FrameSelected) : GetSysBrush(uiColor::FrameUnselected));
	}

	bool CustomDrawButton(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
	{
		if (!pNMHDR) return false;
		const auto lvcd = reinterpret_cast<LPNMCUSTOMDRAW>(pNMHDR);

		// Ensure we only handle painting for buttons
		WCHAR szClass[64] = {0};
		GetClassNameW(lvcd->hdr.hwndFrom, szClass, _countof(szClass));
		if (_wcsicmp(szClass, L"BUTTON") != 0) return false;

		switch (lvcd->dwDrawStage)
		{
		case CDDS_PREPAINT:
		{
			const auto lStyle = BS_TYPEMASK & GetWindowLongA(lvcd->hdr.hwndFrom, GWL_STYLE);

			*pResult = CDRF_SKIPDEFAULT;
			if (BS_AUTOCHECKBOX == lStyle)
			{
				viewpane::CheckPane::Draw(lvcd->hdr.hwndFrom, lvcd->hdc, lvcd->rc, lvcd->uItemState);
			}
			else if (BS_PUSHBUTTON == lStyle || BS_DEFPUSHBUTTON == lStyle)
			{
				DrawButton(lvcd->hdr.hwndFrom, lvcd->hdc, lvcd->rc, lvcd->uItemState);
			}
			else
				*pResult = CDRF_DODEFAULT;
		}
		break;

		default:
			*pResult = CDRF_DODEFAULT;
			break;
		}
		return true;
	}

	void MeasureMenu(_In_ LPMEASUREITEMSTRUCT lpMeasureItemStruct)
	{
		if (!lpMeasureItemStruct) return;
		const auto lpMenuEntry = reinterpret_cast<LPMENUENTRY>(lpMeasureItemStruct->itemData);
		if (!lpMenuEntry) return;

		const auto hdc = GetDC(nullptr);
		if (hdc)
		{
			const auto hfontOld = SelectObject(hdc, GetSegoeFont());

			// In order to compute the right width, we need to drop our prefix characters
			auto szText = strings::StripCharacter(lpMenuEntry->m_MSAA.pszWText, L'&');

			const auto size = GetTextExtentPoint32(hdc, szText);
			lpMeasureItemStruct->itemWidth = size.cx + 4 * GetSystemMetrics(SM_CXEDGE);
			lpMeasureItemStruct->itemHeight = size.cy + 2 * GetSystemMetrics(SM_CYEDGE);

			// Make sure we have room for the flyout icon
			if (!lpMenuEntry->m_bOnMenuBar &&
				IsMenu(reinterpret_cast<HMENU>(static_cast<UINT_PTR>(lpMeasureItemStruct->itemID))))
			{
				lpMeasureItemStruct->itemWidth += GetSystemMetrics(SM_CXSMICON);
			}

			SelectObject(hdc, hfontOld);
			ReleaseDC(nullptr, hdc);

			output::DebugPrint(
				output::dbgLevel::Draw, L"Measure %d, \"%ws\"\n", lpMeasureItemStruct->itemWidth, szText.c_str());
		}
	}

	// Returns true if handled, false otherwise
	bool MeasureItem(_In_ LPMEASUREITEMSTRUCT lpMeasureItemStruct)
	{
		if (ODT_MENU == lpMeasureItemStruct->CtlType)
		{
			MeasureMenu(lpMeasureItemStruct);
			return true;
		}
		else if (ODT_COMBOBOX == lpMeasureItemStruct->CtlType)
		{
			lpMeasureItemStruct->itemHeight = GetEditHeight(nullptr);
			return true;
		}

		return false;
	}

	void DrawMenu(_In_ LPDRAWITEMSTRUCT lpDrawItemStruct)
	{
		if (!lpDrawItemStruct) return;
		const auto lpMenuEntry = reinterpret_cast<LPMENUENTRY>(lpDrawItemStruct->itemData);
		if (!lpMenuEntry) return;

		const auto bSeparator =
			!lpMenuEntry->m_MSAA.cchWText; // Cheap way to find separators. Revisit if odd effects show.
		const auto bAccel = (lpDrawItemStruct->itemState & ODS_NOACCEL) != 0;
		const auto bHot = (lpDrawItemStruct->itemState & (ODS_HOTLIGHT | ODS_SELECTED)) != 0;
		const auto bDisabled = (lpDrawItemStruct->itemState & (ODS_GRAYED | ODS_DISABLED)) != 0;

		output::DebugPrint(
			output::dbgLevel::Draw,
			L"DrawMenu %d, \"%ws\"\n",
			lpDrawItemStruct->rcItem.right - lpDrawItemStruct->rcItem.left,
			lpMenuEntry->m_pName.c_str());

		// Double buffer our menu painting
		CDoubleBuffer db;
		auto hdc = lpDrawItemStruct->hDC;
		db.Begin(hdc, lpDrawItemStruct->rcItem);

		// Draw background
		const auto rcItem = lpDrawItemStruct->rcItem;
		auto rcText = rcItem;

		if (!lpMenuEntry->m_bOnMenuBar)
		{
			auto rectGutter = rcText;
			rcText.left += GetSystemMetrics(SM_CXMENUCHECK);
			rectGutter.right = rcText.left;
			FillRect(hdc, &rectGutter, GetSysBrush(uiColor::BackgroundReadOnly));
		}

		auto cBack = uiColor::Background;
		auto cFore = uiColor::Text;

		if (bHot)
		{
			cBack = uiColor::SelectedBackground;
			cFore = uiColor::GlowText;
			FillRect(hdc, &rcItem, GetSysBrush(cBack));
		}

		if (!bHot || bDisabled)
		{
			if (bDisabled) cFore = uiColor::TextDisabled;
			FillRect(hdc, &rcText, GetSysBrush(cBack));
		}

		if (bSeparator)
		{
			InflateRect(&rcText, -3, 0);
			const auto lMid = (rcText.bottom + rcText.top) / 2;
			const auto hpenOld = SelectObject(hdc, GetPen(uiPen::SolidGreyPen));
			MoveToEx(hdc, rcText.left, lMid, nullptr);
			LineTo(hdc, rcText.right, lMid);
			static_cast<void>(SelectObject(hdc, hpenOld));
		}
		else if (!lpMenuEntry->m_pName.empty())
		{
			UINT uiTextFlags = DT_SINGLELINE | DT_VCENTER;
			if (bAccel) uiTextFlags |= DT_HIDEPREFIX;

			if (lpMenuEntry->m_bOnMenuBar)
				uiTextFlags |= DT_CENTER;
			else
				rcText.left += GetSystemMetrics(SM_CXEDGE);

			output::DebugPrint(
				output::dbgLevel::Draw,
				L"DrawMenu text %d, \"%ws\"\n",
				rcText.right - rcText.left,
				lpMenuEntry->m_pName.c_str());

			DrawSegoeTextW(hdc, lpMenuEntry->m_pName, MyGetSysColor(cFore), rcText, false, uiTextFlags);

			// Triple buffer the check mark so we can copy it over without the background
			if (lpDrawItemStruct->itemState & ODS_CHECKED)
			{
				const UINT nWidth = GetSystemMetrics(SM_CXMENUCHECK);
				const UINT nHeight = GetSystemMetrics(SM_CYMENUCHECK);
				auto rc = RECT{};
				const auto bm = CreateBitmap(nWidth, nHeight, 1, 1, nullptr);
				const auto hdcMem = CreateCompatibleDC(hdc);

				SelectObject(hdcMem, bm);
				SetRect(&rc, 0, 0, nWidth, nHeight);
				static_cast<void>(DrawFrameControl(hdcMem, &rc, DFC_MENU, DFCS_MENUCHECK));

				static_cast<void>(TransparentBlt(
					hdc,
					rcItem.left,
					(rcItem.top + rcItem.bottom - nHeight) / 2,
					nWidth,
					nHeight,
					hdcMem,
					0,
					0,
					nWidth,
					nHeight,
					MyGetSysColor(uiColor::Background)));

#ifdef SKIPBUFFER
				auto frameRect = rcItem;
				frameRect.right = frameRect.left + nWidth;
				frameRect.bottom = frameRect.top + nHeight;
				::FrameRect(hdc, &frameRect, GetSysBrush(cBitmapTransFore));
#endif

				DeleteDC(hdcMem);
				DeleteObject(bm);
			}
		}

		db.End(hdc);
	}

	std::wstring GetLBText(HWND hwnd, const int nIndex)
	{
		const auto len = ComboBox_GetLBTextLen(hwnd, nIndex) + 1;
		auto buffer = std::basic_string<TCHAR>(len, '\0');
		ComboBox_GetLBText(hwnd, nIndex, const_cast<TCHAR*>(buffer.c_str()));
		auto szOut = strings::LPCTSTRToWstring(buffer.c_str());
		return szOut;
	}

	void DrawComboBox(_In_ LPDRAWITEMSTRUCT lpDrawItemStruct)
	{
		if (!lpDrawItemStruct) return;
		if (lpDrawItemStruct->itemID == static_cast<UINT>(-1)) return;

		// Get and display the text for the list item.
		const auto szText = GetLBText(lpDrawItemStruct->hwndItem, lpDrawItemStruct->itemID);
		const auto bHot = 0 != (lpDrawItemStruct->itemState & (ODS_FOCUS | ODS_SELECTED));
		auto cBack = uiColor::Background;
		auto cFore = uiColor::Text;
		if (bHot)
		{
			cBack = uiColor::GlowBackground;
			cFore = uiColor::GlowText;
		}

		FillRect(lpDrawItemStruct->hDC, &lpDrawItemStruct->rcItem, GetSysBrush(cBack));

		DrawSegoeTextW(
			lpDrawItemStruct->hDC,
			szText,
			MyGetSysColor(cFore),
			lpDrawItemStruct->rcItem,
			false,
			DT_LEFT | DT_SINGLELINE | DT_VCENTER);
	}

	// Returns true if handled, false otherwise
	bool DrawItem(_In_ LPDRAWITEMSTRUCT lpDrawItemStruct)
	{
		if (!lpDrawItemStruct) return false;

		if (ODT_MENU == lpDrawItemStruct->CtlType)
		{
			DrawMenu(lpDrawItemStruct);
			return true;
		}
		else if (ODT_COMBOBOX == lpDrawItemStruct->CtlType)
		{
			DrawComboBox(lpDrawItemStruct);
			return true;
		}

		return false;
	}

	// Paint the status bar with double buffering to avoid flicker
	void DrawStatus(
		HWND hwnd,
		const int iStatusHeight,
		const std::wstring& szStatusData1,
		const int iStatusData1,
		const std::wstring& szStatusData2,
		const int iStatusData2,
		const std::wstring& szStatusInfo)
	{
		auto rcStatus = RECT{};
		GetClientRect(hwnd, &rcStatus);
		if (rcStatus.bottom - rcStatus.top > iStatusHeight)
		{
			rcStatus.top = rcStatus.bottom - iStatusHeight;
		}
		auto ps = PAINTSTRUCT{};
		BeginPaint(hwnd, &ps);

		CDoubleBuffer db;
		db.Begin(ps.hdc, rcStatus);

		auto rcGrad = rcStatus;
		auto rcText = rcGrad;
		const auto crFore = MyGetSysColor(uiColor::StatusText);

		// We start painting a little lower to allow our gradiants to line up
		rcGrad.bottom += GetSystemMetrics(SM_CYSIZEFRAME) - BORDER_VISIBLEWIDTH;
		GradientFillRect(ps.hdc, rcGrad, uiColor::Status);

		rcText.left = rcText.right - iStatusData2;
		DrawSegoeTextW(ps.hdc, szStatusData2, crFore, rcText, true, DT_LEFT | DT_SINGLELINE | DT_BOTTOM);
		rcText.right = rcText.left;
		rcText.left = rcText.right - iStatusData1;
		DrawSegoeTextW(ps.hdc, szStatusData1, crFore, rcText, true, DT_LEFT | DT_SINGLELINE | DT_BOTTOM);
		rcText.right = rcText.left;
		rcText.left = 0;
		DrawSegoeTextW(
			ps.hdc, szStatusInfo, crFore, rcText, true, DT_LEFT | DT_SINGLELINE | DT_BOTTOM | DT_END_ELLIPSIS);

		db.End(ps.hdc);
		EndPaint(hwnd, &ps);
	}

	void GetCaptionRects(
		HWND hWnd,
		RECT* lprcFullCaption,
		RECT* lprcIcon,
		RECT* lprcCloseIcon,
		RECT* lprcMaxIcon,
		RECT* lprcMinIcon,
		RECT* lprcCaptionText) noexcept
	{
		auto rcFullCaption = RECT{};
		auto rcIcon = RECT{};
		auto rcCloseIcon = RECT{};
		auto rcMaxIcon = RECT{};
		auto rcMinIcon = RECT{};
		auto rcCaptionText = RECT{};

		auto rcWindow = RECT{};
		const auto dwWinStyle = GetWindowStyle(hWnd);
		const auto bThickFrame = WS_THICKFRAME == (dwWinStyle & WS_THICKFRAME);

		GetWindowRect(hWnd, &rcWindow); // Get our non-client size
		OffsetRect(&rcWindow, -rcWindow.left, -rcWindow.top); // shift the origin to 0 since that's where our DC paints
		// At this point, we have rectangles for our window and client area
		// rcWindow is the outer rectangle for our NC frame

		const auto cxFixedFrame = GetSystemMetrics(SM_CXFIXEDFRAME);
		const auto cyFixedFrame = GetSystemMetrics(SM_CYFIXEDFRAME);
		const auto cxSizeFrame = GetSystemMetrics(SM_CXSIZEFRAME);
		const auto cySizeFrame = GetSystemMetrics(SM_CYSIZEFRAME);

		auto cxFrame = bThickFrame ? cxSizeFrame : cxFixedFrame;
		auto cyFrame = bThickFrame ? cySizeFrame : cyFixedFrame;

		// If padded borders are in effect, we fall back to a single width for both thick and thin frames
		const auto cxPaddedBorder = GetSystemMetrics(SM_CXPADDEDBORDER);
		if (cxPaddedBorder)
		{
			cxFrame = cxSizeFrame + cxPaddedBorder;
			cyFrame = cySizeFrame + cxPaddedBorder;
		}

		rcFullCaption.top = rcWindow.top + BORDER_VISIBLEWIDTH;
		rcFullCaption.left = rcWindow.left + BORDER_VISIBLEWIDTH;
		rcFullCaption.right = rcWindow.right - BORDER_VISIBLEWIDTH;
		rcFullCaption.bottom = rcCaptionText.bottom = rcWindow.top + cyFrame + GetSystemMetrics(SM_CYCAPTION);

		rcIcon.top = rcWindow.top + cyFrame + GetSystemMetrics(SM_CYEDGE);
		rcIcon.left = rcWindow.left + cxFrame + cxFixedFrame;
		rcIcon.bottom = rcIcon.top + GetSystemMetrics(SM_CYSMICON);
		rcIcon.right = rcIcon.left + GetSystemMetrics(SM_CXSMICON);

		const auto cxBorder = GetSystemMetrics(SM_CXBORDER);
		const auto cyBorder = GetSystemMetrics(SM_CYBORDER);
		const auto buttonSize = GetSystemMetrics(SM_CYSIZE) - 2 * cyBorder;

		rcCloseIcon.top = rcMaxIcon.top = rcMinIcon.top = rcWindow.top + cyFrame + cyBorder;
		rcCloseIcon.bottom = rcMaxIcon.bottom = rcMinIcon.bottom = rcCloseIcon.top + buttonSize;
		rcCloseIcon.right = rcWindow.right - cxFrame - cxBorder;
		rcCloseIcon.left = rcMaxIcon.right = rcCloseIcon.right - buttonSize;

		rcMaxIcon.left = rcMinIcon.right = rcMaxIcon.right - buttonSize;

		rcMinIcon.left = rcMinIcon.right - buttonSize;

		InflateRect(&rcCloseIcon, -1, -1);
		InflateRect(&rcMaxIcon, -1, -1);
		InflateRect(&rcMinIcon, -1, -1);

		rcCaptionText.top = rcWindow.top + cxFrame;
		if (DS_MODALFRAME == (dwWinStyle & DS_MODALFRAME))
		{
			rcCaptionText.left = rcWindow.left + BORDER_VISIBLEWIDTH + cxFixedFrame + cxBorder;
		}
		else
		{
			rcCaptionText.left = rcIcon.right + cxFixedFrame + cxBorder;
		}

		rcCaptionText.right = rcCloseIcon.left;

		if (WS_MINIMIZEBOX == (dwWinStyle & WS_MINIMIZEBOX))
			rcCaptionText.right = rcMinIcon.left;
		else if (WS_MAXIMIZEBOX == (dwWinStyle & WS_MAXIMIZEBOX))
			rcCaptionText.right = rcMaxIcon.left;

		if (lprcFullCaption) *lprcFullCaption = rcFullCaption;
		if (lprcIcon) *lprcIcon = rcIcon;
		if (lprcCloseIcon) *lprcCloseIcon = rcCloseIcon;
		if (lprcMaxIcon) *lprcMaxIcon = rcMaxIcon;
		if (lprcMinIcon) *lprcMinIcon = rcMinIcon;
		if (lprcCaptionText) *lprcCaptionText = rcCaptionText;
	}

	void DrawSystemButtons(_In_ HWND hWnd, _In_opt_ HDC hdc, const LONG_PTR iHitTest, const bool bHover) noexcept
	{
		HDC hdcLocal = nullptr;
		if (!hdc)
		{
			hdcLocal = GetWindowDC(hWnd);
			hdc = hdcLocal;
		}

		const auto dwWinStyle = GetWindowStyle(hWnd);
		const auto bMinBox = WS_MINIMIZEBOX == (dwWinStyle & WS_MINIMIZEBOX);
		const auto bMaxBox = WS_MAXIMIZEBOX == (dwWinStyle & WS_MAXIMIZEBOX);

		auto rcCloseIcon = RECT{};
		auto rcMaxIcon = RECT{};
		auto rcMinIcon = RECT{};
		GetCaptionRects(hWnd, nullptr, nullptr, &rcCloseIcon, &rcMaxIcon, &rcMinIcon, nullptr);

		// Draw our system buttons appropriately
		const auto htClose = HTCLOSE == iHitTest;
		const auto htMax = HTMAXBUTTON == iHitTest;
		const auto htMin = HTMINBUTTON == iHitTest;

		DrawBitmap(hdc, rcCloseIcon, uiBitmap::Close, htClose, htClose && !bHover ? 2 : 0);

		if (bMaxBox)
		{
			DrawBitmap(
				hdc,
				rcMaxIcon,
				IsZoomed(hWnd) ? uiBitmap::Restore : uiBitmap::Maximize,
				htMax,
				htMax && !bHover ? 2 : 0);
		}

		if (bMinBox)
		{
			DrawBitmap(hdc, rcMinIcon, uiBitmap::Minimize, htMin, htMin && !bHover ? 2 : 0);
		}

		if (hdcLocal) ReleaseDC(hWnd, hdcLocal);
	}

	void DrawWindowFrame(_In_ HWND hWnd, const bool bActive, int iStatusHeight)
	{
		const auto cxPaddedBorder = GetSystemMetrics(SM_CXPADDEDBORDER);
		const auto cxFixedFrame = GetSystemMetrics(SM_CXFIXEDFRAME);
		const auto cyFixedFrame = GetSystemMetrics(SM_CYFIXEDFRAME);
		const auto cxSizeFrame = GetSystemMetrics(SM_CXSIZEFRAME);
		const auto cySizeFrame = GetSystemMetrics(SM_CYSIZEFRAME);

		const auto dwWinStyle = GetWindowStyle(hWnd);
		const auto bModal = DS_MODALFRAME == (dwWinStyle & DS_MODALFRAME);
		const auto bThickFrame = WS_THICKFRAME == (dwWinStyle & WS_THICKFRAME);

		auto rcWindow = RECT{};
		auto rcClient = RECT{};
		GetWindowRect(hWnd, &rcWindow); // Get our non-client size
		GetClientRect(hWnd, &rcClient); // Get our client size
		// locate our client rect on the screen
		if (::MapWindowPoints(hWnd, nullptr, reinterpret_cast<LPPOINT>(&rcClient), 2) == 0) return;

		// Before we fiddle with our client and window rects further, paint the menu
		// This must be in window coordinates for WM_NCPAINT to work!!!
		auto rcMenu = rcWindow;
		rcMenu.top = rcMenu.top + cySizeFrame + GetSystemMetrics(SM_CYCAPTION);
		rcMenu.bottom = rcClient.top;
		rcMenu.left += cxSizeFrame;
		rcMenu.right -= cxSizeFrame;
		auto hRgnCaption = CreateRectRgnIndirect(&rcMenu);
		::DefWindowProc(hWnd, WM_NCPAINT, reinterpret_cast<WPARAM>(hRgnCaption), NULL);
		DeleteObject(hRgnCaption);

		OffsetRect(&rcClient, -rcWindow.left, -rcWindow.top); // shift the origin to 0 since that's where our DC paints
		OffsetRect(&rcWindow, -rcWindow.left, -rcWindow.top); // shift the origin to 0 since that's where our DC paints
		// At this point, we have rectangles for our window and client area
		// rcWindow is the outer rectangle for our NC frame
		// rcClient is the inner rectangle for our NC frame

		// Make sure our status bar doesn't bleed into the menus on short windows
		if (rcClient.bottom - rcClient.top < iStatusHeight)
		{
			iStatusHeight = rcClient.bottom - rcClient.top;
		}

		auto cxFrame = bThickFrame ? cxSizeFrame : cxFixedFrame;
		auto cyFrame = bThickFrame ? cySizeFrame : cyFixedFrame;

		// If padded borders are in effect, we fall back to a single width for both thick and thin frames
		if (cxPaddedBorder)
		{
			cxFrame = cxSizeFrame + cxPaddedBorder;
			cyFrame = cySizeFrame + cxPaddedBorder;
		}

		// The menu and caption have odd borders we've not painted yet - compute rectangles so we can paint them
		auto rcFullCaption = RECT{};
		auto rcIcon = RECT{};
		auto rcCaptionText = RECT{};
		auto rcMenuGutterLeft = RECT{};
		auto rcMenuGutterRight = RECT{};
		auto rcWindowGutterLeft = RECT{};
		auto rcWindowGutterRight = RECT{};

		GetCaptionRects(hWnd, &rcFullCaption, &rcIcon, nullptr, nullptr, nullptr, &rcCaptionText);

		rcMenuGutterLeft.left = rcWindowGutterLeft.left = rcFullCaption.left;
		rcMenuGutterRight.right = rcWindowGutterRight.right = rcFullCaption.right;
		rcMenuGutterLeft.top = rcMenuGutterRight.top = rcFullCaption.bottom;

		rcMenuGutterLeft.right = rcWindow.left + cxFrame;
		rcMenuGutterRight.left = rcWindow.right - cxFrame;
		rcMenuGutterLeft.bottom = rcMenuGutterRight.bottom = rcWindowGutterLeft.top = rcWindowGutterRight.top =
			rcClient.top;

		rcWindowGutterLeft.bottom = rcWindowGutterRight.bottom = rcClient.bottom - iStatusHeight;

		rcWindowGutterLeft.right = rcClient.left;
		rcWindowGutterRight.left = rcClient.right;

		// Get a DC where the upper left corner of the window is 0,0
		// All of our rectangles were computed with this DC in mind
		const auto hdcWin = GetWindowDC(hWnd);
		if (hdcWin)
		{
			CDoubleBuffer db;
			auto hdc = hdcWin;
			db.Begin(hdc, rcWindow);

			// Copy the current screen over to preserve it
			BitBlt(
				hdc,
				rcWindow.left,
				rcWindow.top,
				rcWindow.right - rcWindow.left,
				rcWindow.bottom - rcWindow.top,
				hdcWin,
				0,
				0,
				SRCCOPY);

			// Draw a line under the menu from gutter to gutter
			const auto hpenOld = SelectObject(hdc, GetPen(uiPen::SolidGreyPen));
			MoveToEx(hdc, rcMenuGutterLeft.right, rcClient.top - 1, nullptr);
			LineTo(hdc, rcMenuGutterRight.left, rcClient.top - 1);
			static_cast<void>(SelectObject(hdc, hpenOld));

			// White out the caption
			FillRect(hdc, &rcFullCaption, GetSysBrush(uiColor::Background));

			DrawSystemButtons(hWnd, hdc, HTNOWHERE, false);

			// Draw our icon
			if (!bModal)
			{
				const auto hIcon = static_cast<HICON>(::LoadImage(
					AfxGetInstanceHandle(),
					MAKEINTRESOURCE(IDR_MAINFRAME),
					IMAGE_ICON,
					rcIcon.right - rcIcon.left,
					rcIcon.bottom - rcIcon.top,
					LR_DEFAULTCOLOR));

				DrawIconEx(
					hdc,
					rcIcon.left,
					rcIcon.top,
					hIcon,
					rcIcon.right - rcIcon.left,
					rcIcon.bottom - rcIcon.top,
					NULL,
					GetSysBrush(uiColor::Background),
					DI_NORMAL);

				DestroyIcon(hIcon);
			}

			// Fill in our gutters
			FillRect(hdc, &rcMenuGutterLeft, GetSysBrush(uiColor::Background));
			FillRect(hdc, &rcMenuGutterRight, GetSysBrush(uiColor::Background));
			FillRect(hdc, &rcWindowGutterLeft, GetSysBrush(uiColor::Background));
			FillRect(hdc, &rcWindowGutterRight, GetSysBrush(uiColor::Background));

			if (iStatusHeight)
			{
				auto rcStatus = rcClient;
				rcStatus.top = rcClient.bottom - iStatusHeight;

				auto rcFullStatus = RECT{};
				rcFullStatus.top = rcStatus.top;
				rcFullStatus.left = rcWindow.left + BORDER_VISIBLEWIDTH;
				rcFullStatus.right = rcWindow.right - BORDER_VISIBLEWIDTH;
				rcFullStatus.bottom = rcWindow.bottom - BORDER_VISIBLEWIDTH;

				ExcludeClipRect(hdc, rcStatus.left, rcStatus.top, rcStatus.right, rcStatus.bottom);
				GradientFillRect(hdc, rcFullStatus, uiColor::Status);
			}
			else
			{
				auto rcBottomGutter = RECT{};
				rcBottomGutter.top = rcWindow.bottom - cyFrame;
				rcBottomGutter.left = rcWindow.left + BORDER_VISIBLEWIDTH;
				rcBottomGutter.right = rcWindow.right - BORDER_VISIBLEWIDTH;
				rcBottomGutter.bottom = rcWindow.bottom - BORDER_VISIBLEWIDTH;

				FillRect(hdc, &rcBottomGutter, GetSysBrush(uiColor::Background));
			}

			const WCHAR szTitle[256] = {};
			DefWindowProcW(hWnd, WM_GETTEXT, static_cast<WPARAM>(_countof(szTitle)), reinterpret_cast<LPARAM>(szTitle));
			DrawSegoeTextW(
				hdc, szTitle, MyGetSysColor(uiColor::Text), rcCaptionText, false, DT_LEFT | DT_SINGLELINE | DT_VCENTER);

			// Finally, we paint our border glow if we're not maximized
			if (!IsZoomed(hWnd))
			{
				auto rcInnerFrame = rcWindow;
				InflateRect(&rcInnerFrame, -BORDER_VISIBLEWIDTH, -BORDER_VISIBLEWIDTH);
				ExcludeClipRect(hdc, rcInnerFrame.left, rcInnerFrame.top, rcInnerFrame.right, rcInnerFrame.bottom);
				FillRect(hdc, &rcWindow, GetSysBrush(bActive ? uiColor::Glow : uiColor::FrameUnselected));
			}

			db.End(hdc);
			ReleaseDC(hWnd, hdcWin);
		}
	}

	_Check_return_ bool
	HandleControlUI(const UINT message, const WPARAM wParam, const LPARAM lParam, _Out_ LRESULT* lpResult)
	{
		if (!lpResult) return false;
		*lpResult = 0;
		switch (message)
		{
		case WM_DRAWITEM:
			if (DrawItem(reinterpret_cast<LPDRAWITEMSTRUCT>(lParam))) return true;
			break;
		case WM_MEASUREITEM:
			if (MeasureItem(reinterpret_cast<LPMEASUREITEMSTRUCT>(lParam))) return true;
			break;
		case WM_ERASEBKGND:
			return true;
		case WM_CTLCOLORSTATIC:
		case WM_CTLCOLOREDIT:
		{
			const auto lsStyle = static_cast<uiLabelStyle>(
				reinterpret_cast<intptr_t>(::GetProp(reinterpret_cast<HWND>(lParam), LABEL_STYLE)));
			auto uiText = uiColor::Text;
			auto uiBackground = registry::uiDiag ? uiColor::TestPink : uiColor::Background;

			if (lsStyle == uiLabelStyle::PaneHeaderLabel || lsStyle == uiLabelStyle::PaneHeaderText)
			{
				uiText = uiColor::PaneHeaderText;
				uiBackground = registry::uiDiag ? uiColor::TestRed : uiColor::PaneHeaderBackground;
			}

			const auto hdc = reinterpret_cast<HDC>(wParam);
			if (hdc)
			{
				SetTextColor(hdc, MyGetSysColor(uiText));
				SetBkMode(hdc, TRANSPARENT);
				SelectObject(hdc, GetSegoeFont());
			}

			*lpResult = reinterpret_cast<LRESULT>(GetSysBrush(uiBackground));
			return true;
		}
		case WM_NOTIFY:
		{
			const auto pHdr = reinterpret_cast<LPNMHDR>(lParam);

			switch (pHdr->code)
			{
			// Paint Buttons
			case NM_CUSTOMDRAW:
				return CustomDrawButton(pHdr, lpResult);
			}
		}
		break;
		}

		return false;
	}

	void DrawHelpText(_In_ HWND hWnd, _In_ const UINT uIDText)
	{
		auto ps = PAINTSTRUCT{};
		BeginPaint(hWnd, &ps);

		if (ps.hdc)
		{
			auto rcWin = RECT{};
			GetClientRect(hWnd, &rcWin);

			CDoubleBuffer db;
			auto hdc = ps.hdc;
			db.Begin(hdc, rcWin);

			auto rcText = rcWin;
			FillRect(hdc, &rcText, GetSysBrush(uiColor::Background));

			DrawSegoeTextW(
				hdc,
				strings::loadstring(uIDText),
				MyGetSysColor(uiColor::TextDisabled),
				rcText,
				true,
				DT_CENTER | DT_SINGLELINE | DT_VCENTER | DT_NOCLIP | DT_END_ELLIPSIS | DT_NOPREFIX);

			db.End(hdc);
		}

		EndPaint(hWnd, &ps);
	}

	// Handle WM_ERASEBKGND so the control won't flicker.
	LRESULT CALLBACK LabelProc(
		_In_ HWND hWnd,
		const UINT uMsg,
		const WPARAM wParam,
		const LPARAM lParam,
		const UINT_PTR uIdSubclass,
		DWORD_PTR /*dwRefData*/) noexcept
	{
		switch (uMsg)
		{
		case WM_NCDESTROY:
			RemoveWindowSubclass(hWnd, LabelProc, uIdSubclass);
			return DefSubclassProc(hWnd, uMsg, wParam, lParam);
		case WM_ERASEBKGND:
			return true;
		}

		return DefSubclassProc(hWnd, uMsg, wParam, lParam);
	}

	void SubclassLabel(_In_ HWND hWnd) noexcept
	{
		SetWindowSubclass(hWnd, LabelProc, 0, 0);
		SendMessageA(hWnd, WM_SETFONT, reinterpret_cast<WPARAM>(GetSegoeFont()), false);
		::SendMessage(hWnd, EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, 0);
	}

	void StyleLabel(_In_ HWND hWnd, uiLabelStyle lsStyle)
	{
		EC_B_S(::SetProp(hWnd, LABEL_STYLE, reinterpret_cast<HANDLE>(lsStyle)));
	}

	// Returns the first visible top level window in the current process
	_Check_return_ HWND GetMainWindow() noexcept
	{
		auto hwndRet = HWND{};
		static auto currentPid = GetCurrentProcessId();
		const auto enumProc = [](auto hwnd, auto lParam) {
			// Use of BOOL return type forced by WNDENUMPROC signature
			if (!lParam) return FALSE;
			const auto ret = reinterpret_cast<HWND*>(lParam);
			auto pid = ULONG{};
			GetWindowThreadProcessId(hwnd, &pid);
			if (currentPid == pid && GetWindow(hwnd, GW_OWNER) == nullptr && IsWindowVisible(hwnd))
			{
				*ret = hwnd;
				return FALSE;
			}

			return TRUE;
		};

		EnumWindows(enumProc, reinterpret_cast<LPARAM>(&hwndRet));

		return hwndRet;
	}

	HDWP WINAPI DeferWindowPos(
		_In_ HDWP hWinPosInfo,
		_In_ HWND hWnd,
		_In_ int x,
		_In_ int y,
		_In_ int cx,
		_In_ int cy,
		_In_ const WCHAR* szName,
		_In_opt_ const WCHAR* szLabel)
	{
		if (szLabel)
		{
			output::DebugPrint(
				output::dbgLevel::Draw,
				L"%ws x:%d y:%d width:%d height:%d v:%d label:\"%ws\"\n",
				szName,
				x,
				y,
				cx,
				cy,
				::IsWindowVisible(hWnd),
				szLabel);
		}
		else
		{
			output::DebugPrint(
				output::dbgLevel::Draw,
				L"%ws x:%d y:%d cx:%d cy:%d v:%d\n",
				szName,
				x,
				y,
				cx,
				cy,
				::IsWindowVisible(hWnd));
		}

		return EC_D(HDWP, ::DeferWindowPos(hWinPosInfo, hWnd, nullptr, x, y, cx, cy, SWP_NOZORDER));
	}

	// Draw a frame but don't touch the inside of the frame
	void WINAPI FrameRect(_In_ HDC hDC, _In_ RECT rect, _In_ int width, _In_ const uiColor color)
	{
		auto rcInnerFrame = rect;
		WC_D_S(::InflateRect(&rcInnerFrame, -width, -width));
		auto rgn = WC_D(HRGN, ::CreateRectRgn(0, 0, 0, 0));
		if (GetClipRgn(hDC, rgn) != 1)
		{
			WC_B_S(::DeleteObject(rgn));
			rgn = nullptr;
		}

		WC_D_S(::ExcludeClipRect(hDC, rcInnerFrame.left, rcInnerFrame.top, rcInnerFrame.right, rcInnerFrame.bottom));
		WC_D_S(::FillRect(hDC, &rect, GetSysBrush(color)));

		WC_D_S(::SelectClipRgn(hDC, rgn));
		if (rgn != nullptr)
		{
			WC_B_S(::DeleteObject(rgn));
		}
	}
} // namespace ui