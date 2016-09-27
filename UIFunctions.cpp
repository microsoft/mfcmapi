// UIFunctions.h : Common UI functions for MFCMAPI
#include "stdafx.h"
#include "UIFunctions.h"
#include "Windowsx.h"
#include "ImportProcs.h"
#include "RichEditOleCallback.h"
#include "SortList/SortListData.h"
#include "SortList/NodeData.h"

HFONT g_hFontSegoe = nullptr;
HFONT g_hFontSegoeBold = nullptr;

#define SEGOE _T("Segoe UI") // STRING_OK
#define SEGOEW L"Segoe UI" // STRING_OK
#define SEGOEBOLD L"Segoe UI Bold" // STRING_OK
#define BUTTON_STYLE _T("ButtonStyle") // STRING_OK
#define LABEL_STYLE _T("LabelStyle") // STRING_OK

#define BORDER_VISIBLEWIDTH 2
#define TRIANGLE_SIZE 4

enum myColor
{
	cWhite,
	cLightGrey,
	cGrey,
	cDarkGrey,
	cBlack,
	cCyan,
	cMagenta,
	cBlue,
	cMedBlue,
	cPaleBlue,
	cColorEnd
};

// Keep in sync with enum myColor
COLORREF g_Colors[cColorEnd] =
{
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
};

// Fixed mapping of UI elements to colors
// Will be overridden by system colors when specified in g_SysColors
// Keep in sync with enum uiColor
// We can swap in cCyan for various entries to test coverage
myColor g_FixedColors[cUIEnd] =
{
 cWhite, // cBackground
 cLightGrey, // cBackgroundReadOnly
 cBlue, // cGlow
 cPaleBlue, // cGlowBackground
 cBlack, // cGlowText
 cBlack, // cFrameSelected
 cGrey, // cFrameUnselected
 cMedBlue, // cSelectedBackground
 cGrey, // cArrow
 cBlack, // cText
 cGrey, // cTextDisabled
 cBlack, // cTextReadOnly
 cMagenta, // cBitmapTransBack
 cCyan, // cBitmapTransFore
 cBlue, // cStatus
 cWhite, // cStatusText
 cPaleBlue, // cPaneHeaderBackground,
 cBlack, // cPaneHeaderText,
};

// Mapping of UI elements to system colors
// NULL entries will get the fixed mapping from g_FixedColors
int g_SysColors[cUIEnd] =
{
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
int g_BitmapResources[cBitmapEnd] =
{
 IDB_ADVISE, // cNotify,
 IDB_CLOSE, // cClose,
 IDB_MINIMIZE, // cMinimize,
 IDB_MAXIMIZE, // cMaximize,
 IDB_RESTORE, // cRestore,
};

HBRUSH g_FixedBrushes[cColorEnd] = { nullptr };
HBRUSH g_SysBrushes[cUIEnd] = { nullptr };
HPEN g_Pens[cPenEnd] = { nullptr };
HBITMAP g_Bitmaps[cBitmapEnd] = { nullptr };

void DrawSegoeTextW(
	_In_ HDC hdc,
	_In_z_ LPCWSTR lpchText,
	_In_ COLORREF color,
	_In_ LPRECT lprc,
	bool bBold,
	_In_ UINT format);

void DrawSegoeTextA(
	_In_ HDC hdc,
	_In_z_ LPCSTR lpchText,
	_In_ COLORREF color,
	_In_ LPRECT lprc,
	bool bBold,
	_In_ UINT format);

#ifdef UNICODE
#define DrawSegoeText DrawSegoeTextW
#else
#define DrawSegoeText DrawSegoeTextA
#endif

CDoubleBuffer::CDoubleBuffer() : m_hdcMem(nullptr), m_hbmpMem(nullptr), m_hdcPaint(nullptr)
{
	ZeroMemory(&m_rcPaint, sizeof m_rcPaint);
}

CDoubleBuffer::~CDoubleBuffer()
{
	Cleanup();
}

void CDoubleBuffer::Begin(_Inout_ HDC& hdc, _In_ RECT CONST* prcPaint)
{
	if (hdc)
	{
		m_hdcMem = ::CreateCompatibleDC(hdc);
		if (m_hdcMem)
		{
			m_hbmpMem = ::CreateCompatibleBitmap(
				hdc,
				prcPaint->right - prcPaint->left,
				prcPaint->bottom - prcPaint->top);

			if (m_hbmpMem)
			{
				(void) ::SelectObject(m_hdcMem, m_hbmpMem);
				(void) ::CopyRect(&m_rcPaint, prcPaint);
				(void) ::OffsetWindowOrgEx(m_hdcMem,
					m_rcPaint.left,
					m_rcPaint.top,
					nullptr);

				(void) ::SelectObject(m_hdcMem, GetCurrentObject(hdc, OBJ_FONT));
				(void) ::SelectObject(m_hdcMem, GetCurrentObject(hdc, OBJ_BRUSH));
				(void) ::SelectObject(m_hdcMem, GetCurrentObject(hdc, OBJ_PEN));
				// cache the original DC and pass out the memory DC
				m_hdcPaint = hdc;
				hdc = m_hdcMem;
			}
			else
			{
				Cleanup();
			}
		}
	}
}

void CDoubleBuffer::End(_Inout_ HDC& hdc)
{
	if (hdc && hdc == m_hdcMem)
	{
		::BitBlt(m_hdcPaint,
			m_rcPaint.left,
			m_rcPaint.top,
			m_rcPaint.right - m_rcPaint.left,
			m_rcPaint.bottom - m_rcPaint.top,
			m_hdcMem,
			m_rcPaint.left,
			m_rcPaint.top,
			SRCCOPY);

		// restore the original DC
		hdc = m_hdcPaint;

		Cleanup();
	}
}

void CDoubleBuffer::Cleanup()
{
	if (m_hbmpMem)
	{
		(void) ::DeleteObject(m_hbmpMem);
		m_hbmpMem = nullptr;
	}

	if (m_hdcMem)
	{
		(void) ::DeleteDC(m_hdcMem);
		m_hdcMem = nullptr;
	}

	m_hdcPaint = nullptr;
	ZeroMemory(&m_rcPaint, sizeof m_rcPaint);
}

void InitializeGDI()
{
}

void UninitializeGDI()
{
	if (g_hFontSegoe) ::DeleteObject(g_hFontSegoe);
	g_hFontSegoe = nullptr;
	if (g_hFontSegoeBold) ::DeleteObject(g_hFontSegoeBold);
	g_hFontSegoeBold = nullptr;

	for (auto i = 0; i < cColorEnd; i++)
	{
		if (g_FixedBrushes[i])
		{
			::DeleteObject(g_FixedBrushes[i]);
			g_FixedBrushes[i] = nullptr;
		}
	}
	for (auto i = 0; i < cUIEnd; i++)
	{
		if (g_SysBrushes[i])
		{
			::DeleteObject(g_SysBrushes[i]);
			g_SysBrushes[i] = nullptr;
		}
	}

	for (auto i = 0; i < cPenEnd; i++)
	{
		if (g_Pens[i])
		{
			::DeleteObject(g_Pens[i]);
			g_Pens[i] = nullptr;
		}
	}

	for (auto i = 0; i < cBitmapEnd; i++)
	{
		if (g_Bitmaps[i])
		{
			::DeleteObject(g_Bitmaps[i]);
			g_Bitmaps[i] = nullptr;
		}
	}
}

_Check_return_ LPMENUENTRY CreateMenuEntry(_In_z_ LPCWSTR szMenu)
{
	auto hRes = S_OK;
	auto lpMenu = new MenuEntry;
	if (lpMenu)
	{
		ZeroMemory(lpMenu, sizeof(MenuEntry));
		lpMenu->m_MSAA.dwMSAASignature = MSAA_MENU_SIG;

		size_t iLen = 0;
		WC_H(StringCchLengthW(szMenu, STRSAFE_MAX_CCH, &iLen));

		lpMenu->m_MSAA.pszWText = new WCHAR[iLen + 1];
		if (lpMenu->m_MSAA.pszWText)
		{
			WC_H(StringCchCopyW(lpMenu->m_MSAA.pszWText, iLen + 1, szMenu));
			lpMenu->m_pName = lpMenu->m_MSAA.pszWText;
			lpMenu->m_MSAA.cchWText = static_cast<DWORD>(iLen);
		}
		return lpMenu;
	}
	return nullptr;
}

_Check_return_ LPMENUENTRY CreateMenuEntry(UINT iudMenu)
{
	return CreateMenuEntry(loadstring(iudMenu).c_str());
}

void DeleteMenuEntry(_In_ LPMENUENTRY lpMenu)
{
	if (!lpMenu) return;
	if (lpMenu->m_MSAA.pszWText) delete[] lpMenu->m_MSAA.pszWText;
	delete lpMenu;
}

void DeleteMenuEntries(_In_ HMENU hMenu)
{
	UINT nCount = ::GetMenuItemCount(hMenu);
	if (-1 == nCount) return;

	for (UINT nPosition = 0; nPosition < nCount; nPosition++)
	{
		MENUITEMINFOW menuiteminfo = { 0 };
		menuiteminfo.cbSize = sizeof(MENUITEMINFOW);
		menuiteminfo.fMask = MIIM_DATA | MIIM_SUBMENU;

		::GetMenuItemInfoW(hMenu, nPosition, true, &menuiteminfo);
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
void ConvertMenuOwnerDraw(_In_ HMENU hMenu, bool bRoot)
{
	UINT nCount = ::GetMenuItemCount(hMenu);
	if (-1 == nCount) return;

	for (UINT nPosition = 0; nPosition < nCount; nPosition++)
	{
		MENUITEMINFOW menuiteminfo = { 0 };
		WCHAR szMenu[128] = { 0 };
		menuiteminfo.cbSize = sizeof(MENUITEMINFOW);
		menuiteminfo.fMask = MIIM_STRING | MIIM_SUBMENU | MIIM_FTYPE;
		menuiteminfo.cch = _countof(szMenu);
		menuiteminfo.dwTypeData = szMenu;

		::GetMenuItemInfoW(hMenu, nPosition, true, &menuiteminfo);
		auto bOwnerDrawn = (menuiteminfo.fType & MF_OWNERDRAW) != 0;
		if (!bOwnerDrawn)
		{
			auto lpMenuEntry = CreateMenuEntry(szMenu);
			if (lpMenuEntry)
			{
				lpMenuEntry->b_OnMenuBar = bRoot;
				menuiteminfo.fMask = MIIM_DATA | MIIM_FTYPE;
				menuiteminfo.fType |= MF_OWNERDRAW;
				menuiteminfo.dwItemData = reinterpret_cast<ULONG_PTR>(lpMenuEntry);
				::SetMenuItemInfoW(hMenu, nPosition, true, &menuiteminfo);
			}
		}

		if (menuiteminfo.hSubMenu)
		{
			ConvertMenuOwnerDraw(menuiteminfo.hSubMenu, false);
		}
	}
}

void UpdateMenuString(_In_ HWND hWnd, UINT uiMenuTag, UINT uidNewString)
{
	auto hRes = S_OK;

	auto szNewString = loadstring(uidNewString);

	DebugPrint(DBGMenu, L"UpdateMenuString: Changing menu item 0x%X on window %p to \"%ws\"\n", uiMenuTag, hWnd, szNewString.c_str());
	auto hMenu = ::GetMenu(hWnd);
	if (!hMenu) return;

	MENUITEMINFOW MenuInfo = { 0 };

	ZeroMemory(&MenuInfo, sizeof MenuInfo);

	MenuInfo.cbSize = sizeof MenuInfo;
	MenuInfo.fMask = MIIM_STRING;
	MenuInfo.dwTypeData = LPWSTR(szNewString.c_str());

	EC_B(SetMenuItemInfoW(
		hMenu,
		uiMenuTag,
		false,
		&MenuInfo));
}

void MergeMenu(_In_ HMENU hMenuDestination, _In_ const HMENU hMenuAdd)
{
	::GetMenuItemCount(hMenuDestination);
	auto iMenuDestItemCount = ::GetMenuItemCount(hMenuDestination);
	auto iMenuAddItemCount = ::GetMenuItemCount(hMenuAdd);

	if (iMenuAddItemCount == 0) return;

	if (iMenuDestItemCount > 0) ::AppendMenu(hMenuDestination, MF_SEPARATOR, NULL, nullptr);

	WCHAR szMenu[128] = { 0 };
	for (auto iLoop = 0; iLoop < iMenuAddItemCount; iLoop++)
	{
		::GetMenuStringW(hMenuAdd, iLoop, szMenu, _countof(szMenu), MF_BYPOSITION);

		auto hSubMenu = ::GetSubMenu(hMenuAdd, iLoop);
		if (!hSubMenu)
		{
			auto nState = GetMenuState(hMenuAdd, iLoop, MF_BYPOSITION);
			auto nItemID = GetMenuItemID(hMenuAdd, iLoop);

			::AppendMenuW(hMenuDestination, nState, nItemID, szMenu);
			iMenuDestItemCount++;
		}
		else
		{
			auto iInsertPosDefault = GetMenuItemCount(hMenuDestination);

			auto hNewPopupMenu = ::CreatePopupMenu();

			MergeMenu(hNewPopupMenu, hSubMenu);

			::InsertMenuW(
				hMenuDestination,
				iInsertPosDefault,
				MF_BYPOSITION | MF_POPUP | MF_ENABLED,
				reinterpret_cast<UINT_PTR>(hNewPopupMenu),
				szMenu);
			iMenuDestItemCount++;
		}
	}
}

void DisplayContextMenu(UINT uiClassMenu, UINT uiControlMenu, _In_ HWND hWnd, int x, int y)
{
	if (!uiClassMenu) uiClassMenu = IDR_MENU_DEFAULT_POPUP;
	auto hPopup = ::CreateMenu();
	auto hContext = ::LoadMenu(nullptr, MAKEINTRESOURCE(uiClassMenu));

	::InsertMenu(
		hPopup,
		0,
		MF_BYPOSITION | MF_POPUP,
		reinterpret_cast<UINT_PTR>(hContext),
		_T(""));

	auto hRealPopup = ::GetSubMenu(hPopup, 0);

	if (hRealPopup)
	{
		if (uiControlMenu)
		{
			auto hAppended = ::LoadMenu(nullptr, MAKEINTRESOURCE(uiControlMenu));
			if (hAppended)
			{
				MergeMenu(hRealPopup, hAppended);
				::DestroyMenu(hAppended);
			}
		}

		if (IDR_MENU_PROPERTY_POPUP == uiClassMenu)
		{
			(void)ExtendAddInMenu(hRealPopup, MENU_CONTEXT_PROPERTY);
		}

		ConvertMenuOwnerDraw(hRealPopup, false);
		::TrackPopupMenu(hRealPopup, TPM_LEFTALIGN | TPM_RIGHTBUTTON, x, y, NULL, hWnd, nullptr);
		DeleteMenuEntries(hRealPopup);
	}

	::DestroyMenu(hContext);
	::DestroyMenu(hPopup);
}

HMENU LocateSubmenu(_In_ HMENU hMenu, UINT uid)
{
	UINT nCount = ::GetMenuItemCount(hMenu);
	if (-1 == nCount) return nullptr;

	for (UINT nPosition = 0; nPosition < nCount; nPosition++)
	{
		MENUITEMINFOW menuiteminfo = { 0 };
		menuiteminfo.cbSize = sizeof(MENUITEMINFOW);
		menuiteminfo.fMask = MIIM_SUBMENU | MIIM_ID;

		::GetMenuItemInfoW(hMenu, nPosition, true, &menuiteminfo);

		if (menuiteminfo.wID == uid) return hMenu;

		if (menuiteminfo.hSubMenu)
		{
			auto hSub = LocateSubmenu(menuiteminfo.hSubMenu, uid);
			if (hSub) return hSub;
		}
	}
	return nullptr;
}

_Check_return_ int GetEditHeight(_In_ HWND hwndEdit)
{
	auto hRes = S_OK;
	HDC hdc = nullptr;
	TEXTMETRIC tmFont = { 0 };
	auto iHeight = 0;

	// Get the DC for the edit control.
	WC_D(hdc, GetDC(hwndEdit));

	if (hdc)
	{
		// Get the metrics for the Segoe font.
		auto hOldFont = static_cast<HFONT>(SelectObject(hdc, GetSegoeFont()));
		WC_B(::GetTextMetrics(hdc, &tmFont));
		SelectObject(hdc, hOldFont);
		ReleaseDC(hwndEdit, hdc);
	}

	// Calculate the new height for the edit control.
	iHeight =
		tmFont.tmHeight
		+ 2 * GetSystemMetrics(SM_CYFIXEDFRAME) // Adjust for the edit border
		+ 2 * GetSystemMetrics(SM_CXEDGE); // Adjust for the edit border
	return iHeight;
}

_Check_return_ int GetTextHeight(_In_ HWND hwndEdit)
{
	auto hRes = S_OK;
	HDC hdc = nullptr;
	TEXTMETRIC tmFont = { 0 };
	auto iHeight = 0;

	// Get the DC for the edit control.
	WC_D(hdc, GetDC(hwndEdit));

	if (hdc)
	{
		// Get the metrics for the Segoe font.
		auto hOldFont = static_cast<HFONT>(SelectObject(hdc, GetSegoeFont()));
		WC_B(::GetTextMetrics(hdc, &tmFont));
		SelectObject(hdc, hOldFont);
		ReleaseDC(hwndEdit, hdc);
	}

	// Calculate the new height for the static control.
	iHeight = tmFont.tmHeight;
	return iHeight;
}

SIZE GetTextExtentPoint32(HDC hdc, wstring szText)
{
	SIZE size = { 0 };
	::GetTextExtentPoint32W(hdc, szText.c_str(), static_cast<int>(szText.length()), &size);
	return size;
}

int CALLBACK EnumFontFamExProcW(
	_In_ LPLOGFONTW lplf,
	_In_ NEWTEXTMETRICEXW* /*lpntme*/,
	DWORD /*FontType*/,
	LPARAM lParam)
{
	// Use a 9 point font
	lplf->lfHeight = -MulDiv(9, GetDeviceCaps(::GetDC(nullptr), LOGPIXELSY), 72);
	lplf->lfWidth = 0;
	lplf->lfCharSet = DEFAULT_CHARSET;
	*reinterpret_cast<HFONT *>(lParam) = CreateFontIndirectW(lplf);
	return 0;
}

// This font is not cached and must be delete manually
HFONT GetFont(_In_z_ LPCWSTR szFont)
{
	HFONT hFont = nullptr;
	LOGFONTW lfFont = { 0 };
	auto hRes = S_OK;
	WC_H(StringCchCopyW(lfFont.lfFaceName, _countof(lfFont.lfFaceName), szFont));

	EnumFontFamiliesExW(GetDC(nullptr), &lfFont, reinterpret_cast<FONTENUMPROCW>(EnumFontFamExProcW), reinterpret_cast<LPARAM>(&hFont), 0);
	if (hFont) return hFont;

	// If we can't get our font, fallback to a system font
	NONCLIENTMETRICSW ncm = { 0 };
	ncm.cbSize = sizeof ncm;
	SystemParametersInfoW(SPI_GETNONCLIENTMETRICS, NULL, &ncm, NULL);
	static auto lf = ncm.lfMessageFont;

	// Create the font, and then return its handle.
	hFont = CreateFontW(lf.lfHeight, lf.lfWidth,
		lf.lfEscapement, lf.lfOrientation, lf.lfWeight,
		lf.lfItalic, lf.lfUnderline, lf.lfStrikeOut, lf.lfCharSet,
		lf.lfOutPrecision, lf.lfClipPrecision, lf.lfQuality,
		lf.lfPitchAndFamily, lf.lfFaceName);
	return hFont;
}

// Cached for deletion in UninitializeGDI
HFONT GetSegoeFont()
{
	if (g_hFontSegoe) return g_hFontSegoe;
	g_hFontSegoe = GetFont(SEGOEW);
	return g_hFontSegoe;
}

// Cached for deletion in UninitializeGDI
HFONT GetSegoeFontBold()
{
	if (g_hFontSegoeBold) return g_hFontSegoeBold;
	g_hFontSegoeBold = GetFont(SEGOEBOLD);
	return g_hFontSegoeBold;
}

_Check_return_ HBRUSH GetSysBrush(uiColor uc)
{
	// Return a cached brush if we have one
	if (g_SysBrushes[uc]) return g_SysBrushes[uc];
	// No cached brush found, cache and return a system brush if requested
	auto iSysColor = g_SysColors[uc];
	if (iSysColor)
	{
		g_SysBrushes[uc] = GetSysColorBrush(iSysColor);
		return g_SysBrushes[uc];
	}

	// No system brush for this color, cache and return a solid brush of the requested color
	auto mc = g_FixedColors[uc];
	if (g_FixedBrushes[mc]) return g_FixedBrushes[mc];
	g_FixedBrushes[mc] = CreateSolidBrush(g_Colors[mc]);
	return g_FixedBrushes[mc];
}

_Check_return_ COLORREF MyGetSysColor(uiColor uc)
{
	// Return a system color if we have one in g_SysColors
	auto iSysColor = g_SysColors[uc];
	if (iSysColor) return ::GetSysColor(iSysColor);

	// No system color listed in g_SysColors, return a hard coded color
	auto mc = g_FixedColors[uc];
	return g_Colors[mc];
}

_Check_return_ HPEN GetPen(uiPen up)
{
	if (g_Pens[up]) return g_Pens[up];
	LOGBRUSH lbr = { 0 };
	lbr.lbStyle = BS_SOLID;

	switch (up)
	{
	case cSolidPen:
	{
		lbr.lbColor = MyGetSysColor(cFrameSelected);
		g_Pens[cSolidPen] = ExtCreatePen(PS_SOLID, 1, &lbr, 0, nullptr);
		return g_Pens[cSolidPen];
	}
	case cSolidGreyPen:
	{
		lbr.lbColor = MyGetSysColor(cFrameUnselected);
		g_Pens[cSolidGreyPen] = ExtCreatePen(PS_SOLID, 1, &lbr, 0, nullptr);
		return g_Pens[cSolidGreyPen];
	}
	case cDashedPen:
	{
		lbr.lbColor = MyGetSysColor(cFrameUnselected);
		DWORD rgStyle[2] = { 1,3 };
		g_Pens[cDashedPen] = ExtCreatePen(PS_GEOMETRIC | PS_USERSTYLE, 1, &lbr, 2, rgStyle);
		return g_Pens[cDashedPen];
	}
	default: break;
	}
	return nullptr;
}

HBITMAP GetBitmap(uiBitmap ub)
{
	if (g_Bitmaps[ub]) return g_Bitmaps[ub];

	g_Bitmaps[ub] = ::LoadBitmap(GetModuleHandle(nullptr), MAKEINTRESOURCE(g_BitmapResources[ub]));
	return g_Bitmaps[ub];
}

void DrawSegoeTextW(
	_In_ HDC hdc,
	_In_z_ LPCWSTR lpchText,
	_In_ COLORREF color,
	_In_ LPRECT lprc,
	bool bBold,
	_In_ UINT format)
{
	auto hfontOld = ::SelectObject(hdc, bBold ? GetSegoeFontBold() : GetSegoeFont());
	auto crText = ::SetTextColor(hdc, color);
	::SetBkMode(hdc, TRANSPARENT);
	::DrawTextW(
		hdc,
		lpchText,
		-1,
		lprc,
		format);
	::SelectObject(hdc, hfontOld);
	(void) ::SetTextColor(hdc, crText);
}

void DrawSegoeTextA(
	_In_ HDC hdc,
	_In_z_ LPCSTR lpchText,
	_In_ COLORREF color,
	_In_ LPRECT lprc,
	bool bBold,
	_In_ UINT format)
{
	auto hfontOld = ::SelectObject(hdc, bBold ? GetSegoeFontBold() : GetSegoeFont());
	auto crText = ::SetTextColor(hdc, color);
	::SetBkMode(hdc, TRANSPARENT);
	::DrawTextA(
		hdc,
		lpchText,
		-1,
		lprc,
		format);
	::SelectObject(hdc, hfontOld);
	(void) ::SetTextColor(hdc, crText);
}

// Clear/initialize formatting on the rich edit control.
// We have to force load the system riched20 to ensure this doesn't break since
// Office's riched20 apparently doesn't handle CFM_COLOR at all.
// Sets our text color, script, and turns off bold, italic, etc formatting.
void ClearEditFormatting(_In_ HWND hWnd, bool bReadOnly)
{
	CHARFORMAT2 cf;
	ZeroMemory(&cf, sizeof cf);
	cf.cbSize = sizeof cf;
	cf.dwMask = CFM_COLOR | CFM_FACE | CFM_BOLD | CFM_ITALIC | CFM_UNDERLINE | CFM_STRIKEOUT;
	cf.crTextColor = MyGetSysColor(bReadOnly ? cTextReadOnly : cText);
	StringCchCopy(cf.szFaceName, _countof(cf.szFaceName), SEGOE);
	(void) ::SendMessage(hWnd, EM_SETCHARFORMAT, SCF_ALL, reinterpret_cast<LPARAM>(&cf));
}

// Lighten the colors of the base, being careful not to overflow
COLORREF LightColor(COLORREF crBase)
{
	auto f = .20;
	auto bRed = static_cast<BYTE>(GetRValue(crBase) + 255 * f);
	auto bGreen = static_cast<BYTE>(GetGValue(crBase) + 255 * f);
	auto bBlue = static_cast<BYTE>(GetBValue(crBase) + 255 * f);
	if (bRed < GetRValue(crBase)) bRed = 0xff;
	if (bGreen < GetGValue(crBase)) bGreen = 0xff;
	if (bBlue < GetBValue(crBase)) bBlue = 0xff;
	return RGB(bRed, bGreen, bBlue);
}

void GradientFillRect(_In_ HDC hdc, RECT rc, uiColor uc)
{
	// Gradient fill the background
	auto crGlow = MyGetSysColor(uc);
	auto crLightGlow = LightColor(crGlow);

	TRIVERTEX vertex[2] = { 0 };
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
	GRADIENT_RECT gRect = { 0 };
	gRect.UpperLeft = 0;
	gRect.LowerRight = 1;
	::GradientFill(hdc, vertex, 2, &gRect, 1, GRADIENT_FILL_RECT_V);
}

void DrawFilledPolygon(_In_ HDC hdc, _In_count_(cpt) CONST POINT *apt, _In_ int cpt, COLORREF cEdge, _In_ HBRUSH hFill)
{
	auto hPen = ::CreatePen(PS_SOLID, 0, cEdge);
	auto hBrushOld = ::SelectObject(hdc, hFill);
	auto hPenOld = ::SelectObject(hdc, hPen);
	::Polygon(hdc, apt, cpt);
	::SelectObject(hdc, hPenOld);
	::SelectObject(hdc, hBrushOld);
	::DeleteObject(hPen);
}

// Draw the frame of our edit controls
LRESULT CALLBACK DrawEditProc(
	_In_ HWND hWnd,
	UINT uMsg,
	WPARAM wParam,
	LPARAM lParam,
	UINT_PTR uIdSubclass,
	DWORD_PTR /*dwRefData*/)
{
	switch (uMsg)
	{
	case WM_NCDESTROY:
		RemoveWindowSubclass(hWnd, DrawEditProc, uIdSubclass);
		return DefSubclassProc(hWnd, uMsg, wParam, lParam);
	case WM_NCPAINT:
	{
		auto hdc = ::GetWindowDC(hWnd);
		if (hdc)
		{
			RECT rc = { 0 };
			::GetWindowRect(hWnd, &rc);
			::OffsetRect(&rc, -rc.left, -rc.top);
			::FrameRect(hdc, &rc, GetSysBrush(cFrameSelected));
			::ReleaseDC(hWnd, hdc);
		}

		// Let the system paint the scroll bar if we have one
		// Be sure to use window coordinates
		auto ws = GetWindowLongPtr(hWnd, GWL_STYLE);
		if (ws & WS_VSCROLL)
		{
			RECT rcScroll = { 0 };
			::GetWindowRect(hWnd, &rcScroll);
			::InflateRect(&rcScroll, -1, -1);
			rcScroll.left = rcScroll.right - GetSystemMetrics(SM_CXHSCROLL);
			auto hRgnCaption = ::CreateRectRgnIndirect(&rcScroll);
			::DefWindowProc(hWnd, WM_NCPAINT, reinterpret_cast<WPARAM>(hRgnCaption), NULL);
			DeleteObject(hRgnCaption);
		}
		return 0;
	}
	}
	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
}

void SubclassEdit(_In_ HWND hWnd, _In_ HWND hWndParent, bool bReadOnly)
{
	SetWindowSubclass(hWnd, DrawEditProc, 0, 0);

	auto lStyle = ::GetWindowLongPtr(hWnd, GWL_EXSTYLE);
	lStyle &= ~WS_EX_CLIENTEDGE;
	(void) ::SetWindowLongPtr(hWnd, GWL_EXSTYLE, lStyle);
	if (bReadOnly) {
		(void) ::SendMessage(hWnd, EM_SETBKGNDCOLOR, static_cast<WPARAM>(0), static_cast<LPARAM>(MyGetSysColor(cBackgroundReadOnly)));
		(void) ::SendMessage(hWnd, EM_SETREADONLY, true, 0L);
	}
	else
	{
		(void) ::SendMessage(hWnd, EM_SETBKGNDCOLOR, static_cast<WPARAM>(0), static_cast<LPARAM>(MyGetSysColor(cBackground)));
	}

	ClearEditFormatting(hWnd, bReadOnly);

	// Set up callback to control paste and context menus
	auto reCallback = new CRichEditOleCallback(hWnd, hWndParent);
	if (reCallback)
	{
		(void) ::SendMessage(
			hWnd,
			EM_SETOLECALLBACK,
			static_cast<WPARAM>(0),
			reinterpret_cast<LPARAM>(reCallback));
		reCallback->Release();
	}
}

void CustomDrawList(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult, int iItemCurHover)
{
	static auto bSelected = false;
	if (!pNMHDR) return;
	auto lvcd = reinterpret_cast<LPNMLVCUSTOMDRAW>(pNMHDR);
	auto iItem = static_cast<int>(lvcd->nmcd.dwItemSpec);

	// If there's nothing to paint, this is a "fake paint" and we don't want to toggle selection highlight
	// Toggling selection highlight causes a repaint, so this logic prevents flicker
	auto bFakePaint = lvcd->nmcd.rc.bottom == 0 &&
		lvcd->nmcd.rc.top == 0 &&
		lvcd->nmcd.rc.left == 0 &&
		lvcd->nmcd.rc.right == 0;

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

		RECT rc = { 0 };
		::GetClientRect(lvcd->nmcd.hdr.hwndFrom, &rc);
		::IntersectRect(&rc, &rc, &lvcd->nmcd.rc);

		if (bResetBottom)
		{
			lvcd->nmcd.rc.bottom = 0;
		}

		if (::IsRectEmpty(&rc))
		{
			*pResult = CDRF_SKIPDEFAULT;
			break;
		}
	}

	// Turn on listview hover highlight
	if (bSelected)
	{
		lvcd->clrText = MyGetSysColor(cGlowText);
		lvcd->clrTextBk = MyGetSysColor(cSelectedBackground);
	}
	else if (iItemCurHover == iItem)
	{
		lvcd->clrText = MyGetSysColor(cGlowText);
		lvcd->clrTextBk = MyGetSysColor(cGlowBackground);
	}
	*pResult = CDRF_NEWFONT;
	break;

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
void DrawListItemGlow(_In_ HWND hWnd, UINT itemID)
{
	if (-1 == itemID) return;

	// If the item already has the selection glow, we don't need to redraw
	auto bSelected = ListView_GetItemState(hWnd, itemID, LVIS_SELECTED) != 0;
	if (bSelected) return;

	RECT rcIcon = { 0 };
	RECT rcLabels = { 0 };
	ListView_GetItemRect(hWnd, itemID, &rcLabels, LVIR_BOUNDS);
	ListView_GetItemRect(hWnd, itemID, &rcIcon, LVIR_ICON);
	rcLabels.left = rcIcon.right;
	if (rcLabels.left >= rcLabels.right) return;
	RECT rcClient = { 0 };
	::GetClientRect(hWnd, &rcClient); // Get our client size
	::IntersectRect(&rcLabels, &rcLabels, &rcClient);
	RedrawWindow(hWnd, &rcLabels, nullptr, RDW_INVALIDATE | RDW_UPDATENOW);
}

void DrawTreeItemGlow(_In_ HWND hWnd, _In_ HTREEITEM hItem)
{
	RECT rect = { 0 };
	RECT rectTree = { 0 };
	TreeView_GetItemRect(hWnd, hItem, &rect, false);
	::GetClientRect(hWnd, &rectTree);
	rect.left = rectTree.left;
	rect.right = rectTree.right;
	::InvalidateRect(hWnd, &rect, false);
}

// Copies ibmWidth x ibmHeight rect from hdcSource to a iWidth x iHeight rect in hdcTarget, replacing colors
// No scaling is performed
void CopyBitmap(HDC hdcSource, HDC hdcTarget, int iWidth, int iHeight, int ibmWidth, int ibmHeight, uiColor cSource, uiColor cReplace)
{
	RECT rcBM = { 0, 0, iWidth, iHeight };

	auto hbmTarget = CreateCompatibleBitmap(
		hdcSource,
		iWidth,
		iHeight);
	(void) ::SelectObject(hdcTarget, hbmTarget);
	::FillRect(hdcTarget, &rcBM, GetSysBrush(cReplace));

	(void)TransparentBlt(
		hdcTarget,
		0,
		0,
		ibmWidth,
		ibmHeight,
		hdcSource,
		0,
		0,
		ibmWidth,
		ibmHeight,
		MyGetSysColor(cSource));
	if (hbmTarget) ::DeleteObject(hbmTarget);
}

// Draws a bitmap on the screen, double buffered, with two color replacement
// Fills rectangle with cBackground
// Replaces cBitmapTransFore (cyan) with cFrameSelected
// Replaces cBitmapTransBack (magenta) with the cBackground
void DrawBitmap(_In_ HDC hdc, _In_ LPRECT rcTarget, uiBitmap iBitmap, bool bHover)
{
	if (!rcTarget) return;
	int iWidth = rcTarget->right - rcTarget->left;
	int iHeight = rcTarget->bottom - rcTarget->top;

	// hdcBitmap: Load the image
	auto hdcBitmap = ::CreateCompatibleDC(hdc);
	auto hbmBitmap = GetBitmap(iBitmap);
	(void) ::SelectObject(hdcBitmap, hbmBitmap);

	BITMAP bm = { 0 };
	::GetObject(hbmBitmap, sizeof bm, &bm);

	// hdcForeReplace: Create a bitmap compatible with hdc, select it, fill with cFrameSelected, copy from hdcBitmap, with cBitmapTransFore transparent
	auto hdcForeReplace = ::CreateCompatibleDC(hdc);
	CopyBitmap(hdcBitmap, hdcForeReplace, iWidth, iHeight, bm.bmWidth, bm.bmHeight, cBitmapTransFore, cFrameSelected);

	// hdcBackReplace: Create a bitmap compatible with hdc, select it, fill with cBackground, copy from hdcForeReplace, with cBitmapTransBack transparent
	auto hdcBackReplace = ::CreateCompatibleDC(hdc);
	CopyBitmap(hdcForeReplace, hdcBackReplace, iWidth, iHeight, bm.bmWidth, bm.bmHeight, cBitmapTransBack, bHover ? cGlowBackground : cBackground);

	// hdc: BitBlt from hdcBackReplace
	(void)BitBlt(
		hdc,
		rcTarget->left,
		rcTarget->top,
		iWidth,
		iHeight,
		hdcBackReplace,
		0,
		0,
		SRCCOPY);

	if (hdcBackReplace) ::DeleteDC(hdcBackReplace);
	if (hdcForeReplace) ::DeleteDC(hdcForeReplace);
	if (hdcBitmap) ::DeleteDC(hdcBitmap);
}

void CustomDrawTree(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult, bool bHover, _In_ HTREEITEM hItemCurHover)
{
	if (!pNMHDR) return;
	auto lvcd = reinterpret_cast<LPNMTVCUSTOMDRAW>(pNMHDR);

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
			lvcd->clrTextBk = MyGetSysColor(cBackground);
			lvcd->clrText = MyGetSysColor(cText);
			int iState = TreeView_GetItemState(lvcd->nmcd.hdr.hwndFrom, hItem, TVIS_SELECTED);
			TreeView_SetItemState(lvcd->nmcd.hdr.hwndFrom, hItem, iState & TVIS_SELECTED ? TVIS_BOLD : NULL, TVIS_BOLD);
		}

		*pResult = CDRF_DODEFAULT | CDRF_NOTIFYPOSTPAINT;

		if (hItem == hItemCurHover)
		{
			lvcd->clrText = MyGetSysColor(cGlowText);
			lvcd->clrTextBk = MyGetSysColor(cGlowBackground);
			*pResult |= CDRF_NEWFONT;
		}
		break;
	}

	case CDDS_ITEMPOSTPAINT:
	{
		// If we've advised on this object, add a little icon to let the user know
		auto hItem = reinterpret_cast<HTREEITEM>(lvcd->nmcd.dwItemSpec);

		if (hItem)
		{
			// Cover over the +/- and paint triangles instead
			DrawExpandTriangle(lvcd->nmcd.hdr.hwndFrom, lvcd->nmcd.hdc, hItem, bHover && hItem == hItemCurHover, hItem == hItemCurHover);

			TVITEM tvi = { 0 };
			tvi.mask = TVIF_PARAM;
			tvi.hItem = hItem;
			TreeView_GetItem(lvcd->nmcd.hdr.hwndFrom, &tvi);
			auto lpData = reinterpret_cast<SortListData*>(tvi.lParam);
			if (lpData && lpData->Node() && lpData->Node()->m_lpAdviseSink)
			{
				RECT rect = { 0 };
				TreeView_GetItemRect(lvcd->nmcd.hdr.hwndFrom, hItem, &rect, 1);
				rect.left = rect.right;
				rect.right += rect.bottom - rect.top;
				DrawBitmap(lvcd->nmcd.hdc, &rect, cNotify, hItem == hItemCurHover);
			}
		}
		break;
	}
	default:
		*pResult = CDRF_DODEFAULT;
		break;
	}
}

// Paints the triangles indicating expansion state
void DrawExpandTriangle(_In_ HWND hWnd, _In_ HDC hdc, _In_ HTREEITEM hItem, bool bGlow, bool bHover)
{
	TVITEM tvitem = { 0 };
	tvitem.hItem = hItem;
	tvitem.mask = TVIF_CHILDREN | TVIF_STATE;
	TreeView_GetItem(hWnd, &tvitem);
	auto bHasChildren = tvitem.cChildren != 0;
	if (bHasChildren)
	{
		RECT rcButton = { 0 };
		TreeView_GetItemRect(hWnd, hItem, &rcButton, true);

		// Erase the +/- icons
		// We erase everything to the left of the label
		rcButton.right = rcButton.left;
		rcButton.left = 0;
		::FillRect(hdc, &rcButton, GetSysBrush(bHover ? cGlowBackground : cBackground));

		// Now we focus on a box 15 pixels wide to the left of the label
		rcButton.left = rcButton.right - 15;

		// Boundary box for the actual triangles
		RECT rcTriangle = { 0 };

		POINT tri[3] = { 0 };
		auto bExpanded = TVIS_EXPANDED == (tvitem.state & TVIS_EXPANDED);
		if (bExpanded)
		{
			rcTriangle.top = (rcButton.top + rcButton.bottom) / 2 - 3;
			rcTriangle.bottom = rcTriangle.top + 5;
			rcTriangle.left = rcButton.left;
			rcTriangle.right = rcTriangle.left + 5;

			tri[0].x = rcTriangle.left;
			tri[0].y = rcTriangle.bottom;
			tri[1].x = rcTriangle.right;
			tri[1].y = rcTriangle.top;
			tri[2].x = rcTriangle.right;
			tri[2].y = rcTriangle.bottom;
		}
		else
		{
			rcTriangle.top = (rcButton.top + rcButton.bottom) / 2 - 4;
			rcTriangle.bottom = rcTriangle.top + 8;
			rcTriangle.left = rcButton.left;
			rcTriangle.right = rcTriangle.left + 4;

			tri[0].x = rcTriangle.left;
			tri[0].y = rcTriangle.top;
			tri[1].x = rcTriangle.left;
			tri[1].y = rcTriangle.bottom;
			tri[2].x = rcTriangle.right;
			tri[2].y = (rcTriangle.top + rcTriangle.bottom) / 2;
		}

		uiColor cEdge;
		uiColor cFill;
		if (bGlow)
		{
			cEdge = cGlow;
			cFill = cGlow;
		}
		else
		{
			cEdge = bExpanded ? cFrameSelected : cArrow;
			cFill = bExpanded ? cFrameSelected : cBackground;
		}
		DrawFilledPolygon(hdc, tri, _countof(tri), MyGetSysColor(cEdge), GetSysBrush(cFill));
	}
}

void DrawTriangle(_In_ HWND hWnd, _In_ HDC hdc, _In_ CONST RECT* lprc, bool bButton, bool bUp)
{
	if (!hWnd || !hdc || !lprc) return;

	CDoubleBuffer db;
	db.Begin(hdc, lprc);

	::FillRect(hdc, lprc, GetSysBrush(bButton ? cPaneHeaderBackground : cBackground));

	POINT tri[3] = { 0 };
	LONG lCenter = 0;
	LONG lTop = 0;
	if (bButton)
	{
		lCenter = lprc->left + (lprc->bottom - lprc->top) / 2;
		lTop = (lprc->top + lprc->bottom - TRIANGLE_SIZE) / 2;
	}
	else
	{
		lCenter = (lprc->left + lprc->right) / 2;
		lTop = lprc->top + GetSystemMetrics(SM_CYBORDER);
	}

	if (bUp)
	{
		tri[0].x = lCenter;
		tri[0].y = lTop;
		tri[1].x = lCenter - TRIANGLE_SIZE;
		tri[1].y = lTop + TRIANGLE_SIZE;
		tri[2].x = lCenter + TRIANGLE_SIZE;
		tri[2].y = lTop + TRIANGLE_SIZE;
	}
	else
	{
		tri[0].x = lCenter;
		tri[0].y = lTop + TRIANGLE_SIZE;
		tri[1].x = lCenter - TRIANGLE_SIZE;
		tri[1].y = lTop;
		tri[2].x = lCenter + TRIANGLE_SIZE;
		tri[2].y = lTop;
	}
	auto uiArrow = cArrow;
	if (bButton) uiArrow = cPaneHeaderText;
	DrawFilledPolygon(hdc, tri, _countof(tri), MyGetSysColor(uiArrow), GetSysBrush(uiArrow));

	db.End(hdc);
}

void DrawHeaderItem(_In_ HWND hWnd, _In_ HDC hdc, UINT itemID, _In_ LPRECT lprc)
{
	auto bSorted = false;
	auto bSortUp = true;

	WCHAR szHeader[255] = { 0 };
	HDITEMW hdItem = { 0 };
	hdItem.mask = HDI_TEXT | HDI_FORMAT;
	hdItem.pszText = szHeader;
	hdItem.cchTextMax = _countof(szHeader);
	::SendMessage(hWnd, HDM_GETITEMW, static_cast<WPARAM>(itemID), reinterpret_cast<LPARAM>(&hdItem));

	if (hdItem.fmt & (HDF_SORTUP | HDF_SORTDOWN))
		bSorted = true;
	if (bSorted)
	{
		bSortUp = (hdItem.fmt & HDF_SORTUP) != 0;
	}

	CDoubleBuffer db;
	db.Begin(hdc, lprc);

	auto rcHeader = *lprc;
	::FillRect(hdc, &rcHeader, GetSysBrush(cBackground));

	if (bSorted)
	{
		DrawTriangle(hWnd, hdc, lprc, false, bSortUp);
	}

	auto rcText = rcHeader;
	rcText.left += GetSystemMetrics(SM_CXEDGE);
	DrawSegoeTextW(
		hdc,
		hdItem.pszText,
		MyGetSysColor(cText),
		&rcText,
		false,
		DT_END_ELLIPSIS | DT_SINGLELINE | DT_VCENTER);

	// Draw a line under for some visual separation
	auto hpenOld = ::SelectObject(hdc, GetPen(cSolidGreyPen));
	::MoveToEx(hdc, rcHeader.left, rcHeader.bottom - 1, nullptr);
	::LineTo(hdc, rcHeader.right, rcHeader.bottom - 1);
	(void) ::SelectObject(hdc, hpenOld);

	// Draw our divider
	// Since no one else uses rcHeader after here, we can modify it in place
	::InflateRect(&rcHeader, 0, -1);
	rcHeader.left = rcHeader.right - 2;
	rcHeader.bottom -= 1;
	::FrameRect(hdc, &rcHeader, GetSysBrush(cFrameUnselected));

	db.End(hdc);
}

// Draw the unused portion of the header
void CustomDrawHeader(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
{
	if (!pNMHDR) return;
	auto lvcd = reinterpret_cast<LPNMCUSTOMDRAW>(pNMHDR);

	switch (lvcd->dwDrawStage)
	{
	case CDDS_PREPAINT:
		*pResult = CDRF_NOTIFYPOSTPAINT | CDRF_NOTIFYITEMDRAW;
		break;

	case CDDS_ITEMPREPAINT:
		DrawHeaderItem(lvcd->hdr.hwndFrom, lvcd->hdc, static_cast<UINT>(lvcd->dwItemSpec), &lvcd->rc);
		*pResult = CDRF_SKIPDEFAULT;
		break;

	case CDDS_POSTPAINT:
	{
		RECT rc = { 0 };
		// Get the rc for the entire header
		::GetClientRect(lvcd->hdr.hwndFrom, &rc);
		// Get the last item in the header
		auto iCount = Header_GetItemCount(lvcd->hdr.hwndFrom);
		if (iCount)
		{
			// Find the index of the last item in the header
			auto iIndex = Header_OrderToIndex(lvcd->hdr.hwndFrom, iCount - 1);

			RECT rcRight = { 0 };
			// Compute the right edge of the last item
			(void)Header_GetItemRect(lvcd->hdr.hwndFrom, iIndex, &rcRight);
			rc.left = rcRight.right;
		}

		// If we have visible non occupied header, paint it
		if (rc.left < rc.right)
		{
			::FillRect(lvcd->hdc, &rc, GetSysBrush(cBackground));
		}

		*pResult = CDRF_DODEFAULT;
		break;
	}

	default:
		*pResult = CDRF_DODEFAULT;
		break;
	}
}

void DrawTrackingBar(_In_ HWND hWndHeader, _In_ HWND hWndList, int x, int iHeaderHeight, bool bRedraw)
{
	RECT rcTracker = { 0 };
	::GetClientRect(hWndList, &rcTracker);
	auto hdc = ::GetDC(hWndList);
	rcTracker.top += iHeaderHeight;
	rcTracker.left = x - 1; // this lines us up under the splitter line we drew in the header
	rcTracker.right = x;
	::MapWindowPoints(hWndHeader, hWndList, reinterpret_cast<LPPOINT>(&rcTracker), 2);
	if (bRedraw)
	{
		::InvalidateRect(hWndList, &rcTracker, true);
	}
	else
	{
		::FillRect(hdc, &rcTracker, GetSysBrush(cFrameSelected));
	}
	::ReleaseDC(hWndList, hdc);
}

void StyleButton(_In_ HWND hWnd, uiButtonStyle bsStyle)
{
	auto hRes = S_OK;

	EC_B(::SetProp(hWnd, BUTTON_STYLE, reinterpret_cast<HANDLE>(bsStyle)));
}

void DrawButton(_In_ HWND hWnd, _In_ HDC hDC, _In_ LPRECT lprc, UINT itemState)
{
	::FillRect(hDC, lprc, GetSysBrush(cBackground));

	auto bsStyle = reinterpret_cast<intptr_t>(::GetProp(hWnd, BUTTON_STYLE));
	switch (bsStyle)
	{
	case bsUnstyled:
	{
		WCHAR szButton[255] = { 0 };
		::GetWindowTextW(hWnd, szButton, _countof(szButton));
		auto iState = static_cast<int>(::SendMessage(hWnd, BM_GETSTATE, NULL, NULL));
		auto bGlow = (iState & BST_HOT) != 0;
		auto bPushed = (iState & BST_PUSHED) != 0;
		auto bDisabled = (itemState & CDIS_DISABLED) != 0;
		auto bFocused = (itemState & CDIS_FOCUS) != 0;

		::FrameRect(hDC, lprc, bFocused || bGlow || bPushed ? GetSysBrush(cFrameSelected) : GetSysBrush(cFrameUnselected));

		DrawSegoeTextW(
			hDC,
			szButton,
			bPushed || bDisabled ? MyGetSysColor(cTextDisabled) : MyGetSysColor(cText),
			lprc,
			false,
			DT_SINGLELINE | DT_VCENTER | DT_CENTER);
	}
	break;
	case bsUpArrow:
		DrawTriangle(hWnd, hDC, lprc, true, true);
		break;
	case bsDownArrow:
		DrawTriangle(hWnd, hDC, lprc, true, false);
		break;
	}
}

void DrawCheckButton(_In_ HWND hWnd, _In_ HDC hDC, _In_ LPRECT lprc, UINT itemState)
{
	WCHAR szButton[255];
	::GetWindowTextW(hWnd, szButton, _countof(szButton));
	auto iState = static_cast<int>(::SendMessage(hWnd, BM_GETSTATE, NULL, NULL));
	auto bGlow = BST_HOT == (iState & BST_HOT);
	auto bChecked = (iState & BST_CHECKED) != 0;
	auto bDisabled = (itemState & CDIS_DISABLED) != 0;
	auto bFocused = (itemState & CDIS_FOCUS) != 0;

	long lSpacing = GetSystemMetrics(SM_CYEDGE);
	auto lCheck = lprc->bottom - lprc->top - 2 * lSpacing;
	RECT rcCheck = { 0 };
	rcCheck.right = rcCheck.left + lCheck;
	rcCheck.top = lprc->top + lSpacing;
	rcCheck.bottom = lprc->bottom - lSpacing;

	::FillRect(hDC, lprc, GetSysBrush(cBackground));
	::FrameRect(hDC, &rcCheck, GetSysBrush(bDisabled ? cFrameUnselected : bGlow || bFocused ? cGlow : cFrameSelected));
	if (bChecked)
	{
		auto rcFill = rcCheck;
		::InflateRect(&rcFill, -3, -3);
		::FillRect(hDC, &rcFill, GetSysBrush(cGlow));
	}

	lprc->left = rcCheck.right + lSpacing;

	DrawSegoeTextW(
		hDC,
		szButton,
		bDisabled ? MyGetSysColor(cTextDisabled) : MyGetSysColor(cText),
		lprc,
		false,
		DT_SINGLELINE | DT_VCENTER);
}

bool CustomDrawButton(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
{
	if (!pNMHDR) return false;
	auto lvcd = reinterpret_cast<LPNMCUSTOMDRAW>(pNMHDR);

	// Ensure we only handle painting for buttons
	WCHAR szClass[64] = { 0 };
	::GetClassNameW(lvcd->hdr.hwndFrom, szClass, _countof(szClass));
	if (_wcsicmp(szClass, L"BUTTON")) return false;

	switch (lvcd->dwDrawStage)
	{
	case CDDS_PREPAINT:
	{
		auto lStyle = BS_TYPEMASK & ::GetWindowLongA(lvcd->hdr.hwndFrom, GWL_STYLE);

		*pResult = CDRF_SKIPDEFAULT;
		if (BS_AUTOCHECKBOX == lStyle)
		{
			DrawCheckButton(lvcd->hdr.hwndFrom, lvcd->hdc, &lvcd->rc, lvcd->uItemState);
		}
		else if (BS_PUSHBUTTON == lStyle ||
			BS_DEFPUSHBUTTON == lStyle)
		{
			DrawButton(lvcd->hdr.hwndFrom, lvcd->hdc, &lvcd->rc, lvcd->uItemState);
		}
		else *pResult = CDRF_DODEFAULT;

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
	auto lpMenuEntry = reinterpret_cast<LPMENUENTRY>(lpMeasureItemStruct->itemData);
	if (!lpMenuEntry) return;

	auto hdc = ::GetDC(nullptr);
	if (hdc)
	{
		auto hfontOld = ::SelectObject(hdc, GetSegoeFont());

		// In order to compute the right width, we need to drop our prefix characters
		auto szText = StripCharacter(lpMenuEntry->m_MSAA.pszWText, L'&');

		auto size = GetTextExtentPoint32(hdc, szText);
		lpMeasureItemStruct->itemWidth = size.cx + 2 * GetSystemMetrics(SM_CXEDGE);
		lpMeasureItemStruct->itemHeight = size.cy + 2 * GetSystemMetrics(SM_CYEDGE);

		// Make sure we have room for the flyout icon
		if (!lpMenuEntry->b_OnMenuBar && ::IsMenu(reinterpret_cast<HMENU>(static_cast<UINT_PTR>(lpMeasureItemStruct->itemID))))
		{
			lpMeasureItemStruct->itemWidth += GetSystemMetrics(SM_CXSMICON);
		}

		::SelectObject(hdc, hfontOld);
		::ReleaseDC(nullptr, hdc);
	}
}

void MeasureItem(_In_ LPMEASUREITEMSTRUCT lpMeasureItemStruct)
{
	if (ODT_MENU == lpMeasureItemStruct->CtlType)
	{
		MeasureMenu(lpMeasureItemStruct);
	}
	// Important to keep this case even if we do not use it - CDialog::OnMeasureItem asserts.
	else if (ODT_COMBOBOX == lpMeasureItemStruct->CtlType)
	{
		// DebugPrint(DBGGeneric,"Combo Box\n");
	}
}

void DrawMenu(_In_ LPDRAWITEMSTRUCT lpDrawItemStruct)
{
	if (!lpDrawItemStruct) return;
	auto lpMenuEntry = reinterpret_cast<LPMENUENTRY>(lpDrawItemStruct->itemData);
	if (!lpMenuEntry) return;

	auto bSeparator = !lpMenuEntry->m_MSAA.cchWText; // Cheap way to find separators. Revisit if odd effects show.
	auto bAccel = (lpDrawItemStruct->itemState & ODS_NOACCEL) != 0;
	auto bHot = (lpDrawItemStruct->itemState & (ODS_HOTLIGHT | ODS_SELECTED)) != 0;
	auto bDisabled = (lpDrawItemStruct->itemState & (ODS_GRAYED | ODS_DISABLED)) != 0;

	// Double buffer our menu painting
	CDoubleBuffer db;
	auto hdc = lpDrawItemStruct->hDC;
	db.Begin(hdc, &lpDrawItemStruct->rcItem);

	// Draw background
	auto rcItem = lpDrawItemStruct->rcItem;
	auto rcText = rcItem;

	if (!lpMenuEntry->b_OnMenuBar)
	{
		auto rectGutter = rcText;
		rcText.left += GetSystemMetrics(SM_CXMENUCHECK);
		rectGutter.right = rcText.left;
		::FillRect(hdc, &rectGutter, GetSysBrush(cBackgroundReadOnly));
	}

	auto cBack = cBackground;
	auto cFore = cText;
	if (bHot && !bDisabled)
	{
		cBack = cSelectedBackground;
		cFore = cGlowText;
		::FillRect(hdc, &rcItem, GetSysBrush(cBack));
	}
	else
	{
		if (bDisabled) cFore = cTextDisabled;
		::FillRect(hdc, &rcText, GetSysBrush(cBack));
	}

	if (bSeparator)
	{
		::InflateRect(&rcText, -3, 0);
		auto lMid = (rcText.bottom + rcText.top) / 2;
		auto hpenOld = ::SelectObject(hdc, GetPen(cSolidGreyPen));
		::MoveToEx(hdc, rcText.left, lMid, nullptr);
		::LineTo(hdc, rcText.right, lMid);
		(void) ::SelectObject(hdc, hpenOld);
	}
	else if (lpMenuEntry->m_pName)
	{
		UINT uiTextFlags = DT_SINGLELINE | DT_VCENTER;
		if (bAccel) uiTextFlags |= DT_HIDEPREFIX;

		if (lpMenuEntry->b_OnMenuBar)
			uiTextFlags |= DT_CENTER;
		else
			rcText.left += GetSystemMetrics(SM_CXEDGE);

		DrawSegoeTextW(
			hdc,
			lpMenuEntry->m_pName,
			MyGetSysColor(cFore),
			&rcText,
			false,
			uiTextFlags);

		// Triple buffer the check mark so we can copy it over without the background
		if (lpDrawItemStruct->itemState & ODS_CHECKED)
		{
			UINT nWidth = GetSystemMetrics(SM_CXMENUCHECK);
			UINT nHeight = GetSystemMetrics(SM_CYMENUCHECK);
			RECT rc = { 0 };
			auto bm = CreateBitmap(nWidth, nHeight, 1, 1, nullptr);
			auto hdcMem = CreateCompatibleDC(hdc);

			SelectObject(hdcMem, bm);
			SetRect(&rc, 0, 0, nWidth, nHeight);
			(void)DrawFrameControl(hdcMem, &rc, DFC_MENU, DFCS_MENUCHECK);

			(void)TransparentBlt(
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
				MyGetSysColor(cBackground));

			DeleteDC(hdcMem);
			DeleteObject(bm);
		}
	}

	db.End(hdc);
}

void DrawComboBox(_In_ LPDRAWITEMSTRUCT lpDrawItemStruct)
{
	if (!lpDrawItemStruct) return;
	if (lpDrawItemStruct->itemID == -1) return;

	TCHAR szText[128] = { 0 };
	// Get and display the text for the list item.
	::SendMessage(lpDrawItemStruct->hwndItem, CB_GETLBTEXT, lpDrawItemStruct->itemID, reinterpret_cast<LPARAM>(szText));
	auto bHot = 0 != (lpDrawItemStruct->itemState & (ODS_FOCUS | ODS_SELECTED));
	auto cBack = cBackground;
	auto cFore = cText;
	if (bHot)
	{
		cBack = cGlowBackground;
		cFore = cGlowText;
	}

	::FillRect(lpDrawItemStruct->hDC, &lpDrawItemStruct->rcItem, GetSysBrush(cBack));

	DrawSegoeText(
		lpDrawItemStruct->hDC,
		szText,
		MyGetSysColor(cFore),
		&lpDrawItemStruct->rcItem,
		false,
		DT_LEFT | DT_SINGLELINE | DT_VCENTER);
}

void DrawItem(_In_ LPDRAWITEMSTRUCT lpDrawItemStruct)
{
	if (!lpDrawItemStruct) return;

	if (ODT_MENU == lpDrawItemStruct->CtlType)
	{
		DrawMenu(lpDrawItemStruct);
	}
	else if (ODT_COMBOBOX == lpDrawItemStruct->CtlType)
	{
		DrawComboBox(lpDrawItemStruct);
	}
}

// Paint the status bar with double buffering to avoid flicker
void DrawStatus(
	HWND hwnd,
	int iStatusHeight,
	wstring& szStatusData1,
	int iStatusData1,
	wstring& szStatusData2,
	int iStatusData2,
	wstring& szStatusInfo)
{
	RECT rcStatus = { 0 };
	::GetClientRect(hwnd, &rcStatus);
	if (rcStatus.bottom - rcStatus.top > iStatusHeight)
	{
		rcStatus.top = rcStatus.bottom - iStatusHeight;
	}
	PAINTSTRUCT ps = { nullptr };
	::BeginPaint(hwnd, &ps);

	CDoubleBuffer db;
	db.Begin(ps.hdc, &rcStatus);

	auto rcGrad = rcStatus;
	auto rcText = rcGrad;
	auto crFore = MyGetSysColor(cStatusText);

	// We start painting a little lower to allow our gradiants to line up
	rcGrad.bottom += GetSystemMetrics(SM_CYSIZEFRAME) - BORDER_VISIBLEWIDTH;
	GradientFillRect(ps.hdc, rcGrad, cStatus);

	rcText.left = rcText.right - iStatusData2;
	::DrawSegoeTextW(
		ps.hdc,
		szStatusData2.c_str(),
		crFore,
		&rcText,
		true,
		DT_LEFT | DT_SINGLELINE | DT_BOTTOM);
	rcText.right = rcText.left;
	rcText.left = rcText.right - iStatusData1;
	::DrawSegoeTextW(
		ps.hdc,
		szStatusData1.c_str(),
		crFore,
		&rcText,
		true,
		DT_LEFT | DT_SINGLELINE | DT_BOTTOM);
	rcText.right = rcText.left;
	rcText.left = 0;
	::DrawSegoeTextW(
		ps.hdc,
		szStatusInfo.c_str(),
		crFore,
		&rcText,
		true,
		DT_LEFT | DT_SINGLELINE | DT_BOTTOM | DT_END_ELLIPSIS);

	db.End(ps.hdc);
	::EndPaint(hwnd, &ps);
}

void GetCaptionRects(HWND hWnd,
	RECT* lprcFullCaption,
	RECT* lprcIcon,
	RECT* lprcCloseIcon,
	RECT* lprcMaxIcon,
	RECT* lprcMinIcon,
	RECT* lprcCaptionText)
{
	RECT rcFullCaption = { 0 };
	RECT rcIcon = { 0 };
	RECT rcCloseIcon = { 0 };
	RECT rcMaxIcon = { 0 };
	RECT rcMinIcon = { 0 };
	RECT rcCaptionText = { 0 };

	RECT rcWindow = { 0 };
	auto dwWinStyle = GetWindowStyle(hWnd);
	auto bModal = DS_MODALFRAME == (dwWinStyle & DS_MODALFRAME);
	auto bThickFrame = WS_THICKFRAME == (dwWinStyle & WS_THICKFRAME);
	auto bMinBox = WS_MINIMIZEBOX == (dwWinStyle & WS_MINIMIZEBOX);
	auto bMaxBox = WS_MAXIMIZEBOX == (dwWinStyle & WS_MAXIMIZEBOX);

	::GetWindowRect(hWnd, &rcWindow); // Get our non-client size
	::OffsetRect(&rcWindow, -rcWindow.left, -rcWindow.top); // shift the origin to 0 since that's where our DC paints
	// At this point, we have rectangles for our window and client area
	// rcWindow is the outer rectangle for our NC frame

	auto cxPaddedBorder = GetSystemMetrics(SM_CXPADDEDBORDER);
	auto cxFixedFrame = GetSystemMetrics(SM_CXFIXEDFRAME);
	auto cyFixedFrame = GetSystemMetrics(SM_CYFIXEDFRAME);
	auto cxSizeFrame = GetSystemMetrics(SM_CXSIZEFRAME);
	auto cySizeFrame = GetSystemMetrics(SM_CYSIZEFRAME);
	auto cySize = GetSystemMetrics(SM_CYSIZE);
	auto cxSizeButton = cySize;
	if (pfnGetThemeSysSize) cxSizeButton = pfnGetThemeSysSize(nullptr, SM_CXSIZE);
	auto cxBorder = GetSystemMetrics(SM_CXBORDER);
	auto cyBorder = GetSystemMetrics(SM_CYBORDER);

	auto cxFrame = bThickFrame ? cxSizeFrame : cxFixedFrame;
	auto cyFrame = bThickFrame ? cySizeFrame : cyFixedFrame;

	// If padded borders are in effect, we fall back to a single width for both thick and thin frames
	if (cxPaddedBorder)
	{
		cxFrame = cxSizeFrame + cxPaddedBorder;
		cyFrame = cySizeFrame + cxPaddedBorder;
	}

	rcFullCaption.top = rcWindow.top + BORDER_VISIBLEWIDTH;
	rcFullCaption.left = rcWindow.left + BORDER_VISIBLEWIDTH;
	rcFullCaption.right = rcWindow.right - BORDER_VISIBLEWIDTH;
	rcFullCaption.bottom =
		rcCaptionText.bottom =
		rcWindow.top + cyFrame + GetSystemMetrics(SM_CYCAPTION);

	rcIcon.top = rcWindow.top + cyFrame + GetSystemMetrics(SM_CYEDGE);
	rcIcon.left = rcWindow.left + cxFrame + cxFixedFrame;
	rcIcon.bottom = rcIcon.top + GetSystemMetrics(SM_CYSMICON);
	rcIcon.right = rcIcon.left + GetSystemMetrics(SM_CXSMICON);

	rcCloseIcon.top = rcMaxIcon.top = rcMinIcon.top = rcWindow.top + cyFrame + cyBorder;
	rcCloseIcon.bottom = rcMaxIcon.bottom = rcMinIcon.bottom = rcCloseIcon.top + cySize - 2 * cyBorder;
	rcCloseIcon.right = rcWindow.right - cxFrame - cxBorder;
	rcCloseIcon.left = rcMaxIcon.right = rcCloseIcon.right - cxSizeButton;

	rcMaxIcon.left = rcMaxIcon.right - cxSizeButton;

	rcMinIcon.right = rcMaxIcon.left + GetSystemMetrics(SM_CXEDGE);
	rcMinIcon.left = rcMinIcon.right - cxSizeButton;

	::InflateRect(&rcCloseIcon, -1, -1);
	::InflateRect(&rcMaxIcon, -1, -1);
	::InflateRect(&rcMinIcon, -1, -1);

	rcCaptionText.top = rcWindow.top + cxFrame;
	if (bModal)
	{
		rcCaptionText.left = rcWindow.left + BORDER_VISIBLEWIDTH + cxFixedFrame + cxBorder;
	}
	else
	{
		rcCaptionText.left = rcIcon.right + cxFixedFrame + cxBorder;
	}
	rcCaptionText.right = rcCloseIcon.left;
	if (bMinBox) rcCaptionText.right = rcMinIcon.left;
	else if (bMaxBox) rcCaptionText.right = rcMaxIcon.left;

	if (lprcFullCaption) *lprcFullCaption = rcFullCaption;
	if (lprcIcon) *lprcIcon = rcIcon;
	if (lprcCloseIcon) *lprcCloseIcon = rcCloseIcon;
	if (lprcCloseIcon) *lprcCloseIcon = rcCloseIcon;
	if (lprcMaxIcon) *lprcMaxIcon = rcMaxIcon;
	if (lprcMinIcon) *lprcMinIcon = rcMinIcon;
	if (lprcCaptionText) *lprcCaptionText = rcCaptionText;
}

void DrawSystemButtons(_In_ HWND hWnd, _In_opt_ HDC hdc, int iHitTest)
{
	HDC hdcLocal = nullptr;
	if (!hdc)
	{
		hdcLocal = ::GetWindowDC(hWnd);
		hdc = hdcLocal;
	}

	auto dwWinStyle = GetWindowStyle(hWnd);
	auto bMinBox = WS_MINIMIZEBOX == (dwWinStyle & WS_MINIMIZEBOX);
	auto bMaxBox = WS_MAXIMIZEBOX == (dwWinStyle & WS_MAXIMIZEBOX);

	RECT rcCloseIcon = { 0 };
	RECT rcMaxIcon = { 0 };
	RECT rcMinIcon = { 0 };
	GetCaptionRects(hWnd, nullptr, nullptr, &rcCloseIcon, &rcMaxIcon, &rcMinIcon, nullptr);

	// Draw our system buttons appropriately
	(void) ::OffsetRect(&rcCloseIcon, HTCLOSE == iHitTest ? 1 : 0, HTCLOSE == iHitTest ? 1 : 0);
	DrawBitmap(hdc, &rcCloseIcon, cClose, false);

	if (bMaxBox)
	{
		(void) ::OffsetRect(&rcMaxIcon, HTMAXBUTTON == iHitTest ? 1 : 0, HTMAXBUTTON == iHitTest ? 1 : 0);
		DrawBitmap(hdc, &rcMaxIcon, ::IsZoomed(hWnd) ? cRestore : cMaximize, false);
	}

	if (bMinBox)
	{
		(void) ::OffsetRect(&rcMinIcon, HTMINBUTTON == iHitTest ? 1 : 0, HTMINBUTTON == iHitTest ? 1 : 0);
		DrawBitmap(hdc, &rcMinIcon, cMinimize, false);
	}

	if (hdcLocal) ::ReleaseDC(hWnd, hdcLocal);
}

void DrawWindowFrame(_In_ HWND hWnd, bool bActive, int iStatusHeight)
{
	auto cxPaddedBorder = GetSystemMetrics(SM_CXPADDEDBORDER);
	auto cxFixedFrame = GetSystemMetrics(SM_CXFIXEDFRAME);
	auto cyFixedFrame = GetSystemMetrics(SM_CYFIXEDFRAME);
	auto cxSizeFrame = GetSystemMetrics(SM_CXSIZEFRAME);
	auto cySizeFrame = GetSystemMetrics(SM_CYSIZEFRAME);

	auto dwWinStyle = GetWindowStyle(hWnd);
	auto bModal = DS_MODALFRAME == (dwWinStyle & DS_MODALFRAME);
	auto bThickFrame = WS_THICKFRAME == (dwWinStyle & WS_THICKFRAME);

	RECT rcWindow = { 0 };
	RECT rcClient = { 0 };
	::GetWindowRect(hWnd, &rcWindow); // Get our non-client size
	::GetClientRect(hWnd, &rcClient); // Get our client size
	::MapWindowPoints(hWnd, nullptr, reinterpret_cast<LPPOINT>(&rcClient), 2); // locate our client rect on the screen

	// Before we fiddle with our client and window rects further, paint the menu
	// This must be in window coordinates for WM_NCPAINT to work!!!
	auto rcMenu = rcWindow;
	rcMenu.top = rcMenu.top + cySizeFrame + ::GetSystemMetrics(SM_CYCAPTION);
	rcMenu.bottom = rcClient.top;
	rcMenu.left += cxSizeFrame;
	rcMenu.right -= cxSizeFrame;
	auto hRgnCaption = ::CreateRectRgnIndirect(&rcMenu);
	::DefWindowProc(hWnd, WM_NCPAINT, reinterpret_cast<WPARAM>(hRgnCaption), NULL);
	DeleteObject(hRgnCaption);

	::OffsetRect(&rcClient, -rcWindow.left, -rcWindow.top); // shift the origin to 0 since that's where our DC paints
	::OffsetRect(&rcWindow, -rcWindow.left, -rcWindow.top); // shift the origin to 0 since that's where our DC paints
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
	RECT rcFullCaption = { 0 };
	RECT rcIcon = { 0 };
	RECT rcCaptionText = { 0 };
	RECT rcMenuGutterLeft = { 0 };
	RECT rcMenuGutterRight = { 0 };
	RECT rcWindowGutterLeft = { 0 };
	RECT rcWindowGutterRight = { 0 };

	GetCaptionRects(hWnd, &rcFullCaption, &rcIcon, nullptr, nullptr, nullptr, &rcCaptionText);

	rcMenuGutterLeft.left =
		rcWindowGutterLeft.left =
		rcFullCaption.left;
	rcMenuGutterRight.right =
		rcWindowGutterRight.right =
		rcFullCaption.right;
	rcMenuGutterLeft.top =
		rcMenuGutterRight.top =
		rcFullCaption.bottom;

	rcMenuGutterLeft.right = rcWindow.left + cxFrame;
	rcMenuGutterRight.left = rcWindow.right - cxFrame;
	rcMenuGutterLeft.bottom =
		rcMenuGutterRight.bottom =
		rcWindowGutterLeft.top =
		rcWindowGutterRight.top =
		rcClient.top;

	rcWindowGutterLeft.bottom = rcWindowGutterRight.bottom = rcClient.bottom - iStatusHeight;

	rcWindowGutterLeft.right = rcClient.left;
	rcWindowGutterRight.left = rcClient.right;

	// Get a DC where the upper left corner of the window is 0,0
	// All of our rectangles were computed with this DC in mind
	auto hdcWin = ::GetWindowDC(hWnd);
	if (hdcWin)
	{
		CDoubleBuffer db;
		auto hdc = hdcWin;
		db.Begin(hdc, &rcWindow);

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
		auto hpenOld = ::SelectObject(hdc, GetPen(cSolidGreyPen));
		::MoveToEx(hdc, rcMenuGutterLeft.right, rcClient.top - 1, nullptr);
		::LineTo(hdc, rcMenuGutterRight.left, rcClient.top - 1);
		(void) ::SelectObject(hdc, hpenOld);

		// White out the caption
		::FillRect(hdc, &rcFullCaption, GetSysBrush(cBackground));

		DrawSystemButtons(hWnd, hdc, 0);

		// Draw our icon
		if (!bModal)
		{
			auto hIcon = static_cast<HICON>(::LoadImage(
				AfxGetInstanceHandle(),
				MAKEINTRESOURCE(IDR_MAINFRAME),
				IMAGE_ICON,
				rcIcon.right - rcIcon.left,
				rcIcon.bottom - rcIcon.top,
				LR_DEFAULTCOLOR));

			::DrawIconEx(hdc,
				rcIcon.left,
				rcIcon.top,
				hIcon,
				rcIcon.right - rcIcon.left,
				rcIcon.bottom - rcIcon.top,
				NULL,
				GetSysBrush(cBackground),
				DI_NORMAL);

			DestroyIcon(hIcon);
		}

		// Fill in our gutters
		::FillRect(hdc, &rcMenuGutterLeft, GetSysBrush(cBackground));
		::FillRect(hdc, &rcMenuGutterRight, GetSysBrush(cBackground));
		::FillRect(hdc, &rcWindowGutterLeft, GetSysBrush(cBackground));
		::FillRect(hdc, &rcWindowGutterRight, GetSysBrush(cBackground));

		if (iStatusHeight)
		{
			auto rcStatus = rcClient;
			rcStatus.top = rcClient.bottom - iStatusHeight;

			RECT rcFullStatus = { 0 };
			rcFullStatus.top = rcStatus.top;
			rcFullStatus.left = rcWindow.left + BORDER_VISIBLEWIDTH;
			rcFullStatus.right = rcWindow.right - BORDER_VISIBLEWIDTH;
			rcFullStatus.bottom = rcWindow.bottom - BORDER_VISIBLEWIDTH;

			::ExcludeClipRect(hdc, rcStatus.left, rcStatus.top, rcStatus.right, rcStatus.bottom);
			GradientFillRect(hdc, rcFullStatus, cStatus);
		}
		else
		{
			RECT rcBottomGutter = { 0 };
			rcBottomGutter.top = rcWindow.bottom - cyFrame;
			rcBottomGutter.left = rcWindow.left + BORDER_VISIBLEWIDTH;
			rcBottomGutter.right = rcWindow.right - BORDER_VISIBLEWIDTH;
			rcBottomGutter.bottom = rcWindow.bottom - BORDER_VISIBLEWIDTH;

			::FillRect(hdc, &rcBottomGutter, GetSysBrush(cBackground));
		}

		TCHAR szTitle[256] = { 0 };
		GetWindowText(hWnd, szTitle, _countof(szTitle));
		DrawSegoeText(
			hdc,
			szTitle,
			MyGetSysColor(cText),
			&rcCaptionText,
			false,
			DT_LEFT | DT_SINGLELINE | DT_VCENTER);

		// Finally, we paint our border glow if we're not maximized
		if (!::IsZoomed(hWnd))
		{
			auto rcInnerFrame = rcWindow;
			::InflateRect(&rcInnerFrame, -BORDER_VISIBLEWIDTH, -BORDER_VISIBLEWIDTH);
			::ExcludeClipRect(hdc, rcInnerFrame.left, rcInnerFrame.top, rcInnerFrame.right, rcInnerFrame.bottom);
			::FillRect(hdc, &rcWindow, GetSysBrush(bActive ? cGlow : cFrameUnselected));
		}

		db.End(hdc);
		::ReleaseDC(hWnd, hdcWin);
	}
}

_Check_return_ bool HandleControlUI(UINT message, WPARAM wParam, LPARAM lParam, _Out_ LRESULT* lpResult)
{
	if (!lpResult) return false;
	*lpResult = 0;
	switch (message)
	{
	case WM_ERASEBKGND:
		return true;
	case WM_CTLCOLORSTATIC:
	case WM_CTLCOLOREDIT:
	{
		auto lsStyle = reinterpret_cast<intptr_t>(::GetProp(reinterpret_cast<HWND>(lParam), LABEL_STYLE));
		auto uiText = cText;
		auto uiBackground = cBackground;

		if (lsStyle == lsPaneHeader)
		{
			uiText = cPaneHeaderText;
			uiBackground = cPaneHeaderBackground;
		}
		auto hdc = reinterpret_cast<HDC>(wParam);
		if (hdc)
		{
			::SetTextColor(hdc, MyGetSysColor(uiText));
			::SetBkMode(hdc, TRANSPARENT);
			::SelectObject(hdc, GetSegoeFont());
		}
		*lpResult = reinterpret_cast<LRESULT>(GetSysBrush(uiBackground));
		return true;
	}
	case WM_NOTIFY:
	{
		auto pHdr = reinterpret_cast<LPNMHDR>(lParam);

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

void DrawHelpText(_In_ HWND hWnd, _In_ UINT uIDText)
{
	auto hRes = S_OK;
	PAINTSTRUCT ps = { nullptr };
	::BeginPaint(hWnd, &ps);

	if (ps.hdc)
	{
		RECT rcWin = { 0 };
		::GetClientRect(hWnd, &rcWin);

		CDoubleBuffer db;
		auto hdc = ps.hdc;
		db.Begin(hdc, &rcWin);

		auto rcText = rcWin;
		::FillRect(hdc, &rcText, GetSysBrush(cBackground));

		TCHAR szHelpText[256];
		auto iRet = NULL;
		WC_D(iRet, LoadString(GetModuleHandle(NULL),
			uIDText,
			szHelpText,
			_countof(szHelpText)));

		DrawSegoeText(
			hdc,
			szHelpText,
			MyGetSysColor(cTextDisabled),
			&rcText,
			true,
			DT_CENTER | DT_SINGLELINE | DT_VCENTER | DT_NOCLIP | DT_END_ELLIPSIS | DT_NOPREFIX);

		db.End(hdc);
	}
	::EndPaint(hWnd, &ps);
}

// Handle WM_ERASEBKGND so the control won't flicker.
LRESULT CALLBACK LabelProc(
	_In_ HWND hWnd,
	UINT uMsg,
	WPARAM wParam,
	LPARAM lParam,
	UINT_PTR uIdSubclass,
	DWORD_PTR /*dwRefData*/)
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

void SubclassLabel(_In_ HWND hWnd)
{
	SetWindowSubclass(hWnd, LabelProc, 0, 0);
	::SendMessageA(hWnd, WM_SETFONT, reinterpret_cast<WPARAM>(GetSegoeFont()), false);
	::SendMessage(hWnd, EM_SETMARGINS, EC_LEFTMARGIN | EC_RIGHTMARGIN, 0);
}

void StyleLabel(_In_ HWND hWnd, uiLabelStyle lsStyle)
{
	auto hRes = S_OK;

	EC_B(::SetProp(hWnd, LABEL_STYLE, reinterpret_cast<HANDLE>(lsStyle)));
}