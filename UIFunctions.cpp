// UIFunctions.h : Common UI functions for MFCMAPI

#include "stdafx.h"
#include "UIFunctions.h"
#include "Windowsx.h"

HFONT g_hFontSegoe = NULL;
HFONT g_hFontSegoeBold = NULL;

#define SEGOE _T("Segoe UI") // STRING_OK
#define SEGOEW L"Segoe UI" // STRING_OK
#define SEGOEBOLD L"Segoe UI Bold" // STRING_OK

enum myColor
{
	cWhite,
	cGrey,
	cBlack,
	cCyan,
	cMagenta,
	cOrange,
	cMedOrange,
	cLightOrange,
	cColorEnd
};

// Keep in sync with enum myColor
COLORREF g_Colors[cColorEnd] =
{
	RGB(0xFF, 0xFF, 0xFF), // cWhite
	RGB(0xAD, 0xAC, 0xAE), // cGrey
	RGB(0x00, 0x00, 0x00), // cBlack
	RGB(0x00, 0xFF, 0xFF), // cCyan
	RGB(0xFF, 0x00, 0xFF), // cMagenta
	RGB(0xFF, 0x99, 0x33), // cOrange
	RGB(0xFF, 0xB3, 0x23), // cMedOrange
	RGB(0xFF, 0xD0, 0xA0), // cLightOrange
};

// Fixed mapping of UI elements to colors
// Will be overridden by system colors when specified in g_SysColors
// Keep in sync with enum uiColor
// We can swap in cCyan for various entries to test coverage
myColor g_FixedColors[cUIEnd] =
{
	cWhite, // cBackground
	cLightOrange, // cBackgroundDisabled
	cOrange, // cGlow
	cBlack, // cFrameSelected
	cGrey, // cFrameUnselected
	cGrey, // cArrow
	cBlack, // cText
	cGrey, // cTextDisabled
	cWhite, // cTextInverted
	cMagenta, // cBitmapTransparency
};

// Mapping of UI elements to system colors
// NULL entries will get the fixed mapping from g_FixedColors
int g_SysColors[cUIEnd] =
{
	COLOR_WINDOW, // cBackground
	NULL, // cBackgroundDisabled
	NULL, // cGlow
	COLOR_WINDOWTEXT, // cFrameSelected
	COLOR_3DLIGHT, // cFrameUnselected
	COLOR_GRAYTEXT, // cArrow
	COLOR_WINDOWTEXT, // cText
	COLOR_GRAYTEXT, // cTextDisabled
	NULL, // cTextInverted
	NULL, // cBitmapTransparency
};

HBRUSH g_FixedBrushes[cColorEnd] = {0};
HBRUSH g_SysBrushes[cUIEnd] = {0};
HPEN g_Pens[cPenEnd] = {0};
HBITMAP g_Bitmaps[cBitmapEnd] = {0};

void DrawSegoeTextW(
	_In_ HDC hdc,
	_In_z_ LPCWSTR lpchText,
	_In_ COLORREF color,
	_In_ LPRECT lprc,
	bool bBold,
	_In_ UINT format);

void DrawSegoeTextA(
	_In_ HDC hdc,
	_In_z_ LPCTSTR lpchText,
	_In_ COLORREF color,
	_In_ LPRECT lprc,
	bool bBold,
	_In_ UINT format);

#ifdef UNICODE
#define DrawSegoeText DrawSegoeTextW
#else
#define DrawSegoeText DrawSegoeTextA
#endif

void InitializeGDI()
{
} // InitializeGDI

void UninitializeGDI()
{
	if (g_hFontSegoe) ::DeleteObject(g_hFontSegoe);
	g_hFontSegoe = NULL;
	if (g_hFontSegoeBold) ::DeleteObject(g_hFontSegoeBold);
	g_hFontSegoeBold = NULL;

	int i = 0;
	for (i = 0 ; i < cColorEnd ; i++)
	{
		if (g_FixedBrushes[i])
		{
			::DeleteObject(g_FixedBrushes[i]);
			g_FixedBrushes[i] = NULL;
		}
	}
	for (i = 0 ; i < cUIEnd ; i++)
	{
		if (g_SysBrushes[i])
		{
			::DeleteObject(g_SysBrushes[i]);
			g_SysBrushes[i] = NULL;
		}
	}

	for (i = 0 ; i < cPenEnd ; i++)
	{
		if (g_Pens[i])
		{
			::DeleteObject(g_Pens[i]);
			g_Pens[i] = NULL;
		}
	}

	for (i = 0 ; i < cBitmapEnd ; i++)
	{
		if (g_Bitmaps[i])
		{
			::DeleteObject(g_Bitmaps[i]);
			g_Bitmaps[i] = NULL;
		}
	}
} // UninitializeGDI

_Check_return_ LPMENUENTRY CreateMenuEntry(_In_z_ LPWSTR szMenu)
{
	HRESULT hRes = S_OK;
	LPMENUENTRY lpMenu = new MenuEntry;
	if (lpMenu)
	{
		ZeroMemory(lpMenu, sizeof(MenuEntry));
		lpMenu->m_MSAA.dwMSAASignature = MSAA_MENU_SIG;

		size_t iLen = 0;
		WC_H(StringCchLengthW(szMenu, STRSAFE_MAX_CCH, &iLen));

		lpMenu->m_MSAA.pszWText = new WCHAR[iLen+1];
		if (lpMenu->m_MSAA.pszWText)
		{
			WC_H(StringCchCopyW(lpMenu->m_MSAA.pszWText, iLen+1, szMenu));
			lpMenu->m_pName = lpMenu->m_MSAA.pszWText;
			lpMenu->m_MSAA.cchWText = (DWORD) iLen;
		}
		return lpMenu;
	}
	return NULL;
} // CreateMenuEntry

_Check_return_ LPMENUENTRY CreateMenuEntry(UINT iudMenu)
{
	WCHAR szMenu[128] = {0};
	::LoadStringW(GetModuleHandle(NULL), iudMenu, szMenu, _countof(szMenu));
	return CreateMenuEntry(szMenu);
} // CreateMenuEntry

void DeleteMenuEntry(_In_ LPMENUENTRY lpMenu)
{
	if (!lpMenu) return;
	if (lpMenu->m_MSAA.pszWText) delete[] lpMenu->m_MSAA.pszWText;
	delete lpMenu;
} // DeleteMenuEntry

void DeleteMenuEntries(_In_ HMENU hMenu)
{
	UINT nPosition = 0;
	UINT nCount = 0;
	nCount = ::GetMenuItemCount(hMenu);
	if (-1 == nCount) return;

	for (nPosition = 0; nPosition < nCount; nPosition++)
	{
		MENUITEMINFOW menuiteminfo = {0};
		menuiteminfo.cbSize = sizeof(MENUITEMINFOW);
		menuiteminfo.fMask = MIIM_DATA | MIIM_SUBMENU;

		::GetMenuItemInfoW(hMenu, nPosition, true, &menuiteminfo);
		if (menuiteminfo.dwItemData)
		{
			DeleteMenuEntry((LPMENUENTRY) menuiteminfo.dwItemData);
		}

		if (menuiteminfo.hSubMenu)
		{
			DeleteMenuEntries(menuiteminfo.hSubMenu);
		}
	}
} // DeleteMenuEntries

// Walk through the menu structure and convert any string entries to owner draw using MENUENTRY
void ConvertMenuOwnerDraw(_In_ HMENU hMenu, bool bRoot)
{
	UINT nPosition = 0;
	UINT nCount = 0;

	nCount = ::GetMenuItemCount(hMenu);
	if (-1 == nCount) return;

	for (nPosition = 0; nPosition < nCount; nPosition++)
	{
		MENUITEMINFOW menuiteminfo = {0};
		WCHAR szMenu[128] = {0};
		menuiteminfo.cbSize = sizeof(MENUITEMINFOW);
		menuiteminfo.fMask = MIIM_STRING | MIIM_SUBMENU | MIIM_FTYPE;
		menuiteminfo.cch = _countof(szMenu);
		menuiteminfo.dwTypeData = szMenu;

		::GetMenuItemInfoW(hMenu, nPosition, true, &menuiteminfo);
		bool bOwnerDrawn = (menuiteminfo.fType & MF_OWNERDRAW) != 0;
		if (!bOwnerDrawn)
		{
			LPMENUENTRY lpMenuEntry = CreateMenuEntry(szMenu);
			if (lpMenuEntry)
			{
				lpMenuEntry->b_OnMenuBar = bRoot;
				menuiteminfo.fMask = MIIM_DATA | MIIM_FTYPE;
				menuiteminfo.fType |= MF_OWNERDRAW;
				menuiteminfo.dwItemData = (ULONG_PTR) lpMenuEntry;
				::SetMenuItemInfoW(hMenu, nPosition, true, &menuiteminfo);
			}
		}

		if (menuiteminfo.hSubMenu)
		{
			ConvertMenuOwnerDraw(menuiteminfo.hSubMenu, false);
		}
	}
} // ConvertMenuOwnerDraw

void UpdateMenuString(_In_ HWND hWnd, UINT uiMenuTag, UINT uidNewString)
{
	HRESULT hRes = S_OK;

	WCHAR szNewString[128] = {0};
	::LoadStringW(GetModuleHandle(NULL), uidNewString, szNewString, _countof(szNewString));

	DebugPrint(DBGMenu,_T("UpdateMenuString: Changing menu item 0x%X on window %p to \"%s\"\n"),uiMenuTag,hWnd,(LPCTSTR) szNewString);
	HMENU hMenu = ::GetMenu(hWnd);
	if (!hMenu) return;

	MENUITEMINFOW MenuInfo = {0};

	ZeroMemory(&MenuInfo,sizeof(MenuInfo));

	MenuInfo.cbSize = sizeof(MenuInfo);
	MenuInfo.fMask = MIIM_STRING;
	MenuInfo.dwTypeData = szNewString;

	EC_B(SetMenuItemInfoW(
		hMenu,
		uiMenuTag,
		false,
		&MenuInfo));
} // UpdateMenuString

void MergeMenu(_In_ HMENU hMenuDestination, _In_ const HMENU hMenuAdd)
{
	::GetMenuItemCount(hMenuDestination);
	int iMenuDestItemCount = ::GetMenuItemCount(hMenuDestination);
	int iMenuAddItemCount = ::GetMenuItemCount(hMenuAdd);

	if (iMenuAddItemCount == 0) return;

	if (iMenuDestItemCount > 0) ::AppendMenu(hMenuDestination, MF_SEPARATOR, NULL, NULL);

	WCHAR szMenu[128] = {0};
	int iLoop = 0;
	for (iLoop = 0; iLoop < iMenuAddItemCount; iLoop++ )
	{
		::GetMenuStringW(hMenuAdd,iLoop,szMenu,_countof(szMenu),MF_BYPOSITION);

		HMENU hSubMenu = ::GetSubMenu(hMenuAdd, iLoop);
		if (!hSubMenu)
		{
			UINT nState = GetMenuState(hMenuAdd, iLoop, MF_BYPOSITION);
			UINT nItemID = GetMenuItemID(hMenuAdd, iLoop);

			::AppendMenuW(hMenuDestination, nState, nItemID, szMenu);
			iMenuDestItemCount++;
		}
		else
		{
			int iInsertPosDefault = GetMenuItemCount(hMenuDestination);

			HMENU hNewPopupMenu = ::CreatePopupMenu();

			MergeMenu(hNewPopupMenu, hSubMenu);

			::InsertMenuW(
				hMenuDestination,
				iInsertPosDefault,
				MF_BYPOSITION | MF_POPUP | MF_ENABLED,
				(UINT_PTR) hNewPopupMenu,
				szMenu);
			iMenuDestItemCount++;
		}
	}
} // MergeMenu

void DisplayContextMenu(UINT uiClassMenu, UINT uiControlMenu, _In_ HWND hWnd, int x, int y)
{
	if (!uiClassMenu) uiClassMenu = IDR_MENU_DEFAULT_POPUP;
	HMENU hPopup = ::CreateMenu();
	HMENU hContext = ::LoadMenu(NULL, MAKEINTRESOURCE(uiClassMenu));

	::InsertMenu(
		hPopup,
		0,
		MF_BYPOSITION | MF_POPUP,
		(UINT_PTR) hContext,
		_T(""));

	HMENU hRealPopup = ::GetSubMenu(hPopup,0);

	if (hRealPopup)
	{
		if (uiControlMenu)
		{
			HMENU hAppended = ::LoadMenu(NULL,MAKEINTRESOURCE(uiControlMenu));
			if (hAppended)
			{
				MergeMenu(hRealPopup, hAppended);
				::DestroyMenu(hAppended);
			}
		}

		if (IDR_MENU_PROPERTY_POPUP == uiClassMenu)
		{
			(void) ExtendAddInMenu(hRealPopup, MENU_CONTEXT_PROPERTY);
		}

		ConvertMenuOwnerDraw(hRealPopup, false);
		::TrackPopupMenu(hRealPopup, TPM_LEFTALIGN | TPM_RIGHTBUTTON, x, y, NULL, hWnd, NULL);
		DeleteMenuEntries(hRealPopup);
	}

	::DestroyMenu(hContext);
	::DestroyMenu(hPopup);
} // DisplayContextMenu

_Check_return_ int GetEditHeight(_In_ HWND hwndEdit)
{
	HRESULT		hRes = S_OK;
	HFONT		hOldFont = 0;
	HDC			hdc = 0;
	TEXTMETRIC	tmFont = {0};
	int			iHeight = 0;

	// Get the DC for the edit control.
	WC_D(hdc, GetDC(hwndEdit));

	if (hdc)
	{
		// Get the metrics for the Segoe font.
		hOldFont = (HFONT) SelectObject(hdc, GetSegoeFont());
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
} // GetEditHeight

_Check_return_ int GetTextHeight(_In_ HWND hwndEdit)
{
	HRESULT		hRes = S_OK;
	HFONT		hOldFont = 0;
	HDC			hdc = 0;
	TEXTMETRIC	tmFont = {0};
	int			iHeight = 0;

	// Get the DC for the edit control.
	WC_D(hdc, GetDC(hwndEdit));

	if (hdc)
	{
		// Get the metrics for the Segoe font.
		hOldFont = (HFONT) SelectObject(hdc, GetSegoeFont());
		WC_B(::GetTextMetrics(hdc, &tmFont));
		SelectObject(hdc, hOldFont);
		ReleaseDC(hwndEdit, hdc);
	}

	// Calculate the new height for the static control.
	iHeight = tmFont.tmHeight;
	return iHeight;
} // GetTextHeight

int CALLBACK EnumFontFamExProcW(
	_In_ LPLOGFONTW lplf,
	_In_ NEWTEXTMETRICEXW* /*lpntme*/,
	DWORD /*FontType*/,
	LPARAM lParam)
{
	// Use a 9 point font
	lplf->lfHeight = -MulDiv(9, GetDeviceCaps(::GetDC(NULL), LOGPIXELSY), 72);
	lplf->lfWidth = 0;
	*((HFONT *) lParam) = CreateFontIndirectW(lplf);
	return 0;
} // EnumFontFamExProcW

// This font is not cached and must be delete manually
HFONT GetFont(_In_z_ LPWSTR szFont)
{
	HFONT hFont = 0;
	LOGFONTW lfFont = {0};
	HRESULT hRes = S_OK;
	WC_H(StringCchCopyW(lfFont.lfFaceName, _countof(lfFont.lfFaceName), szFont));

	EnumFontFamiliesExW(GetDC(NULL), &lfFont, (FONTENUMPROCW) EnumFontFamExProcW, (LPARAM) &hFont, 0);
	if (hFont) return hFont;

	// If we can't get our font, fallback to a stock font
	static LOGFONTW lf = {0};
	// This gets the font 'MS Shell Dlg' in testing
	GetObjectW(GetStockObject(DEFAULT_GUI_FONT), sizeof(lf), &lf);

	// Create the font, and then return its handle.
	hFont = CreateFontW(lf.lfHeight, lf.lfWidth,
		lf.lfEscapement, lf.lfOrientation, lf.lfWeight,
		lf.lfItalic, lf.lfUnderline, lf.lfStrikeOut, lf.lfCharSet,
		lf.lfOutPrecision, lf.lfClipPrecision, lf.lfQuality,
		lf.lfPitchAndFamily, lf.lfFaceName);
	return hFont;
} // GetFont

// Cached for deletion in UninitializeGDI
HFONT GetSegoeFont()
{
	if (g_hFontSegoe) return g_hFontSegoe;
	g_hFontSegoe = GetFont(SEGOEW);
	return g_hFontSegoe;
} // GetSegoeFont

// Cached for deletion in UninitializeGDI
HFONT GetSegoeFontBold()
{
	if (g_hFontSegoeBold) return g_hFontSegoeBold;
	g_hFontSegoeBold = GetFont(SEGOEBOLD);
	return g_hFontSegoeBold;
} // GetSegoeFontBold

_Check_return_ HBRUSH GetSysBrush(uiColor uc)
{
	// Return a cached brush if we have one
	if (g_SysBrushes[uc]) return g_SysBrushes[uc];
	// No cached brush found, cache and return a system brush if requested
	int iSysColor = g_SysColors[uc];
	if (iSysColor)
	{
		g_SysBrushes[uc] = GetSysColorBrush(iSysColor);
		return g_SysBrushes[uc];
	}

	// No system brush for this color, cache and return a solid brush of the requested color
	myColor mc = g_FixedColors[uc];
	if (g_FixedBrushes[mc]) return g_FixedBrushes[mc];
	g_FixedBrushes[mc] = CreateSolidBrush(g_Colors[mc]);
	return g_FixedBrushes[mc];
} // GetSysBrush

_Check_return_ COLORREF MyGetSysColor(uiColor uc)
{
	// Return a system color if we have one in g_SysColors
	int iSysColor = g_SysColors[uc];
	if (iSysColor) return ::GetSysColor(iSysColor);

	// No system color listed in g_SysColors, return a hard coded color
	myColor mc = g_FixedColors[uc];
	return g_Colors[mc];
} // MyGetSysColor

_Check_return_ HPEN GetPen(uiPen up)
{
	if (g_Pens[up]) return g_Pens[up];
	LOGBRUSH lbr = {0};
	lbr.lbStyle = BS_SOLID;

	switch (up)
	{
	case cSolidPen:
		{
			lbr.lbColor = MyGetSysColor(cFrameSelected);
			g_Pens[cSolidPen] = ExtCreatePen(PS_SOLID, 1, &lbr, 0, NULL);
			return g_Pens[cSolidPen];
		}
		break;
	case cSolidGreyPen:
		{
			lbr.lbColor = MyGetSysColor(cFrameUnselected);
			g_Pens[cSolidGreyPen] = ExtCreatePen(PS_SOLID, 1, &lbr, 0, NULL);
			return g_Pens[cSolidGreyPen];
		}
		break;
	case cDashedPen:
		{
			lbr.lbColor = MyGetSysColor(cFrameUnselected);
			DWORD rgStyle[2] = {1,3};
			g_Pens[cDashedPen] = ExtCreatePen(PS_GEOMETRIC | PS_USERSTYLE, 1, &lbr, 2, rgStyle);
			return g_Pens[cDashedPen];
		}
	break;
	}
	return NULL;
} // GetPen

HBITMAP GetBitmap(uiBitmap ub)
{
	if (g_Bitmaps[ub]) return g_Bitmaps[ub];

	switch (ub)
	{
	case cNotify:
		{
			g_Bitmaps[cNotify] = ::LoadBitmap(GetModuleHandle(NULL), MAKEINTRESOURCE(IDB_ADVISE));
			return g_Bitmaps[cNotify];
		}
		break;
	}
	return NULL;
} // GetBitmap

void DrawSegoeTextW(
	_In_ HDC hdc,
	_In_z_ LPCWSTR lpchText,
	_In_ COLORREF color,
	_In_ LPRECT lprc,
	bool bBold,
	_In_ UINT format)
{
	HGDIOBJ hfontOld = ::SelectObject(hdc, bBold?GetSegoeFontBold():GetSegoeFont());
	COLORREF crText = ::SetTextColor(hdc, color);
	::SetBkMode(hdc,TRANSPARENT);
	::DrawTextW(
		hdc,
		lpchText,
		-1,
		lprc,
		format);
	::SelectObject(hdc, hfontOld);
	(void) ::SetTextColor(hdc, crText);
} // DrawSegoeTextW

void DrawSegoeTextA(
	_In_ HDC hdc,
	_In_z_ LPCSTR lpchText,
	_In_ COLORREF color,
	_In_ LPRECT lprc,
	bool bBold,
	_In_ UINT format)
{
	HGDIOBJ hfontOld = ::SelectObject(hdc, bBold?GetSegoeFontBold():GetSegoeFont());
	COLORREF crText = ::SetTextColor(hdc, color);
	::SetBkMode(hdc,TRANSPARENT);
	::DrawTextA(
		hdc,
		lpchText,
		-1,
		lprc,
		format);
	::SelectObject(hdc, hfontOld);
	(void) ::SetTextColor(hdc, crText);
} // DrawSegoeTextA

// Clear/initialize formatting on the rich edit control.
// We have to force load the system riched20 to ensure this doesn't break since
// Office's riched20 apparently doesn't handle CFM_COLOR at all.
// Sets our text color, script, and turns off bold, italic, etc formatting.
void ClearEditFormatting(_In_ HWND hWnd)
{
	CHARFORMAT2 cf;
	ZeroMemory(&cf, sizeof(cf));
	cf.cbSize = sizeof(cf);
	cf.dwMask = CFM_COLOR | CFM_FACE | CFM_BOLD | CFM_ITALIC | CFM_UNDERLINE | CFM_STRIKEOUT;
	cf.crTextColor = MyGetSysColor(cText);
	StringCchCopy(cf.szFaceName, _countof(cf.szFaceName), SEGOE);
	(void) ::SendMessage(hWnd, EM_SETCHARFORMAT, SCF_ALL, (LPARAM)&cf);
} // ClearEditFormatting

void GradientFillRect(_In_ HDC hdc, RECT rc)
{
	// Gradient fill the background
	TRIVERTEX vertex[2] = {0};
	// Medium orange at the top
	vertex[0].x     = rc.left;
	vertex[0].y     = rc.top;
	vertex[0].Red   = GetRValue(g_Colors[cMedOrange]) << 8;
	vertex[0].Green = GetGValue(g_Colors[cMedOrange]) << 8;
	vertex[0].Blue  = GetBValue(g_Colors[cMedOrange]) << 8;
	vertex[0].Alpha = 0x0000;

	// Dark orange at the bottom
	vertex[1].x     = rc.right;
	vertex[1].y     = rc.bottom;
	vertex[1].Red   = GetRValue(g_Colors[cOrange]) << 8;
	vertex[1].Green = GetGValue(g_Colors[cOrange]) << 8;
	vertex[1].Blue  = GetBValue(g_Colors[cOrange]) << 8;
	vertex[1].Alpha = 0x0000;

	// Create a GRADIENT_RECT structure that references the TRIVERTEX vertices. 
	GRADIENT_RECT gRect = {0};
	gRect.UpperLeft  = 0;
	gRect.LowerRight = 1;
	::GradientFill(hdc, vertex, 2, &gRect, 1, GRADIENT_FILL_RECT_V);
} // GradientFillRect

void DrawFilledPolygon(_In_ HDC hdc, _In_count_(cpt) CONST POINT *apt, _In_ int cpt, COLORREF cEdge, _In_ HBRUSH hFill)
{
	HPEN hPen = ::CreatePen(PS_SOLID, 0, cEdge);
	HGDIOBJ hBrushOld = ::SelectObject(hdc, hFill);
	HGDIOBJ hPenOld = ::SelectObject(hdc, hPen);
	::Polygon(hdc, apt, cpt);
	::SelectObject(hdc, hPenOld);
	::SelectObject(hdc, hBrushOld);
	::DeleteObject(hPen);
} // DrawFilledPolygon

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
		break;
	case WM_NCPAINT:
		{
			HDC hdc = ::GetWindowDC(hWnd);
			if (hdc)
			{
				RECT rc = {0};
				::GetWindowRect(hWnd, &rc);
				::OffsetRect(&rc, -rc.left, -rc.top);
				UINT iOpts = (UINT) ::SendMessage(hWnd, EM_GETOPTIONS, NULL, NULL);
				bool bReadOnly = (iOpts & ECO_READONLY) != 0;
				::FrameRect(hdc, &rc, GetSysBrush(bReadOnly?cFrameUnselected:cFrameSelected));
				::ReleaseDC(hWnd, hdc);
			}
			return 0;
		}
		break;
	// Handle measuring and painting for context menus
	case WM_MEASUREITEM:
		MeasureItem((LPMEASUREITEMSTRUCT) lParam);
		return true;
		break;
	case WM_DRAWITEM:
		DrawItem((LPDRAWITEMSTRUCT) lParam);
		return true;
		break;
	}
	return DefSubclassProc(hWnd, uMsg, wParam, lParam);
} // DrawEditProc

void CustomDrawList(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult, int iItemCur)
{
	static bool bSelected = false;
	if (!pNMHDR) return;
	LPNMLVCUSTOMDRAW lvcd = (LPNMLVCUSTOMDRAW) pNMHDR;
	int iItem = (int) lvcd->nmcd.dwItemSpec;

	switch (lvcd->nmcd.dwDrawStage)
	{
	case CDDS_PREPAINT:
		*pResult = CDRF_NOTIFYITEMDRAW | CDRF_NOTIFYPOSTPAINT;
		break;

	case CDDS_ITEMPREPAINT:
		// Let the item draw itself without a selection highlight
		bSelected = ListView_GetItemState(lvcd->nmcd.hdr.hwndFrom, iItem, LVIS_SELECTED) != 0;
		if (bSelected)
		{
			// Turn off listview selection highlight
			ListView_SetItemState(lvcd->nmcd.hdr.hwndFrom, iItem, 0, LVIS_SELECTED)
		}
		*pResult = CDRF_DODEFAULT | CDRF_NOTIFYPOSTPAINT;
		break;

	case CDDS_ITEMPOSTPAINT:
		// And then we'll handle our frame
		if (bSelected)
		{
			// Turn on listview selection highlight
			ListView_SetItemState(lvcd->nmcd.hdr.hwndFrom, iItem, LVIS_SELECTED, LVIS_SELECTED)
		}

		break;

	case CDDS_POSTPAINT:
		{
			if (-1 != iItemCur) DrawListItemFrame(lvcd->nmcd.hdr.hwndFrom, lvcd->nmcd.hdc, iItemCur, LVIS_GLOW);

			UINT iNumItems = (UINT) ::SendMessage(lvcd->nmcd.hdr.hwndFrom, LVM_GETSELECTEDCOUNT, 0, 0);
			UINT iArrayPos = 0;
			int iSelectedItem = -1;

			if (!iNumItems) break;

			for (iArrayPos = 0 ; iArrayPos < iNumItems ; iArrayPos++)
			{
				iSelectedItem = (int) ::SendMessage(lvcd->nmcd.hdr.hwndFrom, LVM_GETNEXTITEM, iSelectedItem, LVNI_SELECTED);
				if (-1 == iSelectedItem) break;

				if (iItemCur != iSelectedItem) DrawListItemFrame(lvcd->nmcd.hdr.hwndFrom, lvcd->nmcd.hdc, iSelectedItem, LVIS_SELECTED);
			}

			break;
		}

	default:
		*pResult = CDRF_DODEFAULT;
		break;
	}
} // CustomDrawList

// Handle highlight and selection frames for list items
void DrawListItemFrame(_In_ HWND hWnd, _In_opt_ HDC hdc, UINT itemID, UINT itemState)
{
	if (-1 == itemID) return;
	HDC hdcLocal = NULL;
	RECT rcIcon = {0};
	RECT rcLabels = {0};
	ListView_GetItemRect(hWnd,itemID, &rcLabels, LVIR_BOUNDS);
	ListView_GetItemRect(hWnd,itemID, &rcIcon, LVIR_ICON);
	rcLabels.left = rcIcon.right;
	if (rcLabels.left >= rcLabels.right) return;
	bool bSelected = ListView_GetItemState(hWnd, itemID, LVIS_SELECTED) != 0;
	if (bSelected) itemState |= LVIS_SELECTED;
	if (!hdc)
	{
		hdcLocal = ::GetDC(hWnd);
		hdc = hdcLocal;
	}

	if (hdc)
	{
		uiColor cFrame = cBackground;
		if (itemState & LVIS_GLOW)
		{
			cFrame = cGlow;
		}
		else if (itemState & LVIS_SELECTED)
		{
			cFrame = cFrameSelected;
		}
		::FrameRect(hdc, &rcLabels, GetSysBrush(cFrame));
	}
	if (hdcLocal) ::ReleaseDC(hWnd, hdcLocal);
} // DrawListItemFrame

void DrawTreeItemFrame(_In_ HWND hWnd, HTREEITEM hItem, bool bDraw)
{
	RECT rect = {0};
	RECT rectTree = {0};
	TreeView_GetItemRect(hWnd, hItem, &rect, 1);
	::GetClientRect(hWnd, &rectTree);
	rect.left = rectTree.left;
	rect.right = rectTree.right;
	HDC hdc = ::GetDC(hWnd);
	if (hdc)
	{
		::FrameRect(hdc, &rect, GetSysBrush(bDraw ? cGlow : cBackground));
		::ReleaseDC(hWnd, hdc);
	}
} // DrawTreeItemFrame

void CustomDrawTree(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult, bool bHover, _In_ HTREEITEM hItemCurHover)
{
	if (!pNMHDR) return;
	static bool bHighlighted = false;
	LPNMTVCUSTOMDRAW lvcd = (LPNMTVCUSTOMDRAW) pNMHDR;

	switch (lvcd->nmcd.dwDrawStage)
	{
	case CDDS_PREPAINT:
		*pResult = CDRF_NOTIFYITEMDRAW;
		break;
	case CDDS_ITEMPREPAINT:
		{
			HTREEITEM hItem = (HTREEITEM) lvcd->nmcd.dwItemSpec;

			if (hItem)
			{
				lvcd->clrTextBk = MyGetSysColor(cBackground);
				lvcd->clrText = MyGetSysColor(cText);
				int iState = TreeView_GetItemState(lvcd->nmcd.hdr.hwndFrom, hItem, TVIS_SELECTED);
				TreeView_SetItemState(lvcd->nmcd.hdr.hwndFrom, hItem, iState & TVIS_SELECTED ? TVIS_BOLD : NULL, TVIS_BOLD);
			}

			*pResult = CDRF_DODEFAULT | CDRF_NOTIFYPOSTPAINT;
			break;
		}
	case CDDS_ITEMPOSTPAINT:
		{
			// If we've advised on this object, add a little icon to let the user know
			HTREEITEM hItem = (HTREEITEM) lvcd->nmcd.dwItemSpec;

			if (hItem)
			{
				// Cover over the +/- and paint triangles instead
				DrawExpandTriangle(lvcd->nmcd.hdr.hwndFrom, lvcd->nmcd.hdc, hItem, bHover && (hItem == hItemCurHover));

				TVITEM tvi = {0};
				tvi.mask = TVIF_PARAM;
				tvi.hItem = hItem;
				TreeView_GetItem(lvcd->nmcd.hdr.hwndFrom, &tvi);
				SortListData* lpData = (SortListData*) tvi.lParam;
				if (lpData && lpData->data.Node.lpAdviseSink)
				{
					RECT rect = {0};
					TreeView_GetItemRect(lvcd->nmcd.hdr.hwndFrom, hItem, &rect, 1);
					HDC dcBitmap = ::CreateCompatibleDC(lvcd->nmcd.hdc);
					if (dcBitmap)
					{
						HBITMAP hbm = GetBitmap(cNotify);
						if (hbm)
						{
							BITMAP bm = {0};
							(void) ::SelectObject(dcBitmap, hbm);
							::GetObject(GetBitmap(cNotify), sizeof(bm), &bm);

							(void) TransparentBlt(
								lvcd->nmcd.hdc,
								rect.right,
								rect.top,
								bm.bmWidth,
								bm.bmHeight,
								dcBitmap,
								0,
								0,
								bm.bmWidth,
								bm.bmHeight,
								MyGetSysColor(cBitmapTransparency));
						}
						::DeleteDC(dcBitmap);
					}
				}

				// If the mouse is hovering over this item, paint the hover
				// frame in case we messed it up
				if (hItem == hItemCurHover)
				{
					DrawTreeItemFrame(lvcd->nmcd.hdr.hwndFrom, hItemCurHover, true);
				}
			}
			break;
		}
	default:
		*pResult = CDRF_DODEFAULT;
		break;
	}
} // CustomDrawTree

// Paints the triangles indicating expansion state
void DrawExpandTriangle(_In_ HWND hWnd, _In_ HDC hdc, _In_ HTREEITEM hItem, bool bGlow)
{
	TVITEM tvitem = {0};
	tvitem.hItem = hItem;
	tvitem.mask  = TVIF_CHILDREN | TVIF_STATE;
	TreeView_GetItem(hWnd, &tvitem);
	bool bHasChildren = tvitem.cChildren != 0;
	if (bHasChildren)
	{
		bool bExpanded = (TVIS_EXPANDED == (tvitem.state & TVIS_EXPANDED));
		RECT rect = {0};
		TreeView_GetItemRect(hWnd, hItem, &rect, 1);

		// Build a box to erase the +/- icons
		RECT rcPlusMinus = rect;
		rcPlusMinus.top++;
		rcPlusMinus.bottom--;
		rcPlusMinus.left = rect.left - 14;
		rcPlusMinus.right = rcPlusMinus.left + 9;
		::FillRect(hdc, &rcPlusMinus, GetSysBrush(cBackground));

		POINT tri[3] = {0};
		if (bExpanded)
		{
			tri[0].x = rect.left-14;
			tri[0].y = rect.top+11;
			tri[1].x = tri[0].x+5;
			tri[1].y = tri[0].y-5;
			tri[2].x = tri[0].x+5;
			tri[2].y = tri[0].y;
		}
		else
		{
			tri[0].x = rect.left-14;
			tri[0].y = rect.top+4;
			tri[1].x = tri[0].x;
			tri[1].y = tri[0].y+8;
			tri[2].x = tri[0].x+4;
			tri[2].y = tri[0].y+4;
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
			cEdge = bExpanded?cFrameSelected:cArrow;
			cFill = bExpanded?cFrameSelected:cBackground;
		}
		DrawFilledPolygon(hdc, tri, _countof(tri), MyGetSysColor(cEdge), GetSysBrush(cFill));
	}
} // DrawExpandTriangle

void DrawHeaderItem(_In_ HWND hWnd, _In_ HDC hdc, UINT itemID, _In_ LPRECT lprc)
{
	bool bSorted = false;
	bool bSortUp = true;

	WCHAR szHeader[255] = {0};
	HDITEMW hdItem = {0};
	hdItem.mask = HDI_TEXT | HDI_FORMAT;
	hdItem.pszText = szHeader;
	hdItem.cchTextMax = _countof(szHeader);
	::SendMessage(hWnd, HDM_GETITEMW, (WPARAM) itemID, (LPARAM) &hdItem);

	if (hdItem.fmt & (HDF_SORTUP | HDF_SORTDOWN))
		bSorted = true;
	if (bSorted)
	{
		bSortUp = (hdItem.fmt & HDF_SORTUP) != 0;
	}

	HDC hdcLocal = CreateCompatibleDC(hdc);
	HBITMAP hbm = CreateCompatibleBitmap(
		hdc,
		lprc->right - lprc->left,
		lprc->bottom - lprc->top);
	if (hdcLocal && hbm)
	{
		HGDIOBJ hbmOld = ::SelectObject(hdcLocal, hbm);
		RECT rcHeader = *lprc;
		OffsetRect(&rcHeader, -rcHeader.left, -rcHeader.top);
		::FillRect(hdcLocal, &rcHeader, GetSysBrush(cBackground));

		if (bSorted)
		{
			POINT tri[3] = {0};
			LONG lCenter = (rcHeader.left + rcHeader.right)/2;
			LONG lTop = rcHeader.top + GetSystemMetrics(SM_CYBORDER);
			if (bSortUp)
			{
				tri[0].x = lCenter;
				tri[0].y = lTop;
				tri[1].x = lCenter-4;
				tri[1].y = lTop+4;
				tri[2].x = lCenter+4;
				tri[2].y = lTop+4;
			}
			else
			{
				tri[0].x = lCenter;
				tri[0].y = lTop+4;
				tri[1].x = lCenter-4;
				tri[1].y = lTop;
				tri[2].x = lCenter+4;
				tri[2].y = lTop;
			}
			DrawFilledPolygon(hdcLocal, tri, _countof(tri), MyGetSysColor(cArrow), GetSysBrush(cArrow));
		}

		RECT rcText = rcHeader;
		rcText.left += GetSystemMetrics(SM_CXEDGE);
		DrawSegoeTextW(
			hdcLocal,
			hdItem.pszText,
			MyGetSysColor(cText),
			&rcText,
			false,
			DT_END_ELLIPSIS | DT_SINGLELINE | DT_VCENTER);

		// Draw our divider
		::InflateRect (&rcHeader, 0, -1);
		rcHeader.left = rcHeader.right - 2;
		::FrameRect(hdcLocal, &rcHeader, GetSysBrush(cFrameUnselected));

		BitBlt(
			hdc,
			lprc->left,
			lprc->top,
			lprc->right - lprc->left,
			lprc->bottom - lprc->top,
			hdcLocal,
			0,
			0,
			SRCCOPY);
		(void) ::SelectObject(hdcLocal, hbmOld);
	}
	if (hdcLocal) DeleteDC(hdcLocal);
	if (hbm) DeleteObject(hbm);
} // DrawHeaderItem

// Draw the unused portion of the header
void CustomDrawHeader(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
{
	if (!pNMHDR) return;
	LPNMCUSTOMDRAW lvcd = (LPNMCUSTOMDRAW) pNMHDR;

	switch (lvcd->dwDrawStage)
	{
	case CDDS_PREPAINT:
		*pResult = CDRF_NOTIFYPOSTPAINT | CDRF_NOTIFYITEMDRAW;
		break;

	case CDDS_ITEMPREPAINT:
		DrawHeaderItem(lvcd->hdr.hwndFrom, lvcd->hdc, (UINT) lvcd->dwItemSpec, &lvcd->rc);
		*pResult = CDRF_SKIPDEFAULT;
		break;

	case CDDS_POSTPAINT:
		{
			RECT rc = {0};
			::GetClientRect(lvcd->hdr.hwndFrom, &rc);
			int iCount = Header_GetItemCount(lvcd->hdr.hwndFrom);
			if (iCount)
			{
				RECT rcRight = {0};
				(void) Header_GetItemRect(lvcd->hdr.hwndFrom, iCount - 1, &rcRight);
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
} // CustomDrawHeader

void DrawTrackingBar(_In_ HWND hWndHeader, _In_ HWND hWndList, int x, int iHeaderHeight, bool bRedraw)
{
	RECT rcTracker = {0};
	::GetClientRect(hWndList, &rcTracker);
	HDC hdc = ::GetDC(hWndList);
	rcTracker.top += iHeaderHeight;
	rcTracker.left = x - 1; // this lines us up under the splitter line we drew in the header
	rcTracker.right = x;
	::MapWindowPoints(hWndHeader, hWndList, (LPPOINT) &rcTracker, 2);
	if (bRedraw)
	{
		::InvalidateRect(hWndList, &rcTracker, true);
	}
	else
	{
		::FillRect(hdc, &rcTracker, GetSysBrush(cFrameSelected));
	}
	::ReleaseDC(hWndList, hdc);
} // DrawTrackingBar

void DrawButton(_In_ HWND hWnd, _In_ HDC hDC, _In_ LPRECT lprc, UINT itemState)
{
	WCHAR szButton[255] = {0};
	::GetWindowTextW(hWnd, szButton, _countof(szButton));
	int iState = (int) ::SendMessage(hWnd, BM_GETSTATE, NULL, NULL);
	bool bGlow = (iState & BST_HOT) != 0;
	bool bPushed = (iState & BST_PUSHED) != 0;
	bool bDisabled = (itemState & CDIS_DISABLED) != 0;
	bool bFocused = (itemState & CDIS_FOCUS) != 0;

	::FillRect(hDC, lprc, GetSysBrush(cBackground));
	::FrameRect(hDC ,lprc, (bFocused || bGlow || bPushed)?GetSysBrush(cFrameSelected):GetSysBrush(cFrameUnselected));

	DrawSegoeTextW(
		hDC,
		szButton,
		(bPushed || bDisabled)?MyGetSysColor(cTextDisabled):MyGetSysColor(cText),
		lprc,
		false,
		DT_SINGLELINE | DT_VCENTER | DT_CENTER);
} // DrawButton

void DrawCheckButton(_In_ HWND hWnd, _In_ HDC hDC, _In_ LPRECT lprc, UINT itemState)
{
	WCHAR szButton[255];
	::GetWindowTextW(hWnd, szButton, _countof(szButton));
	int iState = (int) ::SendMessage(hWnd, BM_GETSTATE, NULL, NULL);
	bool bGlow = (BST_HOT == (iState & BST_HOT));
	bool bChecked = (iState & BST_CHECKED) != 0;
	bool bDisabled = (itemState & CDIS_DISABLED) != 0;
	bool bFocused = (itemState & CDIS_FOCUS) != 0;

	long lSpacing = GetSystemMetrics(SM_CYEDGE);
	long lCheck = lprc->bottom - lprc->top - 2 * lSpacing;
	RECT rcCheck = {0};
	rcCheck.right = rcCheck.left + lCheck;
	rcCheck.top = lprc->top + lSpacing;
	rcCheck.bottom = lprc->bottom - lSpacing;

	::FillRect(hDC, lprc, GetSysBrush(cBackground));
	::FrameRect(hDC, &rcCheck, GetSysBrush(bDisabled?cFrameUnselected:((bGlow || bFocused)?cGlow:cFrameSelected)));
	if (bChecked)
	{
		RECT rcFill = rcCheck;
		::InflateRect(&rcFill, -3, -3);
		::FillRect(hDC, &rcFill, GetSysBrush(cGlow));
	}

	lprc->left = rcCheck.right + lSpacing;

	DrawSegoeTextW(
		hDC,
		szButton,
		bDisabled?MyGetSysColor(cTextDisabled):MyGetSysColor(cText),
		lprc,
		false,
		DT_SINGLELINE | DT_VCENTER);
} // DrawCheckButton

bool CustomDrawButton(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult)
{
	if (!pNMHDR) return false;
	static bool bHighlighted = false;
	LPNMCUSTOMDRAW lvcd = (LPNMCUSTOMDRAW) pNMHDR;

	// Ensure we only handle painting for buttons
	WCHAR szClass[64] = {0};
	::GetClassNameW(lvcd->hdr.hwndFrom, szClass, _countof(szClass));
	if (_wcsicmp(szClass, L"BUTTON")) return false;

	switch (lvcd->dwDrawStage)
	{
	case CDDS_PREPAINT:
		{
			LONG lStyle = BS_TYPEMASK & ::GetWindowLongA(lvcd->hdr.hwndFrom, GWL_STYLE);

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
} // CustomDrawButton

void MeasureMenu(_In_ LPMEASUREITEMSTRUCT lpMeasureItemStruct)
{
	if (!lpMeasureItemStruct) return;
	LPMENUENTRY lpMenuEntry = (LPMENUENTRY) lpMeasureItemStruct->itemData;
	if (!lpMenuEntry) return;

	HDC hdc = ::GetDC(NULL);
	if (hdc)
	{
		HGDIOBJ hfontOld = NULL;
		hfontOld = ::SelectObject(hdc, GetSegoeFont());

		// In order to compute the right width, we need to drop our prefix characters
		CString szText = lpMenuEntry->m_MSAA.pszWText;
		szText.Replace(_T("&"), _T("")); // STRING_OK

		SIZE size = {0};
		GetTextExtentPoint32(hdc, (LPCTSTR) szText, szText.GetLength(), &size);
		lpMeasureItemStruct->itemWidth = size.cx + 2* GetSystemMetrics(SM_CXEDGE);
		lpMeasureItemStruct->itemHeight = size.cy + 2* GetSystemMetrics(SM_CYEDGE);

		// Make sure we have room for the flyout icon
		if (!lpMenuEntry->b_OnMenuBar && ::IsMenu((HMENU) lpMeasureItemStruct->itemID))
		{
			lpMeasureItemStruct->itemWidth += GetSystemMetrics(SM_CXSMICON);
		}

		::SelectObject(hdc, hfontOld);
		::ReleaseDC(NULL, hdc);
	}
} // MeasureMenu

void MeasureItem(_In_ LPMEASUREITEMSTRUCT lpMeasureItemStruct)
{
	if (ODT_MENU == lpMeasureItemStruct->CtlType)
	{
		MeasureMenu(lpMeasureItemStruct);
	}
	// Important to keep this case even if we do not use it - CDialog::OnMeasureItem asserts.
	else if (ODT_COMBOBOX == lpMeasureItemStruct->CtlType)
	{
//		DebugPrint(DBGGeneric,"Combo Box\n");
	}
} // MeasureItem

void DrawMenu(_In_ LPDRAWITEMSTRUCT lpDrawItemStruct)
{
	if (!lpDrawItemStruct) return;
	LPMENUENTRY lpMenuEntry = (LPMENUENTRY) lpDrawItemStruct->itemData;
	if (!lpMenuEntry) return;

	bool bSeparator = !lpMenuEntry->m_MSAA.cchWText; // Cheap way to find separators. Revisit if odd effects show.
	bool bAccel = (lpDrawItemStruct->itemState & ODS_NOACCEL) != 0;
	bool bHot = (lpDrawItemStruct->itemState & (ODS_HOTLIGHT|ODS_SELECTED)) != 0;
	bool bDisabled = (lpDrawItemStruct->itemState & (ODS_GRAYED|ODS_DISABLED)) != 0;

	// Double buffer our menu painting
	HDC hdc = CreateCompatibleDC(lpDrawItemStruct->hDC);
	HBITMAP hbm = CreateCompatibleBitmap(
		lpDrawItemStruct->hDC,
		lpDrawItemStruct->rcItem.right - lpDrawItemStruct->rcItem.left,
		lpDrawItemStruct->rcItem.bottom - lpDrawItemStruct->rcItem.top);
	if (hdc && hbm)
	{
		HGDIOBJ hbmOld = ::SelectObject(hdc, hbm);

		// Draw background
		RECT rcItem = lpDrawItemStruct->rcItem;
		OffsetRect(&rcItem, -rcItem.left, -rcItem.top);
		RECT rcText = rcItem;

		if (!lpMenuEntry->b_OnMenuBar)
		{
			RECT rectGutter = rcText;
			rcText.left += GetSystemMetrics(SM_CXMENUCHECK);
			rectGutter.right = rcText.left;
			::FillRect(hdc, &rectGutter, GetSysBrush(cBackgroundDisabled));
		}

		if (bHot && !bDisabled)
		{
			::FillRect(hdc, &rcItem, GetSysBrush(cGlow));
		}
		else
		{
			::FillRect(hdc, &rcText, GetSysBrush(cBackground));
		}

		if (bSeparator)
		{
			::InflateRect(&rcText, -3, 0);
			LONG lMid = (rcText.bottom + rcText.top) / 2;
			HGDIOBJ hpenOld = ::SelectObject(hdc, GetPen(cSolidGreyPen));
			::MoveToEx(hdc, rcText.left, lMid, NULL);
			::LineTo(hdc, rcText.right, lMid);
			(void) ::SelectObject(hdc, hpenOld);
		}
		else if (lpMenuEntry->m_pName)
		{
			// Set text color
			COLORREF crText = MyGetSysColor(cText);
			if (bHot && !bDisabled)
			{
				crText = MyGetSysColor(cTextInverted);
			}
			else if (bDisabled)
			{
				crText = MyGetSysColor(cTextDisabled);
			}

			UINT uiTextFlags = DT_SINGLELINE | DT_VCENTER;
			if (bAccel) uiTextFlags |= DT_HIDEPREFIX;

			if (lpMenuEntry->b_OnMenuBar)
				uiTextFlags |= DT_CENTER;
			else
				rcText.left += GetSystemMetrics(SM_CXEDGE);

			DrawSegoeTextW(
				hdc,
				lpMenuEntry->m_pName,
				crText,
				&rcText,
				false,
				uiTextFlags);

			// Triple buffer the check mark so we can copy it over without the background
			if (lpDrawItemStruct->itemState & ODS_CHECKED)
			{	
				UINT nWidth = GetSystemMetrics(SM_CXMENUCHECK);
				UINT nHeight = GetSystemMetrics(SM_CYMENUCHECK);
				RECT rc = {0};
				HBITMAP bm = CreateBitmap(nWidth, nHeight, 1, 1, NULL);
				HDC hdcMem = CreateCompatibleDC(hdc);

				SelectObject(hdcMem, bm);
				SetRect(&rc, 0, 0, nWidth, nHeight);
				(void) DrawFrameControl(hdcMem, &rc, DFC_MENU, DFCS_MENUCHECK);

				(void) TransparentBlt(
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

		BitBlt(
			lpDrawItemStruct->hDC,
			lpDrawItemStruct->rcItem.left,
			lpDrawItemStruct->rcItem.top,
			lpDrawItemStruct->rcItem.right - lpDrawItemStruct->rcItem.left,
			lpDrawItemStruct->rcItem.bottom - lpDrawItemStruct->rcItem.top,
			hdc,
			0,
			0,
			SRCCOPY);
		(void) ::SelectObject(hdc,hbmOld);
	}
	if (hdc) DeleteDC(hdc);
	if (hbm) DeleteObject(hbm);
} // DrawMenu

void DrawComboBox(_In_ LPDRAWITEMSTRUCT lpDrawItemStruct)
{
	if (!lpDrawItemStruct) return;
	if (lpDrawItemStruct->itemID == -1) return;

	TCHAR szText[128] = {0};
	// Get and display the text for the list item.
	::SendMessage(lpDrawItemStruct->hwndItem, CB_GETLBTEXT, lpDrawItemStruct->itemID, (LPARAM) szText);
	bool bHot = 0 != (lpDrawItemStruct->itemState & (ODS_FOCUS | ODS_SELECTED));

	if (bHot)
	{
		::FillRect(lpDrawItemStruct->hDC, &lpDrawItemStruct->rcItem, GetSysBrush(cGlow));
	}
	else
	{
		::FillRect(lpDrawItemStruct->hDC, &lpDrawItemStruct->rcItem, GetSysBrush(cBackground));
	}

	DrawSegoeText(
		lpDrawItemStruct->hDC,
		szText,
		MyGetSysColor(cText),
		&lpDrawItemStruct->rcItem,
		false,
		DT_LEFT | DT_SINGLELINE | DT_VCENTER);
} // DrawComboBox

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
} // DrawItem

// Paint the status bar with double buffering to avoid flicker
void DrawStatus(
	HWND hwnd,
	int iStatusHeight,
	LPCTSTR szStatusData1,
	int iStatusData1,
	LPCTSTR szStatusData2,
	int iStatusData2,
	LPCTSTR szStatusInfo)
{
	RECT rcStatus = {0};
	::GetClientRect(hwnd, &rcStatus);
	if (rcStatus.bottom - rcStatus.top > iStatusHeight)
	{
		rcStatus.top = rcStatus.bottom - iStatusHeight;
	}
	PAINTSTRUCT ps = {0};
	::BeginPaint(hwnd, &ps);
	HDC hdc = CreateCompatibleDC(ps.hdc);
	HBITMAP hbm = CreateCompatibleBitmap(
		ps.hdc,
		rcStatus.right - rcStatus.left,
		rcStatus.bottom - rcStatus.top);
	if (hdc && hbm)
	{
		HGDIOBJ hbmOld = ::SelectObject(hdc, hbm);
		RECT rcGrad = rcStatus;
		OffsetRect(&rcGrad, -rcGrad.left, -rcGrad.top);

		GradientFillRect(hdc, rcGrad);
		RECT rcText = rcGrad;

		rcText.left = rcText.right - iStatusData2;
		::DrawSegoeText(
			hdc,
			szStatusData2,
			MyGetSysColor(cTextInverted),
			&rcText,
			true,
			DT_LEFT | DT_SINGLELINE | DT_BOTTOM);
		rcText.right = rcText.left;
		rcText.left = rcText.right - iStatusData1;
		::DrawSegoeText(
			hdc,
			szStatusData1,
			MyGetSysColor(cTextInverted),
			&rcText,
			true,
			DT_LEFT | DT_SINGLELINE | DT_BOTTOM);
		rcText.right = rcText.left;
		rcText.left = 0;
		::DrawSegoeText(
			hdc,
			szStatusInfo,
			MyGetSysColor(cTextInverted),
			&rcText,
			true,
			DT_LEFT | DT_SINGLELINE | DT_BOTTOM | DT_END_ELLIPSIS);
		BitBlt(
			ps.hdc,
			rcStatus.left,
			rcStatus.top,
			rcStatus.right - rcStatus.left,
			rcStatus.bottom - rcStatus.top,
			hdc,
			0,
			0,
			SRCCOPY);
		(void) ::SelectObject(hdc, hbmOld);
	}
	if (hdc) DeleteDC(hdc);
	if (hbm) DeleteObject(hbm);
	::EndPaint(hwnd, &ps);
} // DrawStatus

#define BORDER_VISIBLEWIDTH 2

void DrawWindowFrame(_In_ HWND hWnd, bool bActive, int iStatusHeight)
{
	HDC hdc = ::GetWindowDC(hWnd);
	if (hdc)
	{
		RECT rcWindow = {0};
		RECT rcClient = {0};
		DWORD dwWinStyle = GetWindowStyle(hWnd);
		bool bModal = (DS_MODALFRAME == (dwWinStyle & DS_MODALFRAME));
		bool bThickFrame = (WS_THICKFRAME == (dwWinStyle & WS_THICKFRAME));
		bool bMinBox = (WS_MINIMIZEBOX == (dwWinStyle & WS_MINIMIZEBOX));
		bool bMaxBox = (WS_MAXIMIZEBOX == (dwWinStyle & WS_MAXIMIZEBOX));

		::GetWindowRect(hWnd, &rcWindow); // Get our non-client size
		::GetClientRect(hWnd, &rcClient); // get our client size
		::MapWindowPoints(hWnd, NULL, (LPPOINT) &rcClient, 2); // locate our client rect on the screen
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

		int cxFixedFrame = GetSystemMetrics(SM_CXFIXEDFRAME);
		int cyFixedFrame = GetSystemMetrics(SM_CYFIXEDFRAME);
		int cxSizeFrame = GetSystemMetrics(SM_CXSIZEFRAME);
		int cySizeFrame = GetSystemMetrics(SM_CYSIZEFRAME);
		int cySize = GetSystemMetrics(SM_CYSIZE);
		int cxBorder = GetSystemMetrics(SM_CXBORDER);
		int cyBorder = GetSystemMetrics(SM_CYBORDER);

		int cxFrame = bThickFrame?cxSizeFrame:cxFixedFrame;
		int cyFrame = bThickFrame?cySizeFrame:cyFixedFrame;

		// The menu and caption have odd borders we've not painted yet - compute rectangles so we can paint them
		RECT rcFullCaption = {0};
		RECT rcIcon = {0};
		RECT rcCloseIcon = {0};
		RECT rcMaxIcon = {0};
		RECT rcMinIcon = {0};
		RECT rcMenu = {0};
		RECT rcCaptionText = {0};
		RECT rcMenuGutterLeft = {0};
		RECT rcMenuGutterRight = {0};
		RECT rcWindowGutterLeft = {0};
		RECT rcWindowGutterRight = {0};

		rcFullCaption.top = rcWindow.top + BORDER_VISIBLEWIDTH;
		rcFullCaption.left =
			rcMenuGutterLeft.left =
			rcWindowGutterLeft.left =
			rcWindow.left + BORDER_VISIBLEWIDTH;
		rcFullCaption.right =
			rcMenuGutterRight.right =
			rcWindowGutterRight.right =
			rcWindow.right - BORDER_VISIBLEWIDTH;
		rcFullCaption.bottom =
			rcMenu.top =
			rcCaptionText.bottom =
			rcMenuGutterLeft.top =
			rcMenuGutterRight.top =
			rcWindow.top + cyFrame + GetSystemMetrics(SM_CYCAPTION);

		rcIcon.top = rcWindow.top + cyFrame + GetSystemMetrics(SM_CYEDGE);
		rcIcon.left = rcWindow.left + cxFrame + cxFixedFrame;
		rcIcon.bottom = rcIcon.top + GetSystemMetrics(SM_CYSMICON);
		rcIcon.right = rcIcon.left + GetSystemMetrics(SM_CXSMICON);

		rcCloseIcon.top = rcMaxIcon.top = rcMinIcon.top = rcWindow.top + cyFrame + cyBorder;
		rcCloseIcon.bottom = rcMaxIcon.bottom = rcMinIcon.bottom = rcCloseIcon.top + cySize - 2 * cyBorder;
		rcCloseIcon.right = rcWindow.right - cxFrame - cxBorder;
		rcCloseIcon.left = rcMaxIcon.right = rcCloseIcon.right - cySize;

		rcMaxIcon.left = rcMaxIcon.right - cySize;

		rcMinIcon.right = rcMaxIcon.left + GetSystemMetrics(SM_CXEDGE);
		rcMinIcon.left = rcMinIcon.right - cySize;

		::InflateRect(&rcCloseIcon, -1, -1);
		::InflateRect(&rcMaxIcon, -1, -1);
		::InflateRect(&rcMinIcon, -1, -1);

		rcMenu.left = rcMenuGutterLeft.right = rcWindow.left + cxFrame;
		rcMenu.right = rcMenuGutterRight.left = rcWindow.right - cxFrame;
		rcMenu.bottom =
			rcMenuGutterLeft.bottom =
			rcMenuGutterRight.bottom =
			rcWindowGutterLeft.top =
			rcWindowGutterRight.top =
			rcClient.top;

		rcCaptionText.top = rcWindow.top + cxFrame;
		if (bModal)
		{
			rcCaptionText.left = rcFullCaption.left + cxFixedFrame + cxBorder;
		}
		else
		{
			rcCaptionText.left = rcIcon.right + cxFixedFrame + cxBorder;
		}
		rcCaptionText.right = rcCloseIcon.left;
		if (bMinBox) rcCaptionText.right = rcMinIcon.left;
		else if (bMaxBox) rcCaptionText.right = rcMaxIcon.left;

		rcWindowGutterLeft.bottom = rcWindowGutterRight.bottom = rcClient.bottom - iStatusHeight;

		rcWindowGutterLeft.right = rcClient.left;
		rcWindowGutterRight.left = rcClient.right;

		HGDIOBJ hpenOld = ::SelectObject(hdc, GetPen(cSolidGreyPen));
		::MoveToEx(hdc, rcMenu.left, rcMenu.bottom - 1, NULL);
		::LineTo(hdc, rcMenu.right, rcMenu.bottom - 1);
		(void) ::SelectObject(hdc,hpenOld);

		// Protect the system buttons before we paint the caption
		::ExcludeClipRect(hdc, rcCloseIcon.left, rcCloseIcon.top, rcCloseIcon.right, rcCloseIcon.bottom);
		if (bMaxBox) ::ExcludeClipRect(hdc, rcMaxIcon.left, rcMaxIcon.top, rcMaxIcon.right, rcMaxIcon.bottom);
		if (bMinBox) ::ExcludeClipRect(hdc, rcMinIcon.left, rcMinIcon.top, rcMinIcon.right, rcMinIcon.bottom);
		::FillRect(hdc, &rcFullCaption, GetSysBrush(cBackground));

		// Draw our icon
		if (!bModal)
		{
			HICON hIcon = (HICON) ::LoadImage(
				AfxGetInstanceHandle(),
				MAKEINTRESOURCE(IDR_MAINFRAME),
				IMAGE_ICON,
				rcIcon.right - rcIcon.left,
				rcIcon.bottom - rcIcon.top,
				LR_DEFAULTCOLOR);

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

		::FillRect(hdc, &rcMenuGutterLeft, GetSysBrush(cBackground));
		::FillRect(hdc, &rcMenuGutterRight, GetSysBrush(cBackground));
		::FillRect(hdc, &rcWindowGutterLeft, GetSysBrush(cBackground));
		::FillRect(hdc, &rcWindowGutterRight, GetSysBrush(cBackground));

		if (iStatusHeight)
		{
			RECT rcStatus = rcClient;
			rcStatus.top = rcClient.bottom - iStatusHeight;

			RECT rcFullStatus = {0};
			rcFullStatus.top = rcStatus.top;
			rcFullStatus.left = rcWindow.left + BORDER_VISIBLEWIDTH;
			rcFullStatus.right = rcWindow.right - BORDER_VISIBLEWIDTH;
			rcFullStatus.bottom = rcWindow.bottom - BORDER_VISIBLEWIDTH;

			::ExcludeClipRect(hdc, rcStatus.left, rcStatus.top, rcStatus.right, rcStatus.bottom);
			::FillRect(hdc, &rcFullStatus, GetSysBrush(cGlow));
		}
		else
		{
			RECT rcBottomGutter = {0};
			rcBottomGutter.top = rcWindow.bottom - cyFrame;
			rcBottomGutter.left = rcWindow.left + BORDER_VISIBLEWIDTH;
			rcBottomGutter.right = rcWindow.right - BORDER_VISIBLEWIDTH;
			rcBottomGutter.bottom = rcWindow.bottom - BORDER_VISIBLEWIDTH;

			::FillRect(hdc, &rcBottomGutter, GetSysBrush(cBackground));
		}

		TCHAR szTitle[256] = {0};
		GetWindowText(hWnd, szTitle, _countof(szTitle));
		DrawSegoeText(
			hdc,
			szTitle,
			GetSysColor(cTextInverted),
			&rcCaptionText,
			false,
			DT_LEFT | DT_SINGLELINE | DT_VCENTER);

		// Finally, we paint our border glow
		RECT rcInnerFrame = rcWindow;
		::InflateRect(&rcInnerFrame, -BORDER_VISIBLEWIDTH, -BORDER_VISIBLEWIDTH);
		::ExcludeClipRect(hdc, rcInnerFrame.left, rcInnerFrame.top, rcInnerFrame.right, rcInnerFrame.bottom);
		::FillRect(hdc, &rcWindow, GetSysBrush(bActive?cGlow:cFrameUnselected));

		::ReleaseDC(hWnd, hdc);
	}
} // DrawWindowFrame