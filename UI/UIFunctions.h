#pragma once
// Common UI functions for MFC MAPI

#include <oleacc.h>

namespace ui
{
	enum class uiColor
	{
		Background,
		BackgroundReadOnly,
		Glow,
		GlowBackground,
		GlowText,
		FrameSelected,
		FrameUnselected,
		SelectedBackground,
		Arrow,
		Text,
		TextDisabled,
		TextReadOnly,
		BitmapTransBack,
		BitmapTransFore,
		Status,
		StatusText,
		PaneHeaderBackground,
		PaneHeaderText,
		TextHighlightBackground,
		TextHighlight,
		TestPink,
		TestLavender,
		UIEnd
	};

	enum class uiPen
	{
		SolidPen,
		SolidGreyPen,
		DashedPen,
		PenEnd
	};

	enum class uiBitmap
	{
		Notify,
		Close,
		Minimize,
		Maximize,
		Restore,
		BitmapEnd
	};

	enum class uiButtonStyle
	{
		Unstyled,
		UpArrow,
		DownArrow
	};

	enum class uiLabelStyle
	{
		Unstyled,
		PaneHeaderLabel,
		PaneHeaderText,
	};

	void InitializeGDI() noexcept;
	void UninitializeGDI() noexcept;

	void UpdateMenuString(_In_ HWND hWnd, UINT uiMenuTag, UINT uidNewString);

	void DisplayContextMenu(UINT uiClassMenu, UINT uiControlMenu, _In_ HWND hWnd, int x, int y);

	HMENU LocateSubmenu(_In_ HMENU hMenu, UINT uid);

	_Check_return_ int GetEditHeight(_In_opt_ HWND hwndEdit);
	_Check_return_ int GetTextHeight(_In_opt_ HWND hwndEdit);
	SIZE GetTextExtentPoint32(HDC hdc, const std::wstring& szText);

	HFONT GetSegoeFont();
	HFONT GetSegoeFontBold();

	_Check_return_ HBRUSH GetSysBrush(uiColor uc);
	_Check_return_ COLORREF MyGetSysColor(uiColor uc);

	_Check_return_ HPEN GetPen(uiPen up);

	HBITMAP GetBitmap(uiBitmap ub);

	struct SCALE
	{
		LONG x;
		LONG y;
		LONG denominator;
	};

	SCALE GetDPIScale();
	HBITMAP ScaleBitmap(HBITMAP hBitmap, const SCALE& scale);

	void ClearEditFormatting(_In_ HWND hWnd, bool bReadOnly);

	// Application-specific owner-drawn menu info struct. Owner-drawn data
	// is a pointer to one of these. MSAAMENUINFO must be the first
	// member.
	struct MenuEntry
	{
		MSAAMENUINFO m_MSAA{}; // MSAA info - must be first element.
		std::wstring m_pName; // Menu text for display. Empty for separator item.
		BOOL m_bOnMenuBar{};
		ULONG_PTR m_AddInData{};
	};
	typedef MenuEntry* LPMENUENTRY;

	_Check_return_ LPMENUENTRY CreateMenuEntry(_In_ const std::wstring& szMenu);
	void ConvertMenuOwnerDraw(_In_ HMENU hMenu, bool bRoot);
	void DeleteMenuEntries(_In_ HMENU hMenu);

	// Edit box
	LRESULT CALLBACK DrawEditProc(
		_In_ HWND hWnd,
		UINT uMsg,
		WPARAM wParam,
		LPARAM lParam,
		UINT_PTR uIdSubclass,
		DWORD_PTR /*dwRefData*/);
	void SubclassEdit(_In_ HWND hWnd, _In_opt_ HWND hWndParent, bool bReadOnly);

	// List
	void CustomDrawList(_In_ LPNMLVCUSTOMDRAW lvcd, _In_ LRESULT* pResult, DWORD_PTR iItemCurHover);
	void DrawListItemGlow(_In_ HWND hWnd, UINT itemID);

	// Tree
	void DrawBitmap(_In_ HDC hdc, _In_ const RECT& rcTarget, uiBitmap iBitmap, bool bHover, int offset = 0);
	void CustomDrawTree(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult, bool bHover, _In_ HTREEITEM hItemCurHover);
	void DrawTreeItemGlow(_In_ HWND hWnd, _In_ HTREEITEM hItem);
	void DrawExpandTriangle(_In_ HWND hWnd, _In_ HDC hdc, _In_ HTREEITEM hItem, bool bGlow, bool bHover);

	// Header
	void DrawTriangle(_In_ HWND hWnd, _In_ HDC hdc, _In_ const RECT& rc, bool bButton, bool bUp);
	void CustomDrawHeader(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
	void DrawTrackingBar(_In_ HWND hWndHeader, _In_ HWND hWndList, int x, int iHeaderHeight, bool bRedraw);

	// Menu and Combo box
	bool MeasureItem(_In_ LPMEASUREITEMSTRUCT lpMeasureItemStruct);
	bool DrawItem(_In_ LPDRAWITEMSTRUCT lpDrawItemStruct);
	std::wstring GetLBText(HWND hwnd, int nIndex);

	// Status Bar
	void DrawStatus(
		HWND hwnd,
		int iStatusHeight,
		const std::wstring& szStatusData1,
		int iStatusData1,
		const std::wstring& szStatusData2,
		int iStatusData2,
		const std::wstring& szStatusInfo);

	void GetCaptionRects(
		HWND hWnd,
		RECT* lprcFullCaption,
		RECT* lprcIcon,
		RECT* lprcCloseIcon,
		RECT* lprcMaxIcon,
		RECT* lprcMinIcon,
		RECT* lprcCaptionText) noexcept;
	void DrawSystemButtons(_In_ HWND hWnd, _In_opt_ HDC hdc, LONG_PTR iHitTest, bool bHover);
	void DrawWindowFrame(_In_ HWND hWnd, bool bActive, int iStatusHeight);

	// Winproc handler for custom controls
	_Check_return_ bool HandleControlUI(UINT message, WPARAM wParam, LPARAM lParam, _Out_ LRESULT* lpResult);

	// Help Text
	void DrawHelpText(_In_ HWND hWnd, _In_ UINT uIDText);

	// Text labels
	void SubclassLabel(_In_ HWND hWnd);
	void StyleLabel(_In_ HWND hWnd, uiLabelStyle lsStyle);

	void StyleButton(_In_ HWND hWnd, uiButtonStyle bsStyle);
	void DrawSegoeTextW(
		_In_ HDC hdc,
		_In_ const std::wstring& lpchText,
		_In_ COLORREF color,
		_In_ const RECT& rc,
		bool bBold,
		_In_ UINT format);

	_Check_return_ HWND GetMainWindow();
} // namespace ui