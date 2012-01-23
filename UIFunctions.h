#pragma once
// UIFunctions.h : Common UI functions for MFC MAPI

enum uiColor
{
	cBackground,
	cBackgroundReadOnly,
	cGlow,
	cGlowBackground,
	cGlowText,
	cFrameSelected,
	cFrameUnselected,
	cArrow,
	cText,
	cTextDisabled,
	cTextReadOnly,
	cBitmapTransBack,
	cBitmapTransFore,
	cStatus,
	cStatusText,
	cUIEnd
};

enum uiPen
{
	cSolidPen,
	cSolidGreyPen,
	cDashedPen,
	cPenEnd
};

enum uiBitmap
{
	cNotify,
	cClose,
	cMinimize,
	cMaximize,
	cRestore,
	cBitmapEnd
};

class CDoubleBuffer
{
private:
	HDC m_hdcMem;
	HBITMAP m_hbmpMem;
	HDC m_hdcPaint;
	RECT m_rcPaint;

	void Cleanup();

public:
	CDoubleBuffer();
	~CDoubleBuffer();
	void Begin(_Inout_ HDC& hdc, _In_ RECT* prcPaint);
	void End(_Inout_ HDC& hdc);
};

void InitializeGDI();
void UninitializeGDI();

void UpdateMenuString(_In_ HWND hWnd, UINT uiMenuTag, UINT uidNewString);

void DisplayContextMenu(UINT uiClassMenu, UINT uiControlMenu, _In_ HWND hWnd, int x, int y);

_Check_return_ int GetEditHeight(_In_ HWND hwndEdit);
_Check_return_ int GetTextHeight(_In_ HWND hwndEdit);

HFONT GetSegoeFont();
HFONT GetSegoeFontBold();

_Check_return_ HBRUSH GetSysBrush(uiColor uc);
_Check_return_ COLORREF MyGetSysColor(uiColor uc);

_Check_return_ HPEN GetPen(uiPen up);

HBITMAP GetBitmap(uiBitmap ub);

void ClearEditFormatting(_In_ HWND hWnd, bool bReadOnly);

// Application-specific owner-drawn menu info struct. Owner-drawn data 
// is a pointer to one of these. MSAAMENUINFO must be the first 
// member. 
struct MenuEntry
{
	MSAAMENUINFO m_MSAA; // MSAA info - must be first element.
	LPWSTR       m_pName; // Menu text for display. NULL for separator item.
	BOOL         b_OnMenuBar;
	ULONG_PTR    m_AddInData;
};
typedef MenuEntry* LPMENUENTRY;

_Check_return_ LPMENUENTRY CreateMenuEntry(_In_z_ LPWSTR szMenu);
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

// List
void CustomDrawList(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult, int iItemCur);
void DrawListItemFrame(_In_ HWND hWnd, _In_opt_ HDC hdc, UINT itemID, UINT itemState);

// Tree
void CustomDrawTree(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult, bool bHover, _In_ HTREEITEM hItemCurHover);
void DrawTreeItemFrame(_In_ HWND hWnd, _In_ HTREEITEM hItem, bool bDraw);
void DrawExpandTriangle(_In_ HWND hWnd, _In_ HDC hdc, _In_ HTREEITEM hItem, bool bGlow);

// Header
void CustomDrawHeader(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);
void DrawTrackingBar(_In_ HWND hWndHeader, _In_ HWND hWndList, int x, int iHeaderHeight, bool bRedraw);

// Menu and Combo box
void MeasureItem(_In_ LPMEASUREITEMSTRUCT lpMeasureItemStruct);
void DrawItem(_In_ LPDRAWITEMSTRUCT lpDrawItemStruct);

// Status Bar
void DrawStatus(
	HWND hwnd,
	int iStatusHeight,
	LPCTSTR szStatusData1,
	int iStatusData1,
	LPCTSTR szStatusData2,
	int iStatusData2,
	LPCTSTR szStatusInfo);

void DrawSystemButtons(_In_ HWND hWnd, _In_opt_ HDC hdc, int iHitTest);
void DrawWindowFrame(_In_ HWND hWnd, bool bActive, int iStatusHeight);

// Winproc handler for custom controls
_Check_return_ bool HandleControlUI(UINT message, WPARAM wParam, LPARAM lParam, _Out_ LRESULT* lpResult);

// Help Text
void DrawHelpText(_In_ HWND hWnd, _In_ UINT uIDText);