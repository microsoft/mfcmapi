#pragma once
// SortHeader.h : header file for our custom sort list control header

#include "InterpretProp.h"

typedef struct _HeaderData	FAR * LPHEADERDATA;

typedef struct _HeaderData
{
	ULONG	ulTagArrayRow;
	ULONG	ulPropTag;
	BOOL	bIsAB;
	TCHAR	szTipString[TAG_MAX_LEN];
} HeaderData;

class CSortHeader : public CHeaderCtrl
{
public:
	CSortHeader();
	_Check_return_ BOOL Init(_In_ CHeaderCtrl *pHeader, _In_ HWND hwndParent);

private:
	_Check_return_ LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	void    RegisterHeaderTooltip();

	// Custom messages
	_Check_return_ LRESULT msgOnSaveColumnOrder(WPARAM wParam, LPARAM lParam);

	HWND		m_hwndTip;
	TOOLINFO	m_ti;
	HWND		m_hwndParent;
	BOOL		m_bTooltipDisplayed;

	DECLARE_MESSAGE_MAP()
};