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
	BOOL Init(CHeaderCtrl *pHeader, HWND hwndParent);

private:
	LRESULT	WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	void	RegisterHeaderTooltip();

	// Custom messages
	LRESULT	msgOnSaveColumnOrder(WPARAM wParam, LPARAM lParam);

	HWND		m_hwndTip;
	TOOLINFO	m_ti;
	HWND		m_hwndParent;

	DECLARE_MESSAGE_MAP()
};