#pragma once
// SortHeader.h : header file for our custom sort list control header

struct HeaderData
{
	ULONG ulTagArrayRow;
	ULONG ulPropTag;
	bool bIsAB;
	wstring szTipString;
};
typedef HeaderData* LPHEADERDATA;

class CSortHeader : public CHeaderCtrl
{
public:
	CSortHeader();
	_Check_return_ bool Init(_In_ CHeaderCtrl *pHeader, _In_ HWND hwndParent);

private:
	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	void RegisterHeaderTooltip();

	// Custom messages
	_Check_return_ LRESULT msgOnSaveColumnOrder(WPARAM wParam, LPARAM lParam);
	void OnCustomDraw(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);

	HWND m_hwndTip;
	TOOLINFOW m_ti;
	HWND m_hwndParent;
	bool m_bTooltipDisplayed;

	DECLARE_MESSAGE_MAP()
};