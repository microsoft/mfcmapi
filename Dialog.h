#pragma once
// Dialog.h : Base class for dialog class customization
#include "ParentWnd.h"

class CMyDialog : public CDialog
{
public:
	CMyDialog();
	CMyDialog(UINT nIDTemplate, CWnd* pParentWnd = NULL);
	virtual ~CMyDialog();
	void DisplayParentedDialog(CParentWnd* lpNonModalParent, UINT iAutoCenterWidth);

protected:
	_Check_return_ LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	_Check_return_ virtual BOOL CheckAutoCenter();
	void SetStatusHeight(int iHeight);
	int GetStatusHeight();

private:
	void Constructor();
	void OnMeasureItem(int nIDCtl, _In_ LPMEASUREITEMSTRUCT lpMeasureItemStruct);
	void OnDrawItem(int nIDCtl, _In_ LPDRAWITEMSTRUCT lpDrawItemStruct);
	CParentWnd* m_lpNonModalParent;
	CWnd* m_hwndCenteringWindow;
	UINT m_iAutoCenterWidth;
	bool m_bStatus;
	int m_iStatusHeight;

	DECLARE_MESSAGE_MAP()
};