#pragma once
// Base class for dialog class customization
#include <UI/ParentWnd.h>

namespace dialog
{
	class CMyDialog : public CDialog
	{
	public:
		CMyDialog();
		CMyDialog(UINT nIDTemplate, CWnd* pParentWnd = nullptr);
		virtual ~CMyDialog();
		void DisplayParentedDialog(ui::CParentWnd* lpNonModalParent, UINT iAutoCenterWidth);

	protected:
		LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;
		BOOL CheckAutoCenter() override;
		void SetStatusHeight(int iHeight);
		int GetStatusHeight() const;

	private:
		void Constructor();
		LRESULT NCHitTest(WPARAM wParam, LPARAM lParam);
		ui::CParentWnd* m_lpNonModalParent;
		CWnd* m_hwndCenteringWindow;
		UINT m_iAutoCenterWidth;
		bool m_bStatus;
		int m_iStatusHeight;
		HWND m_hWndPrevious;
	};
}