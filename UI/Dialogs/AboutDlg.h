#pragma once
// Displays an about dialog
#include <UI/Dialogs/Dialog.h>

namespace dialog
{
	void DisplayAboutDlg(_In_ CWnd* lpParentWnd);

	class CAboutDlg : public CMyDialog
	{
	public:
		CAboutDlg(_In_ CWnd* pParentWnd);
		virtual ~CAboutDlg();

	private:
		void OnOK() override;
		LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;
		BOOL OnInitDialog() override;

		CRichEditCtrl m_HelpText;
		CButton m_DisplayAboutCheck;
	};
} // namespace dialog