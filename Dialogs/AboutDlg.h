#pragma once
// AboutDlg.h : Displays an about dialog
#include <Dialogs/Dialog.h>

void DisplayAboutDlg(_In_ CWnd* lpParentWnd);

class CAboutDlg : public CMyDialog
{
public:
	CAboutDlg(
		_In_ CWnd* pParentWnd
		);
	virtual ~CAboutDlg();

private:
	void    OnOK();
	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	BOOL    OnInitDialog();

	CRichEditCtrl m_HelpText;
	CButton m_DisplayAboutCheck;
};
