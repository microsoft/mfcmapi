#pragma once
// AboutDlg.h : Displays an about dialog

void DisplayAboutDlg(CWnd* lpParentWnd);

class CAboutDlg : public CDialog
{
public:
	CAboutDlg(
		CWnd* pParentWnd
		);
	virtual ~CAboutDlg();

private:
	void	OnOK();
	LRESULT	WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	BOOL	OnInitDialog();

	HICON			m_hIcon;
	CRichEditCtrl	m_HelpText;
	CButton	 		m_DisplayAboutCheck;
};
