#pragma once
// AboutDlg.h : header file
//

void DisplayAboutDlg(CWnd* lpParentWnd);

class CAboutDlg : public CDialog
{
public:
	CAboutDlg(
		CWnd* pParentWnd
		);
	~CAboutDlg();

	BOOL OnInitDialog();

protected:
	HICON					m_hIcon;
	CRichEditCtrl 			m_HelpText;
	CButton	 				m_DisplayAboutCheck;

	void	OnOK();
	LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam);
	DECLARE_MESSAGE_MAP()
};
