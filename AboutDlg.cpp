// AboutDlg.cpp : implementation file
//

#include "stdafx.h"
#include "AboutDlg.h"
#include "ParentWnd.h"
#include "UIFunctions.h"

void DisplayAboutDlg(_In_ CWnd* lpParentWnd)
{
	CAboutDlg AboutDlg(lpParentWnd);
	HRESULT hRes = S_OK;
	INT_PTR iDlgRet = 0;

	EC_D_DIALOG(AboutDlg.DoModal());
} // DisplayAboutDlg

static TCHAR* CLASS = _T("CAboutDlg");

CAboutDlg::CAboutDlg(
					 _In_ CWnd* pParentWnd
					 ):CMyDialog(IDD_ABOUT,pParentWnd)
{
	TRACE_CONSTRUCTOR(CLASS);
} // CAboutDlg::CAboutDlg

CAboutDlg::~CAboutDlg()
{
	TRACE_DESTRUCTOR(CLASS);
} // CAboutDlg::~CAboutDlg

BOOL CAboutDlg::OnInitDialog()
{
	HRESULT hRes = S_OK;
	BOOL bRet = CMyDialog::OnInitDialog();
	TCHAR szFullPath[256];
	DWORD dwVerHnd = 0;
	DWORD dwRet = 0;
	DWORD dwVerInfoSize = 0;
	int iTextHeight = GetTextHeight(m_hWnd);
	int iCheckHeight = iTextHeight + GetSystemMetrics(SM_CYEDGE)*2;
	int iMargin = GetSystemMetrics(SM_CXHSCROLL)/2+1;
	int i = 0;
	int iRet = 0;
	static TCHAR szProductName[128];

	WC_D(iRet,LoadString(
		GetModuleHandle(NULL),
		ID_PRODUCTNAME,
		szProductName,
		_countof(szProductName)));

	SetWindowText(szProductName);

	RECT rcClient = {0};
	RECT rcIcon = {0};
	RECT rcButton = {0};
	RECT rcText = {0};
	RECT rcHelpText = {0};
	RECT rcCheck = {0};

	::GetClientRect(m_hWnd, &rcClient);

	HWND hWndIcon = ::GetDlgItem(m_hWnd, IDC_STATIC);
	::GetWindowRect(hWndIcon, &rcIcon);
	::OffsetRect(&rcIcon, iMargin - rcIcon.left, iMargin - rcIcon.top);
	::MoveWindow(hWndIcon, rcIcon.left, rcIcon.top, rcIcon.right - rcIcon.left, rcIcon.bottom - rcIcon.top, false);

	HWND hWndButton = ::GetDlgItem(m_hWnd, IDOK);
	::GetWindowRect(hWndButton, &rcButton);
	::OffsetRect(&rcButton, rcClient.right - rcButton.right - iMargin, iMargin + ((IDD_ABOUTVERLAST - IDD_ABOUTVERFIRST + 1) * iTextHeight - iCheckHeight) / 2 - rcButton.top);
	::MoveWindow(hWndButton, rcButton.left, rcButton.top, rcButton.right - rcButton.left, rcButton.bottom - rcButton.top, false);

	// Position our about text with proper height
	rcText.left = rcIcon.right + iMargin;
	rcText.right = rcButton.left - iMargin;
	for (i = IDD_ABOUTVERFIRST; i <= IDD_ABOUTVERLAST; i++)
	{
		HWND hWndAboutText = ::GetDlgItem(m_hWnd, i);
		rcText.top = rcIcon.top + iTextHeight * (i - IDD_ABOUTVERFIRST);
		rcText.bottom = rcText.top + iTextHeight;
		::MoveWindow(hWndAboutText, rcText.left, rcText.top, rcText.right - rcText.left, rcText.bottom - rcText.top, false);
	}

	rcHelpText.left = rcClient.left + iMargin;
	rcHelpText.top = rcText.bottom + iMargin * 2;
	rcHelpText.right = rcClient.right - iMargin;
	rcHelpText.bottom = rcClient.bottom - iMargin * 2 - iCheckHeight;

	EC_B(m_HelpText.Create(
		WS_TABSTOP
		| WS_CHILD
		| WS_VISIBLE
		| WS_HSCROLL
		| WS_VSCROLL
		| WS_BORDER
		| ES_AUTOHSCROLL
		| ES_AUTOVSCROLL
		| ES_MULTILINE
		| ES_READONLY,
		rcHelpText,
		this,
		IDD_HELP));
	m_HelpText.ModifyStyleEx(WS_EX_CLIENTEDGE, 0, 0);
	::SendMessage(m_HelpText.m_hWnd, EM_SETEVENTMASK,NULL, ENM_LINK);
	::SendMessage(m_HelpText.m_hWnd, EM_AUTOURLDETECT,true, NULL);
	m_HelpText.SetBackgroundColor(false, MyGetSysColor(cBackground));
	m_HelpText.SetFont(GetFont());
	ClearEditFormatting(m_HelpText.m_hWnd, false); // Lie. This control is read only, but we render it as if it wasn't.

	CString szHelpText;
	szHelpText.FormatMessage(IDS_HELPTEXT,szProductName);
	m_HelpText.SetWindowText(szHelpText);

	rcCheck = rcHelpText;
	rcCheck.top = rcHelpText.bottom + iMargin;
	rcCheck.bottom = rcCheck.top + iCheckHeight;

	EC_B(m_DisplayAboutCheck.Create(
		NULL,
		WS_TABSTOP
		| WS_CHILD
		| WS_CLIPSIBLINGS
		| WS_VISIBLE
		| BS_AUTOCHECKBOX,
		rcCheck,
		this,
		IDD_DISPLAYABOUT));
	m_DisplayAboutCheck.SetCheck(RegKeys[regkeyDISPLAY_ABOUT_DIALOG].ulCurDWORD?BST_CHECKED:BST_UNCHECKED);
	CString szDisplayAboutCheck;
	EC_B(szDisplayAboutCheck.LoadString(IDS_DISPLAYABOUTCHECK));
	m_DisplayAboutCheck.SetWindowText(szDisplayAboutCheck);

	// Get version information from the application.
	EC_D(dwRet,GetModuleFileName(NULL, szFullPath, _countof(szFullPath)));
	EC_D(dwVerInfoSize,GetFileVersionInfoSize(szFullPath, &dwVerHnd));
	if (dwVerInfoSize)
	{
		// If we were able to get the information, process it.
		size_t cchRoot = 0;

		BYTE* pbData = NULL;
		pbData = new BYTE[dwVerInfoSize];

		if (pbData)
		{
			EC_D(bRet,GetFileVersionInfo(szFullPath,
				dwVerHnd, dwVerInfoSize, (void*)pbData));

			struct LANGANDCODEPAGE {
				WORD wLanguage;
				WORD wCodePage;
			} *lpTranslate = {0};

			UINT	cbTranslate = 0;
			TCHAR	szGetName[256];

			// Read the list of languages and code pages.
			EC_B(VerQueryValue(
				pbData,
				_T("\\VarFileInfo\\Translation"), // STRING_OK
				(LPVOID*)&lpTranslate,
				&cbTranslate));

			// Read the file description for the first language/codepage
			if (S_OK == hRes && lpTranslate)
			{
				EC_H(StringCchPrintf(
					szGetName,
					_countof(szGetName),
					_T("\\StringFileInfo\\%04x%04x\\"), // STRING_OK
					lpTranslate[0].wLanguage,
					lpTranslate[0].wCodePage));

				EC_H(StringCchLength(szGetName,256,&cchRoot));

				// Walk through the dialog box items that we want to replace.
				if (!FAILED(hRes)) for (i = IDD_ABOUTVERFIRST; i <= IDD_ABOUTVERLAST; i++)
				{
					UINT  cchVer = 0;
					TCHAR*pszVer = NULL;
					TCHAR szResult[256];

					hRes = S_OK;

					UINT uiRet = 0;
					EC_D(uiRet,GetDlgItemText(i, szResult, _countof(szResult)));
					EC_H(StringCchCopy(&szGetName[cchRoot], _countof(szGetName)-cchRoot,szResult));

					EC_B(VerQueryValue(
						(void*)pbData,
						szGetName,
						(void**)&pszVer,
						&cchVer));

					if (S_OK == hRes && cchVer && pszVer)
					{
						// Replace the dialog box item text with version information.
						EC_H(StringCchCopy(szResult, _countof(szResult), pszVer));

						SetDlgItemText(i, szResult);
					}
				}
			}
			delete[] pbData;
		}
	}
	return bRet;
} // CAboutDlg::OnInitDialog

LRESULT CAboutDlg::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
	case WM_HELP:
		return true;
		break;
	case WM_NOTIFY:
		{
			ENLINK* lpel = (ENLINK*) lParam;
			if (lpel)
			{
				switch (lpel->nmhdr.code)
				{
				case EN_LINK:
					{
						if (WM_LBUTTONUP == lpel->msg)
						{
							TCHAR szLink[128] = {0};
							TEXTRANGE tr = {0};
							tr.lpstrText = szLink;
							tr.chrg = lpel->chrg;
							::SendMessage(lpel->nmhdr.hwndFrom, EM_GETTEXTRANGE, NULL, (LPARAM) &tr);
							ShellExecute(NULL, _T("open"), szLink, NULL, NULL, SW_SHOWNORMAL); // STRING_OK
						}
						return NULL;
					}
					break;
				}
			}
			break;
		}
	case WM_ERASEBKGND:
		{
			RECT rect = {0};
			::GetClientRect(m_hWnd, &rect);
			HGDIOBJ hOld = ::SelectObject((HDC)wParam, GetSysBrush(cBackground));
			BOOL bRet = ::PatBlt((HDC) wParam, 0, 0, rect.right - rect.left, rect.bottom - rect.top, PATCOPY);
			::SelectObject((HDC)wParam, hOld);
			return bRet;
		}
		break;
	} // end switch
	return CMyDialog::WindowProc(message,wParam,lParam);
} // CAboutDlg::WindowProc

void CAboutDlg::OnOK()
{
	CMyDialog::OnOK();
	int iCheckState = m_DisplayAboutCheck.GetCheck();

	if (BST_CHECKED == iCheckState)
		RegKeys[regkeyDISPLAY_ABOUT_DIALOG].ulCurDWORD = true;
	else
		RegKeys[regkeyDISPLAY_ABOUT_DIALOG].ulCurDWORD = false;
} // CAboutDlg::OnOK