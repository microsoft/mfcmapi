#include "stdafx.h"
#include <UI/Dialogs/AboutDlg.h>
#include <UI/UIFunctions.h>

void DisplayAboutDlg(_In_ CWnd* lpParentWnd)
{
	CAboutDlg AboutDlg(lpParentWnd);
	auto hRes = S_OK;
	INT_PTR iDlgRet = 0;

	EC_D_DIALOG(AboutDlg.DoModal());
}

static std::wstring CLASS = L"CAboutDlg";

CAboutDlg::CAboutDlg(
	_In_ CWnd* pParentWnd
) :CMyDialog(IDD_ABOUT, pParentWnd)
{
	TRACE_CONSTRUCTOR(CLASS);
}

CAboutDlg::~CAboutDlg()
{
	TRACE_DESTRUCTOR(CLASS);
}

BOOL CAboutDlg::OnInitDialog()
{
	auto hRes = S_OK;
	auto bRet = CMyDialog::OnInitDialog();

	auto szProductName = strings::loadstring(ID_PRODUCTNAME);
	::SetWindowTextW(m_hWnd, szProductName.c_str());

	RECT rcClient = { 0 };
	::GetClientRect(m_hWnd, &rcClient);

	auto hWndIcon = ::GetDlgItem(m_hWnd, IDC_STATIC);
	RECT rcIcon = { 0 };
	::GetWindowRect(hWndIcon, &rcIcon);
	auto iMargin = GetSystemMetrics(SM_CXHSCROLL) / 2 + 1;
	::OffsetRect(&rcIcon, iMargin - rcIcon.left, iMargin - rcIcon.top);
	::MoveWindow(hWndIcon, rcIcon.left, rcIcon.top, rcIcon.right - rcIcon.left, rcIcon.bottom - rcIcon.top, false);

	auto hWndButton = ::GetDlgItem(m_hWnd, IDOK);
	RECT rcButton = { 0 };
	::GetWindowRect(hWndButton, &rcButton);
	auto iTextHeight = GetTextHeight(m_hWnd);
	auto iCheckHeight = iTextHeight + GetSystemMetrics(SM_CYEDGE) * 2;
	::OffsetRect(&rcButton, rcClient.right - rcButton.right - iMargin, iMargin + ((IDD_ABOUTVERLAST - IDD_ABOUTVERFIRST + 1) * iTextHeight - iCheckHeight) / 2 - rcButton.top);
	::MoveWindow(hWndButton, rcButton.left, rcButton.top, rcButton.right - rcButton.left, rcButton.bottom - rcButton.top, false);

	// Position our about text with proper height
	RECT rcText = { 0 };
	rcText.left = rcIcon.right + iMargin;
	rcText.right = rcButton.left - iMargin;
	for (auto i = IDD_ABOUTVERFIRST; i <= IDD_ABOUTVERLAST; i++)
	{
		auto hWndAboutText = ::GetDlgItem(m_hWnd, i);
		rcText.top = rcIcon.top + iTextHeight * (i - IDD_ABOUTVERFIRST);
		rcText.bottom = rcText.top + iTextHeight;
		::MoveWindow(hWndAboutText, rcText.left, rcText.top, rcText.right - rcText.left, rcText.bottom - rcText.top, false);
	}

	RECT rcHelpText = { 0 };
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
		| ES_MULTILINE,
		rcHelpText,
		this,
		IDD_HELP));
	SubclassEdit(m_HelpText.m_hWnd, m_hWnd, true);

	::SendMessage(m_HelpText.m_hWnd, EM_SETEVENTMASK, NULL, ENM_LINK);
	::SendMessage(m_HelpText.m_hWnd, EM_AUTOURLDETECT, true, NULL);
	m_HelpText.SetFont(GetFont());

	auto szHelpText = strings::formatmessage(IDS_HELPTEXT, szProductName.c_str());
	::SetWindowTextW(m_HelpText.m_hWnd, szHelpText.c_str());

	auto rcCheck = rcHelpText;
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
	m_DisplayAboutCheck.SetCheck(RegKeys[regkeyDISPLAY_ABOUT_DIALOG].ulCurDWORD ? BST_CHECKED : BST_UNCHECKED);
	auto szDisplayAboutCheck = strings::loadstring(IDS_DISPLAYABOUTCHECK);
	::SetWindowTextW(m_DisplayAboutCheck.m_hWnd, szDisplayAboutCheck.c_str());

	// Get version information from the application.
	WCHAR szFullPath[256];
	DWORD dwRet = 0;
	EC_D(dwRet, GetModuleFileNameW(NULL, szFullPath, _countof(szFullPath)));
	DWORD dwVerInfoSize = 0;
	EC_D(dwVerInfoSize, GetFileVersionInfoSizeW(szFullPath, nullptr));
	if (dwVerInfoSize)
	{
		// If we were able to get the information, process it.
		auto pbData = new BYTE[dwVerInfoSize];

		if (pbData)
		{
			EC_D(bRet, GetFileVersionInfoW(szFullPath,
				0, dwVerInfoSize, static_cast<void*>(pbData)));

			struct LANGANDCODEPAGE {
				WORD wLanguage;
				WORD wCodePage;
			} *lpTranslate = { nullptr };

			UINT cbTranslate = 0;

			// Read the list of languages and code pages.
			EC_B(VerQueryValueW(
				pbData,
				L"\\VarFileInfo\\Translation", // STRING_OK
				reinterpret_cast<LPVOID*>(&lpTranslate),
				&cbTranslate));

			// Read the file description for the first language/codepage
			if (S_OK == hRes && lpTranslate)
			{
				auto szGetName = strings::format(
					L"\\StringFileInfo\\%04x%04x\\", // STRING_OK
					lpTranslate[0].wLanguage,
					lpTranslate[0].wCodePage);

				// Walk through the dialog box items that we want to replace.
				if (!FAILED(hRes)) for (auto i = IDD_ABOUTVERFIRST; i <= IDD_ABOUTVERLAST; i++)
				{
					UINT cchVer = 0;
					WCHAR* pszVer = nullptr;
					WCHAR szResult[256];

					hRes = S_OK;
					UINT uiRet = 0;
					EC_D(uiRet, ::GetDlgItemTextW(m_hWnd, i, szResult, _countof(szResult)));

					EC_B(VerQueryValueW(
						static_cast<void*>(pbData),
						(szGetName + szResult).c_str(),
						reinterpret_cast<void**>(&pszVer),
						&cchVer));

					if (S_OK == hRes && cchVer && pszVer)
					{
						// Replace the dialog box item text with version information.
						::SetDlgItemTextW(m_hWnd, i, pszVer);
					}
				}
			}

			delete[] pbData;
		}
	}

	return bRet;
}

LRESULT CAboutDlg::WindowProc(UINT message, WPARAM wParam, LPARAM lParam)
{
	switch (message)
	{
	case WM_HELP:
		return true;
	case WM_NOTIFY:
	{
		auto lpel = reinterpret_cast<ENLINK*>(lParam);
		if (lpel)
		{
			switch (lpel->nmhdr.code)
			{
			case EN_LINK:
			{
				if (WM_LBUTTONUP == lpel->msg)
				{
					TCHAR szLink[128] = { 0 };
					TEXTRANGE tr = { 0 };
					tr.lpstrText = szLink;
					tr.chrg = lpel->chrg;
					::SendMessage(lpel->nmhdr.hwndFrom, EM_GETTEXTRANGE, NULL, reinterpret_cast<LPARAM>(&tr));
					ShellExecute(nullptr, _T("open"), szLink, nullptr, nullptr, SW_SHOWNORMAL); // STRING_OK
				}

				return NULL;
			}
			}
		}
		break;
	}
	case WM_ERASEBKGND:
	{
		RECT rect = { 0 };
		::GetClientRect(m_hWnd, &rect);
		auto hOld = ::SelectObject(reinterpret_cast<HDC>(wParam), GetSysBrush(cBackground));
		auto bRet = ::PatBlt(reinterpret_cast<HDC>(wParam), 0, 0, rect.right - rect.left, rect.bottom - rect.top, PATCOPY);
		::SelectObject(reinterpret_cast<HDC>(wParam), hOld);
		return bRet;
	}
	}

	return CMyDialog::WindowProc(message, wParam, lParam);
}

void CAboutDlg::OnOK()
{
	CMyDialog::OnOK();
	auto iCheckState = m_DisplayAboutCheck.GetCheck();

	if (BST_CHECKED == iCheckState)
		RegKeys[regkeyDISPLAY_ABOUT_DIALOG].ulCurDWORD = true;
	else
		RegKeys[regkeyDISPLAY_ABOUT_DIALOG].ulCurDWORD = false;
}