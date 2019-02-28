#pragma once

namespace dialog
{
	class CMainDlg;
	bool HandleQuickStart(_In_ WORD wMenuSelect, _In_ CMainDlg* lpHostDlg, _In_ HWND hwnd);
	LPMDB OpenStoreForQuickStart(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd);
}