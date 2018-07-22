#pragma once

namespace dialog
{
	bool HandleQuickStart(_In_ WORD wMenuSelect, _In_ dialog::CMainDlg* lpHostDlg, _In_ HWND hwnd);
	LPMDB OpenStoreForQuickStart(_In_ dialog::CMainDlg* lpHostDlg, _In_ HWND hwnd);
}