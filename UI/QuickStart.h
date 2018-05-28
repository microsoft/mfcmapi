#pragma once

bool HandleQuickStart(_In_ WORD wMenuSelect, _In_ dialog::CMainDlg* lpHostDlg, _In_ HWND hwnd);
HRESULT OpenStoreForQuickStart(_In_ dialog::CMainDlg* lpHostDlg, _In_ HWND hwnd, _Out_ LPMDB* lppMDB);