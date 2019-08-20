#pragma once
_Check_return_ LPMAPISESSION MrMAPILogonEx(const std::wstring& lpszProfile);
_Check_return_ LPMDB OpenExchangeOrDefaultMessageStore(_In_ LPMAPISESSION lpMAPISession);