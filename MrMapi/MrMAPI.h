#pragma once
#define ulNoMatch 0xffffffff
#include <MrMapi/cli.h>
#include <Interpret/ExtraPropTags.h>
#include <MrMapi/MMMapiMime.h>

_Check_return_ LPMAPISESSION MrMAPILogonEx(const std::wstring& lpszProfile);
_Check_return_ LPMDB OpenExchangeOrDefaultMessageStore(_In_ LPMAPISESSION lpMAPISession);