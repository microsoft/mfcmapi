#pragma once

namespace mapi::mapiui
{
	extern std::function<LPMAPIPROGRESS(const std::wstring& lpszContext, HWND hWnd)> getMAPIProgress;
	_Check_return_ LPMAPIPROGRESS GetMAPIProgress(const std::wstring& lpszContext, _In_ HWND hWnd);
} // namespace mapi::mapiui