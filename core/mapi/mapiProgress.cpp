#include <core/stdafx.h>
#include <core/mapi/mapiProgress.h>

namespace mapi::mapiui
{
	std::function<LPMAPIPROGRESS(const std::wstring& lpszContext, HWND hWnd)> getMAPIProgress;

	_Check_return_ LPMAPIPROGRESS GetMAPIProgress(const std::wstring& lpszContext, _In_ HWND hWnd)
	{
		return getMAPIProgress ? getMAPIProgress(lpszContext, hWnd) : nullptr;
	}
} // namespace mapi::mapiui