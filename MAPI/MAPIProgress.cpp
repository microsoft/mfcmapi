#include <StdAfx.h>
#include <MAPI/MAPIProgress.h>

namespace mapi
{
	namespace mapiui
	{
		std::function<LPMAPIPROGRESS(const std::wstring& lpszContext, HWND hWnd)> getMAPIProgress;

		_Check_return_ LPMAPIPROGRESS GetMAPIProgress(const std::wstring& lpszContext, _In_ HWND hWnd)
		{
			return getMAPIProgress ? getMAPIProgress(lpszContext, hWnd) : nullptr;
		}
	} // namespace mapiui
} // namespace mapi