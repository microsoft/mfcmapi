#include <core/stdafx.h>
#include <core/utility/clipboard.h>

namespace clipboard
{
	const void CopyTo(_In_ const std::wstring& str) noexcept
	{
		if (str.empty()) return;
		OpenClipboard(nullptr);
		EmptyClipboard();
		const auto cb = (str.size() + 1) * sizeof(wchar_t);
		const auto hg = GlobalAlloc(GMEM_MOVEABLE, cb);
		if (!hg)
		{
			CloseClipboard();
			return;
		}

		const auto mem = GlobalLock(hg);
		if (!mem)
		{
			CloseClipboard();
			return;
		}

		memcpy(mem, str.c_str(), cb);
		GlobalUnlock(hg);
		SetClipboardData(CF_UNICODETEXT, hg);
		CloseClipboard();
		GlobalFree(hg);
	}
} // namespace clipboard