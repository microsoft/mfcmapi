#pragma once
#include <UI/ParentWnd.h>

namespace dialog::editor
{
	void DisplayDbgView(_In_ ui::CParentWnd* pParentWnd);
	void OutputToDbgView(const std::wstring& szMsg);
} // namespace dialog::editor