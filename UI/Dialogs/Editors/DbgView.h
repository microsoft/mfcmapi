#pragma once
#include <UI/ParentWnd.h>

namespace editor
{
	void DisplayDbgView(_In_ CParentWnd* pParentWnd);
	void OutputToDbgView(const std::wstring& szMsg);
}