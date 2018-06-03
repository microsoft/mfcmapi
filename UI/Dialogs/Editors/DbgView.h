#pragma once
#include <UI/ParentWnd.h>

namespace dialog
{
	namespace editor
	{
		void DisplayDbgView(_In_ ui::CParentWnd* pParentWnd);
		void OutputToDbgView(const std::wstring& szMsg);
	}
}