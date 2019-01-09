#pragma once

namespace ui
{
	namespace addinui
	{
		_Check_return_ ULONG ExtendAddInMenu(HMENU hMenu, ULONG ulAddInContext);
		_Check_return_ LPMENUITEM GetAddinMenuItem(HWND hWnd, UINT uidMsg);
		void InvokeAddInMenu(_In_opt_ LPADDINMENUPARAMS lpParams);
	} // namespace addinui
} // namespace ui