#include <StdAfx.h>
#include <UI/Controls/StyleTreeCtrl.h>
#include <UI/UIFunctions.h>

namespace controls
{
	static std::wstring CLASS = L"StyleTreeCtrl";

	void StyleTreeCtrl::Create(_In_ CWnd* pCreateParent, const UINT nIDContextMenu)
	{
		m_nIDContextMenu = nIDContextMenu;

		CTreeCtrl::Create(
			TVS_HASBUTTONS | TVS_LINESATROOT | TVS_EDITLABELS | TVS_DISABLEDRAGDROP | TVS_SHOWSELALWAYS |
				TVS_FULLROWSELECT | WS_CHILD | WS_TABSTOP | WS_CLIPSIBLINGS | WS_VISIBLE,
			CRect(0, 0, 0, 0),
			pCreateParent,
			IDC_FOLDER_TREE);
		TreeView_SetBkColor(m_hWnd, ui::MyGetSysColor(ui::cBackground));
		TreeView_SetTextColor(m_hWnd, ui::MyGetSysColor(ui::cText));
		::SendMessageA(m_hWnd, WM_SETFONT, reinterpret_cast<WPARAM>(ui::GetSegoeFont()), false);
	}
} // namespace controls