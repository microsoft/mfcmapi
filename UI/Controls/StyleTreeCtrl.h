#pragma once
#include <Enums.h>

namespace controls
{
	class StyleTreeCtrl : public CTreeCtrl
	{
	public:
		void Create(_In_ CWnd* pCreateParent, UINT nIDContextMenu);

		// TODO: Make this private
		UINT m_nIDContextMenu{0};

	private:
		// Overrides from base class
		LRESULT WindowProc(UINT message, WPARAM wParam, LPARAM lParam) override;

		void OnCustomDraw(_In_ NMHDR* pNMHDR, _In_ LRESULT* pResult);

		HTREEITEM m_hItemCurHover{nullptr};
		bool m_HoverButton{false};

		// TODO: Kill this and use WindowProc instead
		DECLARE_MESSAGE_MAP()
	};
} // namespace controls