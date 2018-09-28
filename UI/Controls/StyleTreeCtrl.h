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
	};
} // namespace controls