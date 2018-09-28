#pragma once
#include <Enums.h>

namespace controls
{
	class StyleTreeCtrl : public CTreeCtrl
	{
	public:
		void Create(_In_ CWnd* pCreateParent, UINT nIDContextMenu);

		STDMETHODIMP_(ULONG) AddRef();
		STDMETHODIMP_(ULONG) Release();

		// TODO: Make this private
		UINT m_nIDContextMenu{0};

	private:
		LONG m_cRef{1};
	};
} // namespace controls