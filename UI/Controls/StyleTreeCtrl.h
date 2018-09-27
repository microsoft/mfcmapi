#pragma once
#include <Enums.h>

namespace controls
{
	class StyleTreeCtrl : public CTreeCtrl
	{
	public:
		STDMETHODIMP_(ULONG) AddRef();
		STDMETHODIMP_(ULONG) Release();

	private:
		LONG m_cRef{1};
	};
} // namespace controls