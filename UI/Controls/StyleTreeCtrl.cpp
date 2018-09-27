#include <StdAfx.h>
#include <UI/Controls/StyleTreeCtrl.h>

namespace controls
{
	static std::wstring CLASS = L"StyleTreeCtrl";

	STDMETHODIMP_(ULONG) StyleTreeCtrl::AddRef()
	{
		const auto lCount = InterlockedIncrement(&m_cRef);
		TRACE_ADDREF(CLASS, lCount);
		return lCount;
	}

	STDMETHODIMP_(ULONG) StyleTreeCtrl::Release()
	{
		const auto lCount = InterlockedDecrement(&m_cRef);
		TRACE_RELEASE(CLASS, lCount);
		if (!lCount) delete this;
		return lCount;
	}

} // namespace controls