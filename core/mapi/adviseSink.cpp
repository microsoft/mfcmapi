#include <core/stdafx.h>
#include <core/mapi/adviseSink.h>
#include <core/mapi/mapiOutput.h>
#include <core/utility/output.h>

namespace mapi
{
	std::function<void(HWND hWndParent, HTREEITEM hTreeParent, ULONG cNotify, LPNOTIFICATION lpNotifications)>
		onNotifyCallback;

	static std::wstring CLASS = L"adviseSink";

	adviseSink::adviseSink(_In_ HWND hWndParent, _In_opt_ HTREEITEM hTreeParent)
		: m_hWndParent(hWndParent), m_hTreeParent(hTreeParent)
	{
		TRACE_CONSTRUCTOR(CLASS);
	}

	adviseSink::~adviseSink()
	{
		TRACE_DESTRUCTOR(CLASS);
		if (m_lpAdviseTarget) m_lpAdviseTarget->Release();
	}

	STDMETHODIMP adviseSink::QueryInterface(REFIID riid, LPVOID* ppvObj)
	{
		*ppvObj = nullptr;
		if (riid == IID_IMAPIAdviseSink || riid == IID_IUnknown)
		{
			*ppvObj = static_cast<LPVOID>(this);
			AddRef();
			return S_OK;
		}

		return E_NOINTERFACE;
	}

	STDMETHODIMP_(ULONG) adviseSink::AddRef()
	{
		const auto lCount = InterlockedIncrement(&m_cRef);
		TRACE_ADDREF(CLASS, lCount);
		return lCount;
	}

	STDMETHODIMP_(ULONG) adviseSink::Release()
	{
		const auto lCount = InterlockedDecrement(&m_cRef);
		TRACE_RELEASE(CLASS, lCount);
		if (!lCount) delete this;
		return lCount;
	}

	STDMETHODIMP_(ULONG) adviseSink::OnNotify(ULONG cNotify, LPNOTIFICATION lpNotifications)
	{
		output::outputNotifications(output::DBGNotify, nullptr, cNotify, lpNotifications, m_lpAdviseTarget);
		if (onNotifyCallback)
		{
			onNotifyCallback(m_hWndParent, m_hTreeParent, cNotify, lpNotifications);
		}

		return S_OK;
	}

	void adviseSink::SetAdviseTarget(LPMAPIPROP lpProp) noexcept
	{
		if (m_lpAdviseTarget) m_lpAdviseTarget->Release();
		m_lpAdviseTarget = lpProp;
		if (m_lpAdviseTarget) m_lpAdviseTarget->AddRef();
	}
} // namespace mapi