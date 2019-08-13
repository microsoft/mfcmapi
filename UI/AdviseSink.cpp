#include <StdAfx.h>
#include <UI/AdviseSink.h>
#include <core/mapi/mapiOutput.h>
#include <core/utility/output.h>
#include <UI/OnNotify.h>

namespace mapi
{
	namespace mapiui
	{
		static std::wstring CLASS = L"CAdviseSink";

		CAdviseSink::CAdviseSink(_In_ HWND hWndParent, _In_opt_ HTREEITEM hTreeParent)
		{
			TRACE_CONSTRUCTOR(CLASS);
			m_cRef = 1;
			m_hWndParent = hWndParent;
			m_hTreeParent = hTreeParent;
			m_lpAdviseTarget = nullptr;
		}

		CAdviseSink::~CAdviseSink()
		{
			TRACE_DESTRUCTOR(CLASS);
			if (m_lpAdviseTarget) m_lpAdviseTarget->Release();
		}

		STDMETHODIMP CAdviseSink::QueryInterface(REFIID riid, LPVOID* ppvObj)
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

		STDMETHODIMP_(ULONG) CAdviseSink::AddRef()
		{
			const auto lCount = InterlockedIncrement(&m_cRef);
			TRACE_ADDREF(CLASS, lCount);
			return lCount;
		}

		STDMETHODIMP_(ULONG) CAdviseSink::Release()
		{
			const auto lCount = InterlockedDecrement(&m_cRef);
			TRACE_RELEASE(CLASS, lCount);
			if (!lCount) delete this;
			return lCount;
		}

		STDMETHODIMP_(ULONG) CAdviseSink::OnNotify(ULONG cNotify, LPNOTIFICATION lpNotifications)
		{
			output::outputNotifications(output::DBGNotify, nullptr, cNotify, lpNotifications, m_lpAdviseTarget);
			mapi::mapiui::OnNotify(m_hWndParent, m_hTreeParent, cNotify, lpNotifications);
			return S_OK;
		}

		void CAdviseSink::SetAdviseTarget(LPMAPIPROP lpProp)
		{
			if (m_lpAdviseTarget) m_lpAdviseTarget->Release();
			m_lpAdviseTarget = lpProp;
			if (m_lpAdviseTarget) m_lpAdviseTarget->AddRef();
		}
	} // namespace mapiui
} // namespace mapi