#include <StdAfx.h>
#include <UI/AdviseSink.h>
#include <core/mapi/mapiOutput.h>
#include <core/utility/output.h>

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

			if (!m_hWndParent) return 0;

			for (ULONG i = 0; i < cNotify; i++)
			{
				if (fnevNewMail == lpNotifications[i].ulEventType)
				{
					MessageBeep(MB_OK);
				}
				else if (fnevTableModified == lpNotifications[i].ulEventType)
				{
					switch (lpNotifications[i].info.tab.ulTableEvent)
					{
					case TABLE_ERROR:
					case TABLE_CHANGED:
					case TABLE_RELOAD:
						EC_H_S(static_cast<HRESULT>(::SendMessage(
							m_hWndParent, WM_MFCMAPI_REFRESHTABLE, reinterpret_cast<WPARAM>(m_hTreeParent), 0)));
						break;
					case TABLE_ROW_ADDED:
						EC_H_S(static_cast<HRESULT>(::SendMessage(
							m_hWndParent,
							WM_MFCMAPI_ADDITEM,
							reinterpret_cast<WPARAM>(&lpNotifications[i].info.tab),
							reinterpret_cast<LPARAM>(m_hTreeParent))));
						break;
					case TABLE_ROW_DELETED:
						EC_H_S(static_cast<HRESULT>(::SendMessage(
							m_hWndParent,
							WM_MFCMAPI_DELETEITEM,
							reinterpret_cast<WPARAM>(&lpNotifications[i].info.tab),
							reinterpret_cast<LPARAM>(m_hTreeParent))));
						break;
					case TABLE_ROW_MODIFIED:
						EC_H_S(static_cast<HRESULT>(::SendMessage(
							m_hWndParent,
							WM_MFCMAPI_MODIFYITEM,
							reinterpret_cast<WPARAM>(&lpNotifications[i].info.tab),
							reinterpret_cast<LPARAM>(m_hTreeParent))));
						break;
					}
				}
			}

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