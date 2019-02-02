#pragma once

// TODO: Consider rewriting this with callbacks and avoid the windows message passing
namespace mapi
{
	namespace mapiui
	{
		class CAdviseSink : public IMAPIAdviseSink
		{
		public:
			CAdviseSink(_In_ HWND hWndParent, _In_opt_ HTREEITEM hTreeParent);
			virtual ~CAdviseSink();

			STDMETHODIMP QueryInterface(REFIID riid, LPVOID* ppvObj) override;
			STDMETHODIMP_(ULONG) AddRef() override;
			STDMETHODIMP_(ULONG) Release() override;
			STDMETHODIMP_(ULONG) OnNotify(ULONG cNotify, LPNOTIFICATION lpNotifications) override;

			void SetAdviseTarget(LPMAPIPROP lpProp);

		private:
			LONG m_cRef;
			HWND m_hWndParent;
			HTREEITEM m_hTreeParent;
			LPMAPIPROP m_lpAdviseTarget; // Used only for named prop lookups in debug logging.
		};
	} // namespace mapiui
} // namespace mapi