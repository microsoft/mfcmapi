#pragma once

namespace mapi
{
	namespace mapiui
	{
		extern std::function<
			void(HWND hWndParent, HTREEITEM hTreeParent, ULONG cNotify, LPNOTIFICATION lpNotifications)>
			onNotifyCallback;

		class CAdviseSink : public IMAPIAdviseSink
		{
		public:
			CAdviseSink(_In_ HWND hWndParent, _In_opt_ HTREEITEM hTreeParent);
			~CAdviseSink();

			STDMETHODIMP QueryInterface(REFIID riid, LPVOID* ppvObj) override;
			STDMETHODIMP_(ULONG) AddRef() override;
			STDMETHODIMP_(ULONG) Release() override;
			STDMETHODIMP_(ULONG) OnNotify(ULONG cNotify, LPNOTIFICATION lpNotifications) override;

			void SetAdviseTarget(LPMAPIPROP lpProp);

		private:
			LONG m_cRef{1};
			const HWND m_hWndParent{};
			const HTREEITEM m_hTreeParent{};
			LPMAPIPROP m_lpAdviseTarget{}; // Used only for named prop lookups in debug logging.
		};
	} // namespace mapiui
} // namespace mapi