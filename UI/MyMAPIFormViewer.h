#pragma once
// Interface for the CMyMAPIFormViewer class.

namespace controls::sortlistctrl
{
	class CContentsTableListCtrl;
} // namespace controls::sortlistctrl

namespace mapi::mapiui
{
	class CMyMAPIFormViewer : public IMAPIMessageSite, public IMAPIViewAdviseSink, public IMAPIViewContext
	{
	public:
		CMyMAPIFormViewer(
			_In_ HWND hwndParent,
			_In_ LPMDB lpMDB,
			_In_ LPMAPISESSION lpMAPISession,
			_In_ LPMAPIFOLDER lpFolder,
			_In_ LPMESSAGE lpMessage,
			_In_opt_ controls::sortlistctrl::CContentsTableListCtrl* lpContentsTableListCtrl,
			int iItem);
		virtual ~CMyMAPIFormViewer();

		// IUnknown
		STDMETHODIMP QueryInterface(REFIID riid, LPVOID* ppvObj) override;
		STDMETHODIMP_(ULONG) AddRef() override;
		STDMETHODIMP_(ULONG) Release() override;

		MAPI_GETLASTERROR_METHOD(IMPL);
		MAPI_IMAPIMESSAGESITE_METHODS(IMPL);
		MAPI_IMAPIVIEWADVISESINK_METHODS(IMPL);
		MAPI_IMAPIVIEWCONTEXT_METHODS(IMPL);

		_Check_return_ HRESULT CallDoVerb(_In_ LPMAPIFORM lpMapiForm, LONG lVerb, _In_opt_ LPCRECT lpRect);

	private:
		// helper function for ActivateNext
		_Check_return_ HRESULT GetNextMessage(
			ULONG ulDir,
			_Out_ int* piNewItem,
			_Out_ ULONG* pulStatus,
			_Deref_out_opt_ LPMESSAGE* ppMessage) const;
		_Check_return_ HRESULT SetPersist(_In_opt_ LPMAPIFORM lpForm, _In_opt_ LPPERSISTMESSAGE lpPersist);
		void ShutdownPersist();
		void ReleaseObjects();

		HWND m_hwndParent;
		LONG m_cRef;
		// primary guts of message, folder, MDB for our message site
		LPMAPIFOLDER m_lpFolder;
		LPMESSAGE m_lpMessage;
		LPMDB m_lpMDB;
		LPMAPISESSION m_lpMAPISession;
		controls::sortlistctrl::CContentsTableListCtrl* m_lpContentsTableListCtrl;
		int m_iItem; // index in list control of item being displayed
		// Set in SetAdviseSink, used in ActivateNext
		LPMAPIFORMADVISESINK m_lpMapiFormAdviseSink;
		LPPERSISTMESSAGE m_lpPersistMessage;
	};
} // namespace mapi::mapiui