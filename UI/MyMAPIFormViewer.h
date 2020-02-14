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

		STDMETHODIMP GetLastError(HRESULT hResult, ULONG ulFlags, LPMAPIERROR* lppMAPIError) override;

		// IMAPIMessageSite
		STDMETHODIMP GetSession(LPMAPISESSION* ppSession) override;
		STDMETHODIMP GetStore(LPMDB* ppStore) override;
		STDMETHODIMP GetFolder(LPMAPIFOLDER* ppFolder) override;
		STDMETHODIMP GetMessage(LPMESSAGE* ppmsg) override;
		STDMETHODIMP GetFormManager(LPMAPIFORMMGR* ppFormMgr) override;
		STDMETHODIMP NewMessage(
			ULONG fComposeInFolder,
			LPMAPIFOLDER pFolderFocus,
			LPPERSISTMESSAGE pPersistMessage,
			LPMESSAGE* ppMessage,
			LPMAPIMESSAGESITE* ppMessageSite,
			LPMAPIVIEWCONTEXT* ppViewContext) override;
		STDMETHODIMP CopyMessage(LPMAPIFOLDER pFolderDestination) override;
		STDMETHODIMP
		MoveMessage(LPMAPIFOLDER pFolderDestination, LPMAPIVIEWCONTEXT pViewContext, LPCRECT prcPosRect) override;
		STDMETHODIMP DeleteMessage(LPMAPIVIEWCONTEXT pViewContext, LPCRECT prcPosRect) override;
		STDMETHODIMP SaveMessage() override;
		STDMETHODIMP SubmitMessage(ULONG ulFlags) override;
		STDMETHODIMP GetSiteStatus(LPULONG lpulStatus) override;

		// IMAPIViewAdviseSink
		STDMETHODIMP OnShutdown() override;
		STDMETHODIMP OnNewMessage() override;
		STDMETHODIMP OnPrint(ULONG dwPageNumber, HRESULT hrStatus) override;
		STDMETHODIMP OnSubmitted() override;
		STDMETHODIMP OnSaved() override;

		// IMAPIViewContext
		STDMETHODIMP SetAdviseSink(LPMAPIFORMADVISESINK pmvns) override;
		STDMETHODIMP ActivateNext(ULONG ulDir, LPCRECT lpRect) override;
		STDMETHODIMP GetPrintSetup(ULONG ulFlags, LPFORMPRINTSETUP* lppFormPrintSetup) override;
		STDMETHODIMP GetSaveStream(ULONG* pulFlags, ULONG* pulFormat, LPSTREAM* ppstm) override;
		STDMETHODIMP GetViewStatus(LPULONG lpulStatus) override;

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