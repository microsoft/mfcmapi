#pragma once
// MyMAPIFormViewer.h: interface for the CMyMAPIFormViewer class.

class CContentsTableListCtrl;

class CMyMAPIFormViewer :
	public IMAPIMessageSite ,
	public IMAPIViewAdviseSink,
	public IMAPIViewContext
{
public:
	CMyMAPIFormViewer(
		_In_ HWND hwndParent,
		_In_ LPMDB lpMDB,
		_In_ LPMAPISESSION lpMAPISession,
		_In_ LPMAPIFOLDER lpFolder,
		_In_ LPMESSAGE lpMessage,
		_In_opt_ CContentsTableListCtrl* lpContentsTableListCtrl,
		int iItem);
	virtual ~CMyMAPIFormViewer();

	// IUnknown
	STDMETHODIMP QueryInterface (REFIID riid, LPVOID * ppvObj);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	STDMETHODIMP GetLastError(
		HRESULT hResult,
		ULONG ulFlags,
		LPMAPIERROR* lppMAPIError);

	// IMAPIMessageSite
	STDMETHODIMP GetSession (
		LPMAPISESSION* ppSession);
	STDMETHODIMP GetStore (
		LPMDB* ppStore);
	STDMETHODIMP GetFolder (
		LPMAPIFOLDER* ppFolder);
	STDMETHODIMP GetMessage (
		LPMESSAGE* ppmsg);
	STDMETHODIMP GetFormManager (
		LPMAPIFORMMGR* ppFormMgr);
	STDMETHODIMP NewMessage (
		ULONG fComposeInFolder,
		LPMAPIFOLDER pFolderFocus,
		LPPERSISTMESSAGE pPersistMessage,
		LPMESSAGE* ppMessage,
		LPMAPIMESSAGESITE* ppMessageSite,
		LPMAPIVIEWCONTEXT* ppViewContext);
	STDMETHODIMP CopyMessage (
		LPMAPIFOLDER pFolderDestination);
	STDMETHODIMP MoveMessage (
		LPMAPIFOLDER pFolderDestination,
		LPMAPIVIEWCONTEXT pViewContext,
		LPCRECT prcPosRect);
	STDMETHODIMP DeleteMessage (
		LPMAPIVIEWCONTEXT pViewContext,
		LPCRECT prcPosRect);
	STDMETHODIMP SaveMessage();
	STDMETHODIMP SubmitMessage(
		ULONG ulFlags);
	STDMETHODIMP GetSiteStatus(
		LPULONG lpulStatus);

	// IMAPIViewAdviseSink
	STDMETHODIMP OnShutdown();
	STDMETHODIMP OnNewMessage();
	STDMETHODIMP OnPrint(
		ULONG dwPageNumber,
		HRESULT hrStatus);
	STDMETHODIMP OnSubmitted();
	STDMETHODIMP OnSaved();

	// IMAPIViewContext
	STDMETHODIMP SetAdviseSink(
		LPMAPIFORMADVISESINK pmvns);
	STDMETHODIMP ActivateNext(
		ULONG ulDir,
		LPCRECT prcPosRect);
	STDMETHODIMP GetPrintSetup(
		ULONG ulFlags,
		LPFORMPRINTSETUP* lppFormPrintSetup);
	STDMETHODIMP GetSaveStream(
		ULONG* pulFlags,
		ULONG* pulFormat,
		LPSTREAM* ppstm);
	STDMETHODIMP GetViewStatus(
		LPULONG lpulStatus);

	_Check_return_ HRESULT CallDoVerb(_In_ LPMAPIFORM lpMapiForm,
		LONG lVerb,
		_In_opt_ LPCRECT lpRect);

private:
	// helper function for ActivateNext
	_Check_return_ HRESULT GetNextMessage(
		ULONG ulDir,
		_Out_ int* piNewItem,
		_Out_ ULONG* pulStatus,
		_Deref_out_opt_ LPMESSAGE* ppMessage);
	_Check_return_ HRESULT SetPersist(_In_opt_ LPMAPIFORM lpForm, _In_opt_ LPPERSISTMESSAGE lpPersist);
	void ShutdownPersist();
	void ReleaseObjects();

	HWND					m_hwndParent;
	LONG					m_cRef;
	// primary guts of message, folder, MDB for our message site
	LPMAPIFOLDER			m_lpFolder;
	LPMESSAGE				m_lpMessage;
	LPMDB					m_lpMDB;
	LPMAPISESSION			m_lpMAPISession;
	CContentsTableListCtrl*	m_lpContentsTableListCtrl;
	int						m_iItem; // index in list control of item being displayed
	// Set in SetAdviseSink, used in ActivateNext
	LPMAPIFORMADVISESINK	m_lpMapiFormAdviseSink;
	LPPERSISTMESSAGE		m_lpPersistMessage;
};
