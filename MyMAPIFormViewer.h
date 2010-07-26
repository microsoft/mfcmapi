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
	_Check_return_ STDMETHODIMP QueryInterface (REFIID riid, _Deref_out_opt_ LPVOID * ppvObj);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	_Check_return_ STDMETHODIMP GetLastError(
		HRESULT hResult,
		ULONG ulFlags,
		_Deref_out_ LPMAPIERROR FAR * lppMAPIError);

	// IMAPIMessageSite
	_Check_return_ STDMETHODIMP GetSession (
		_Deref_out_opt_ LPMAPISESSION FAR * ppSession);
	_Check_return_ STDMETHODIMP GetStore (
		_Deref_out_opt_ LPMDB FAR * ppStore);
	_Check_return_ STDMETHODIMP GetFolder (
		_Deref_out_opt_ LPMAPIFOLDER FAR * ppFolder);
	_Check_return_ STDMETHODIMP GetMessage (
		_Deref_out_opt_ LPMESSAGE FAR * ppmsg);
	_Check_return_ STDMETHODIMP GetFormManager (
		_Deref_out_ LPMAPIFORMMGR FAR * ppFormMgr);
	_Check_return_ STDMETHODIMP NewMessage (
		ULONG fComposeInFolder,
		_In_ LPMAPIFOLDER pFolderFocus,
		_In_ LPPERSISTMESSAGE pPersistMessage,
		_Deref_out_opt_ LPMESSAGE FAR * ppMessage,
		_Deref_out_opt_ LPMAPIMESSAGESITE FAR * ppMessageSite,
		_Deref_out_opt_ LPMAPIVIEWCONTEXT FAR * ppViewContext);
	_Check_return_ STDMETHODIMP CopyMessage (
		_In_ LPMAPIFOLDER pFolderDestination);
	_Check_return_ STDMETHODIMP MoveMessage (
		_In_ LPMAPIFOLDER pFolderDestination,
		_In_ LPMAPIVIEWCONTEXT pViewContext,
		_In_ LPCRECT prcPosRect);
	_Check_return_ STDMETHODIMP DeleteMessage (
		_In_ LPMAPIVIEWCONTEXT pViewContext,
		_In_ LPCRECT prcPosRect);
	_Check_return_ STDMETHODIMP SaveMessage();
	_Check_return_ STDMETHODIMP SubmitMessage(
		ULONG ulFlags);
	_Check_return_ STDMETHODIMP GetSiteStatus(
		_Inout_ LPULONG lpulStatus);

	// IMAPIViewAdviseSink
	_Check_return_ STDMETHODIMP OnShutdown();
	_Check_return_ STDMETHODIMP OnNewMessage();
	_Check_return_ STDMETHODIMP OnPrint(
		ULONG dwPageNumber,
		HRESULT hrStatus);
	_Check_return_ STDMETHODIMP OnSubmitted();
	_Check_return_ STDMETHODIMP OnSaved();

	// IMAPIViewContext
	_Check_return_ STDMETHODIMP SetAdviseSink(
		_In_ LPMAPIFORMADVISESINK pmvns);
	_Check_return_ STDMETHODIMP ActivateNext(
		ULONG ulDir,
		_In_ LPCRECT prcPosRect);
	_Check_return_ STDMETHODIMP GetPrintSetup(
		ULONG ulFlags,
		_Deref_out_ LPFORMPRINTSETUP FAR * lppFormPrintSetup);
	_Check_return_ STDMETHODIMP GetSaveStream(
		_Out_ ULONG FAR * pulFlags,
		_Out_ ULONG FAR * pulFormat,
		_Deref_out_ LPSTREAM FAR * ppstm);
	_Check_return_ STDMETHODIMP GetViewStatus(
		_In_ LPULONG lpulStatus);

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
