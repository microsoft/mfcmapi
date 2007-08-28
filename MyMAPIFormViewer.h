// MyMAPIFormViewer.h: interface for the CMyMAPIFormViewer class.
//
//////////////////////////////////////////////////////////////////////

#pragma once

class CContentsTableListCtrl;

class CMyMAPIFormViewer :
	public IMAPIMessageSite ,
	public IMAPIViewAdviseSink,
	public IMAPIViewContext
{
public:
	CMyMAPIFormViewer(
		HWND hwndParent,
		LPMDB lpMDB,
		LPMAPISESSION lpMAPISession,
		LPMAPIFOLDER lpFolder,
		LPMESSAGE lpMessage,
		CContentsTableListCtrl *lpContentsTableListCtrl,
		int iItem);
	virtual ~CMyMAPIFormViewer();

	STDMETHODIMP QueryInterface (REFIID riid, LPVOID * ppvObj);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	STDMETHODIMP GetLastError(
		HRESULT hResult,
		ULONG ulFlags,
		LPMAPIERROR FAR * lppMAPIError);

	////////////////////////////////////////////////////////////
	// IMAPIMessageSite Functions
	////////////////////////////////////////////////////////////
	STDMETHODIMP GetSession (
		LPMAPISESSION FAR * ppSession);
	STDMETHODIMP GetStore (
		LPMDB FAR * ppStore);
	STDMETHODIMP GetFolder (
		LPMAPIFOLDER FAR * ppFolder);
	STDMETHODIMP GetMessage (
		LPMESSAGE FAR * ppmsg);
	STDMETHODIMP GetFormManager (
		LPMAPIFORMMGR FAR * ppFormMgr);
	STDMETHODIMP NewMessage (
		ULONG fComposeInFolder,
		LPMAPIFOLDER pFolderFocus,
		LPPERSISTMESSAGE pPersistMessage,
		LPMESSAGE FAR * ppMessage,
		LPMAPIMESSAGESITE FAR * ppMessageSite,
		LPMAPIVIEWCONTEXT FAR * ppViewContext);
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
	////////////////////////////////////////////////////////////
	// IMAPIViewAdviseSink Functions
	////////////////////////////////////////////////////////////
	STDMETHODIMP OnShutdown();
	STDMETHODIMP OnNewMessage();
	STDMETHODIMP OnPrint(
		ULONG dwPageNumber,
		HRESULT hrStatus);
	STDMETHODIMP OnSubmitted();
	STDMETHODIMP OnSaved();

	////////////////////////////////////////////////////////////
	// IMAPIViewContext Functions
	////////////////////////////////////////////////////////////
	STDMETHODIMP SetAdviseSink(
		LPMAPIFORMADVISESINK pmvns);
	STDMETHODIMP ActivateNext(
		ULONG ulDir,
		LPCRECT prcPosRect);
	STDMETHODIMP GetPrintSetup(
		ULONG ulFlags,
		LPFORMPRINTSETUP FAR * lppFormPrintSetup);
	STDMETHODIMP GetSaveStream(
		ULONG FAR * pulFlags,
		ULONG FAR * pulFormat,
		LPSTREAM FAR * ppstm);
	STDMETHODIMP GetViewStatus(
		LPULONG lpulStatus);

	HRESULT CallDoVerb(LPMAPIFORM lpMapiForm,
				   LONG lVerb,
				   LPCRECT lpRect);
private :
	HWND					m_hwndParent;
	LONG					m_cRef;

	//primary guts of message, folder, MDB for our message site
	LPMAPIFOLDER			m_lpFolder;
	LPMESSAGE				m_lpMessage;
	LPMDB					m_lpMDB;
	LPMAPISESSION			m_lpMAPISession;
	CContentsTableListCtrl*	m_lpContentsTableListCtrl;
	int						m_iItem;//index in list control of item being displayed
	//Set in SetAdviseSink, used in ActivateNext
	LPMAPIFORMADVISESINK	m_lpMapiFormAdviseSink;

	//helper function for ActivateNext
	HRESULT GetNextMessage(
		ULONG ulDir,
		int* piNewItem,
		ULONG* pulStatus,
		LPMESSAGE* ppMessage);
	HRESULT SetPersist(LPMAPIFORM lpForm, LPPERSISTMESSAGE lpPersist);
//	LPMAPIFORM				m_lpForm;
	LPPERSISTMESSAGE		m_lpPersistMessage;
//	ULONG					m_pulConnection;
	void ShutdownPersist();
	void ReleaseObjects();
};
