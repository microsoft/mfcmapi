// MyMAPIFormViewer.cpp: implementation of the CMyMAPIFormViewer class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "MyMAPIFormViewer.h"
#include "MAPIFunctions.h"
#include "MAPIFormFunctions.h"
#include "ContentsTableListCtrl.h"
#include "Editor.h"
#include "InterpretProp.h"
#include "InterpretProp2.h"
#include "MyWinApp.h"
extern CMyWinApp theApp;

static TCHAR* CLASS = _T("CMyMAPIFormViewer");

//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CMyMAPIFormViewer::CMyMAPIFormViewer(
									 HWND hwndParent,
									 LPMDB lpMDB,
									 LPMAPISESSION lpMAPISession,
									 LPMAPIFOLDER lpFolder,
									 LPMESSAGE lpMessage,
									 CContentsTableListCtrl *lpContentsTableListCtrl,
									 int iItem)
{
	TRACE_CONSTRUCTOR(CLASS);

	m_cRef = 1;
	m_iItem = iItem;

	m_lpMDB = lpMDB;
	m_lpMAPISession = lpMAPISession;
	m_lpFolder = lpFolder;
	m_lpMessage = lpMessage;
	m_lpContentsTableListCtrl = lpContentsTableListCtrl;

	if (m_lpMDB) m_lpMDB->AddRef();
	if (m_lpMAPISession) m_lpMAPISession->AddRef();
	if (m_lpFolder) m_lpFolder->AddRef();
	if (m_lpMessage) m_lpMessage->AddRef();
	if (m_lpContentsTableListCtrl) m_lpContentsTableListCtrl->AddRef();

	m_lpPersistMessage = NULL;
	m_lpMapiFormAdviseSink = NULL;

	m_hwndParent = hwndParent;
	if (!m_hwndParent)
	{
		m_hwndParent = GetActiveWindow();
	}
	if (!m_hwndParent && theApp.m_pMainWnd)
	{
		m_hwndParent = theApp.m_pMainWnd->m_hWnd;
	}
	if (!m_hwndParent)
	{
		ErrDialog(__FILE__,__LINE__,IDS_EDFORMVIEWERNULLPARENT);
		DebugPrint(DBGFormViewer,_T("Form Viewer created with a NULL parent!\n"));
	}
}

CMyMAPIFormViewer::~CMyMAPIFormViewer()
{
	TRACE_DESTRUCTOR(CLASS);
	ReleaseObjects();
}

void CMyMAPIFormViewer::ReleaseObjects()
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("ReleaseObjects"),_T("\n"));
	ShutdownPersist();
	if (m_lpMessage) m_lpMessage->Release();
	if (m_lpFolder) m_lpFolder->Release();
	if (m_lpMDB) m_lpMDB->Release();
	if (m_lpMAPISession) m_lpMAPISession->Release();
	if (m_lpMapiFormAdviseSink) m_lpMapiFormAdviseSink->Release();
	if (m_lpContentsTableListCtrl) m_lpContentsTableListCtrl->Release(); // this must be last!!!!!
	m_lpMessage = NULL;
	m_lpFolder = NULL;
	m_lpMDB = NULL;
	m_lpMAPISession = NULL;
	m_lpMapiFormAdviseSink = NULL;
	m_lpContentsTableListCtrl = NULL;
}

STDMETHODIMP CMyMAPIFormViewer::QueryInterface (REFIID riid,
												 LPVOID * ppvObj)
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("QueryInterface"),_T("\n"));
	LPTSTR szGuid = GUIDToStringAndName((LPGUID)&riid);
	if (szGuid)
	{
		DebugPrint(DBGFormViewer,_T("GUID Requested: %s\n"),szGuid);
		delete[] szGuid;
		szGuid = NULL;
	}

	*ppvObj = 0;
	if (riid == IID_IMAPIMessageSite)
	{
		DebugPrint(DBGFormViewer,_T("Requested IID_IMAPIMessageSite\n"));
		*ppvObj = (IMAPIMessageSite *)this;
		AddRef();
		return S_OK;
	}
	if (riid == IID_IMAPIViewContext)
	{
		DebugPrint(DBGFormViewer,_T("Requested IID_IMAPIViewContext\n"));
		*ppvObj = (IMAPIViewContext *)this;
		AddRef();
		return S_OK;
	}
	if (riid == IID_IMAPIViewAdviseSink)
	{
		DebugPrint(DBGFormViewer,_T("Requested IID_IMAPIViewAdviseSink\n"));
		*ppvObj = (IMAPIViewAdviseSink *)this;
		AddRef();
		return S_OK;
	}
	if (riid == IID_IUnknown)
	{
		DebugPrint(DBGFormViewer,_T("Requested IID_IUnknown\n"));
		*ppvObj = (LPUNKNOWN)((IMAPIMessageSite *)this);
		AddRef();
		return S_OK;
	}
	DebugPrint(DBGFormViewer,_T("Unknown interface requested\n"));
	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) CMyMAPIFormViewer::AddRef()
{
	LONG lCount = InterlockedIncrement(&m_cRef);
	TRACE_ADDREF(CLASS,lCount);
	return lCount;
}

STDMETHODIMP_(ULONG) CMyMAPIFormViewer::Release()
{
	LONG lCount = InterlockedDecrement(&m_cRef);
	TRACE_RELEASE(CLASS,lCount);
	if (!lCount)  delete this;
	return lCount;
}

///////////////////////////////////////////////////////////////////////////////
// IMAPIMessageSite implementation
///////////////////////////////////////////////////////////////////////////////
STDMETHODIMP CMyMAPIFormViewer::GetSession (LPMAPISESSION FAR * ppSession)
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("GetSession"),_T("\n"));
	if (ppSession) *ppSession = m_lpMAPISession;
	if (ppSession && *ppSession)
	{
		m_lpMAPISession->AddRef();
		return S_OK;
	}
	else
	{
		return S_FALSE;
	}
}

STDMETHODIMP CMyMAPIFormViewer::GetStore (LPMDB FAR * ppStore)
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("GetStore"),_T("\n"));
	if (ppStore) *ppStore = m_lpMDB;
	if (ppStore && *ppStore)
	{
		m_lpMDB->AddRef();
		return S_OK;
	}
	else
	{
		return S_FALSE;
	}
}

STDMETHODIMP CMyMAPIFormViewer::GetFolder (LPMAPIFOLDER FAR * ppFolder)
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("GetFolder"),_T("\n"));
	if (ppFolder) *ppFolder = m_lpFolder;
	if (ppFolder && *ppFolder)
	{
		m_lpFolder->AddRef();
		return S_OK;
	}
	else
	{
		return S_FALSE;
	}
}

STDMETHODIMP CMyMAPIFormViewer::GetMessage(LPMESSAGE FAR * ppmsg)
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("GetMessage"),_T("\n"));
	if (ppmsg) *ppmsg = m_lpMessage;
	if (ppmsg && *ppmsg)
	{
		m_lpMessage->AddRef();
		return S_OK;
	}
	else
	{
		return S_FALSE;
	}
}

STDMETHODIMP CMyMAPIFormViewer::GetFormManager(LPMAPIFORMMGR FAR * ppFormMgr)
{
	HRESULT hRes = S_OK;
	DebugPrintEx(DBGFormViewer,CLASS,_T("GetFormManager"),_T("\n"));
	EC_H(MAPIOpenFormMgr(m_lpMAPISession,ppFormMgr));
	return hRes;
}

STDMETHODIMP CMyMAPIFormViewer::NewMessage(ULONG fComposeInFolder,
											 LPMAPIFOLDER pFolderFocus,
											 LPPERSISTMESSAGE pPersistMessage,
											 LPMESSAGE FAR * ppMessage,
											 LPMAPIMESSAGESITE FAR * ppMessageSite,
											 LPMAPIVIEWCONTEXT FAR * ppViewContext)
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("NewMessage"),_T("fComposeInFolder = 0x%X pFolderFocus = 0x%X, pPersistMessage = 0x%X\n"),fComposeInFolder,pFolderFocus,pPersistMessage);
	HRESULT hRes = S_OK;

	*ppMessage = NULL;
	*ppMessageSite = NULL;
	if (ppViewContext) *ppViewContext = NULL;

	if ((fComposeInFolder == FALSE) || !pFolderFocus)
	{
		pFolderFocus = m_lpFolder;
	}

	if (pFolderFocus)
	{
		EC_H(pFolderFocus->CreateMessage(
			NULL, // IID
			NULL, // flags
			ppMessage));

		if (*ppMessage) // not going to release this because we're returning it
		{
			CMyMAPIFormViewer *lpMAPIFormViewer = NULL; // don't free since we're passing it back
			lpMAPIFormViewer = new CMyMAPIFormViewer(
				m_hwndParent,
				m_lpMDB,
				m_lpMAPISession,
				pFolderFocus,
				*ppMessage,
				NULL, // m_lpContentsTableListCtrl, // don't need this on a new message
				-1);
			if (lpMAPIFormViewer) // not going to release this because we're returning it in ppMessageSite
			{
				EC_H(lpMAPIFormViewer->SetPersist(NULL,pPersistMessage));
				*ppMessageSite = (LPMAPIMESSAGESITE)lpMAPIFormViewer;
			}
		}
	}

	return hRes;
}

STDMETHODIMP CMyMAPIFormViewer::CopyMessage(LPMAPIFOLDER /*pFolderDestination*/)
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("CopyMessage"),_T("\n"));
	return MAPI_E_NO_SUPPORT;
}

STDMETHODIMP CMyMAPIFormViewer::MoveMessage(LPMAPIFOLDER /*pFolderDestination*/,
											  LPMAPIVIEWCONTEXT /*pViewContext*/,
											  LPCRECT /*prcPosRect*/)
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("MoveMessage"),_T("\n"));
	return MAPI_E_NO_SUPPORT;
}

STDMETHODIMP CMyMAPIFormViewer::DeleteMessage(LPMAPIVIEWCONTEXT /*pViewContext*/,
												LPCRECT /*prcPosRect*/)
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("DeleteMessage"),_T("\n"));
	return MAPI_E_NO_SUPPORT;
}

STDMETHODIMP CMyMAPIFormViewer::SaveMessage()
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("SaveMessage"),_T("\n"));
	HRESULT hRes = S_OK;

	if (!m_lpPersistMessage || !m_lpMessage) return MAPI_E_INVALID_PARAMETER;

	EC_H(m_lpPersistMessage->Save(
		NULL, // m_lpMessage,
		TRUE));
	if (FAILED(hRes))
	{
		LPMAPIERROR lpErr = NULL;
		hRes = m_lpPersistMessage->GetLastError(hRes,fMapiUnicode,&lpErr);
		if (lpErr)
		{
			EC_MAPIERR(fMapiUnicode,lpErr);
			MAPIFreeBuffer(lpErr);
		}
		else CHECKHRES(hRes);
	}
	else
	{
		EC_H(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
		EC_H(m_lpPersistMessage->SaveCompleted(NULL));
	}

	return hRes;
}

STDMETHODIMP CMyMAPIFormViewer::SubmitMessage (ULONG ulFlags)
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("SubmitMessage"),_T("ulFlags = 0x%08X\n"),ulFlags);
	HRESULT hRes = S_OK;
	if (!m_lpPersistMessage || !m_lpMessage) return MAPI_E_INVALID_PARAMETER;

	EC_H(m_lpPersistMessage->Save(
		m_lpMessage,
		TRUE));
	if (FAILED(hRes))
	{
		LPMAPIERROR lpErr = NULL;
		hRes = m_lpPersistMessage->GetLastError(hRes,fMapiUnicode,&lpErr);
		if (lpErr)
		{
			EC_MAPIERR(fMapiUnicode,lpErr);
			MAPIFreeBuffer(lpErr);
		}
		else CHECKHRES(hRes);
	}
	else
	{
		EC_H(m_lpPersistMessage->HandsOffMessage());

		EC_H(m_lpMessage->SubmitMessage(NULL));
	}

	m_lpMessage->Release();
	m_lpMessage = NULL;
	return S_OK;
}

STDMETHODIMP CMyMAPIFormViewer::GetSiteStatus (LPULONG lpulStatus)
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("GetSiteStatus"),_T("\n"));
	*lpulStatus =
		VCSTATUS_NEW_MESSAGE;
	if (m_lpPersistMessage)
	{
		*lpulStatus |=
			VCSTATUS_SAVE |
			VCSTATUS_SUBMIT |
			NULL;
	}
	return S_OK;
}

STDMETHODIMP CMyMAPIFormViewer::GetLastError(HRESULT hResult,
						  ULONG ulFlags,
						  LPMAPIERROR FAR * /*lppMAPIError*/)
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("GetLastError"),_T("hResult = 0x%08X, ulFlags = 0x%08X\n"),hResult,ulFlags);
	return MAPI_E_NO_SUPPORT;
}

///////////////////////////////////////////////////////////////////////////////
// End IMAPIMessageSite implementation
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// IMAPIViewAdviseSink implementation
///////////////////////////////////////////////////////////////////////////////

// Assuming we've advised on this form, we need to Unadvise it now, or it will never unload
STDMETHODIMP CMyMAPIFormViewer::OnShutdown()
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("OnShutdown"),_T("\n"));
	return MAPI_E_NO_SUPPORT;
}

STDMETHODIMP CMyMAPIFormViewer::OnNewMessage()
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("OnNewMessage"),_T("\n"));
	return MAPI_E_NO_SUPPORT;
}

STDMETHODIMP CMyMAPIFormViewer::OnPrint(
					 ULONG dwPageNumber,
					 HRESULT hrStatus)
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("OnPrint"),
		_T("Page Number %u\nStatus: 0x%08X\n"), // STRING_OK
		dwPageNumber,
		hrStatus);
	return MAPI_E_NO_SUPPORT;
}

STDMETHODIMP CMyMAPIFormViewer::OnSubmitted()
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("OnSubmitted"),_T("\n"));
	return MAPI_E_NO_SUPPORT;
}

STDMETHODIMP CMyMAPIFormViewer::OnSaved()
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("OnSaved"),_T("\n"));
	return MAPI_E_NO_SUPPORT;
}

///////////////////////////////////////////////////////////////////////////////
// End IMAPIViewAdviseSink implementation
///////////////////////////////////////////////////////////////////////////////


// Only call this when we really plan on shutting down
void CMyMAPIFormViewer::ShutdownPersist()
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("ShutdownPersist"),_T("\n"));

	if (m_lpPersistMessage)
	{
		m_lpPersistMessage->HandsOffMessage();
		m_lpPersistMessage->Release();
		m_lpPersistMessage = NULL;
	}
}

// Used to set or clear our m_lpPersistMessage
// Can take either a persist or a form interface, but not both
// This should only be called with a persist if the verb is EXCHIVERB_OPEN
HRESULT CMyMAPIFormViewer::SetPersist(LPMAPIFORM lpForm, LPPERSISTMESSAGE lpPersist)
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("SetPersist"),_T("\n"));
	HRESULT hRes = S_OK;
	ShutdownPersist();

	SizedSPropTagArray(1, sptaFlags) = {1,{PR_MESSAGE_FLAGS}};
	ULONG			cValues		= 0L;
	LPSPropValue	lpPropArray	= NULL;

	EC_H(m_lpMessage->GetProps((LPSPropTagArray)&sptaFlags, 0, &cValues, &lpPropArray));
	BOOL bComposing = (lpPropArray && (lpPropArray->Value.l & MSGFLAG_UNSENT));
	MAPIFreeBuffer(lpPropArray);

	if (bComposing && !RegKeys[regkeyALLOW_PERSIST_CACHE].ulCurDWORD)
	{
		//	Message is in compose mode - caching the IPersistMessage might hose Outlook
		//	However, if we don't cache the IPersistMessage, we can't support SaveMessage or SubmitMessage
		DebugPrint(DBGFormViewer,_T("Not caching the persist\n"));
		return S_OK;
	}
	DebugPrint(DBGFormViewer,_T("Caching the persist\n"));

	// Doing this part (saving the persist) will leak Winword in compose mode - which is why we make the above check
	// trouble is - compose mode is when we need the persist, for SaveMessage and SubmitMessage
	if (lpPersist)
	{
		m_lpPersistMessage = lpPersist;
		m_lpPersistMessage->AddRef();
	}
	else if (lpForm)
	{
		EC_H(lpForm->QueryInterface(IID_IPersistMessage ,(LPVOID *)&m_lpPersistMessage));
	}

	return hRes;
}

HRESULT CMyMAPIFormViewer::CallDoVerb(LPMAPIFORM lpMapiForm,
				   LONG lVerb,
				   LPCRECT lpRect)
{
	HRESULT hRes = S_OK;
	if (EXCHIVERB_OPEN == lVerb)
	{
		SetPersist(lpMapiForm,NULL);
	}

	WC_H(lpMapiForm->DoVerb(
		lVerb,
		// (IMAPIViewContext *) this, // view context
		NULL, // view context
		(ULONG_PTR) m_hwndParent, // parent window
		lpRect)); // RECT structure with size
	if (S_OK != hRes)
	{
		hRes = S_OK;
		RECT Rect;

		Rect.left = 0;
		Rect.right = 500;
		Rect.top = 0;
		Rect.bottom = 400;
		EC_H(lpMapiForm->DoVerb(
			lVerb,
			// (IMAPIViewContext *) this, // view context
			NULL, // view context
			(ULONG_PTR) m_hwndParent, // parent window
			&Rect)); // RECT structure with size
	}
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
// IMAPIViewContext implementation
///////////////////////////////////////////////////////////////////////////////

STDMETHODIMP CMyMAPIFormViewer::SetAdviseSink(LPMAPIFORMADVISESINK pmvns)
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("SetAdviseSink"),_T("pmvns = 0x%X\n"),pmvns);
	if (m_lpMapiFormAdviseSink) m_lpMapiFormAdviseSink->Release();
	if (pmvns)
	{
		m_lpMapiFormAdviseSink = pmvns;
		m_lpMapiFormAdviseSink->AddRef();
	}
	else
	{
		m_lpMapiFormAdviseSink = NULL;
	}
	return S_OK;
}

STDMETHODIMP CMyMAPIFormViewer::ActivateNext(ULONG ulDir,
											  LPCRECT lpRect)
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("ActivateNext"),_T("ulDir = 0x%X\n"),ulDir);
	HRESULT	hRes = S_OK;

	enum {ePR_MESSAGE_FLAGS,ePR_MESSAGE_CLASS_A,NUM_COLS};
	SizedSPropTagArray(NUM_COLS,sptaShowForm) = { NUM_COLS, {
		PR_MESSAGE_FLAGS,
			PR_MESSAGE_CLASS_A}
	};

	int			iNewItem = -1;
	LPMESSAGE	lpNewMessage = NULL;
	ULONG		ulMessageStatus = NULL;
	BOOL		bUsedCurrentSite = false;

	WC_H(GetNextMessage(ulDir,&iNewItem,&ulMessageStatus,&lpNewMessage));
	if (lpNewMessage)
	{
		ULONG			cValuesShow = 0;
		LPSPropValue	lpspvaShow = NULL;

		EC_H_GETPROPS(lpNewMessage->GetProps(
			(LPSPropTagArray) &sptaShowForm, // property tag array
			fMapiUnicode, // flags
			&cValuesShow, // Count of values returned
			&lpspvaShow)); // Values returned
		if (lpspvaShow)
		{
			LPPERSISTMESSAGE lpNewPersistMessage = NULL;

			if (m_lpMapiFormAdviseSink)
			{
				DebugPrint(DBGFormViewer,_T("Calling OnActivateNext: szClass = %hs, ulStatus = 0x%X, ulFlags = 0x%X\n"),
					lpspvaShow[ePR_MESSAGE_CLASS_A].Value.lpszA,
					ulMessageStatus,
					lpspvaShow[ePR_MESSAGE_FLAGS].Value.ul);

				WC_H(m_lpMapiFormAdviseSink->OnActivateNext(
					lpspvaShow[ePR_MESSAGE_CLASS_A].Value.lpszA,
					ulMessageStatus, // message status
					lpspvaShow[ePR_MESSAGE_FLAGS].Value.ul, // message flags
					&lpNewPersistMessage));
			}
			else
			{
				hRes = S_FALSE;
			}

			if (S_OK == hRes) // we can handle the message ourselves
			{
				if (lpNewPersistMessage)
				{
					DebugPrintEx(DBGFormViewer,CLASS,_T("ActivateNext"),_T("Got new persist from OnActivateNext\n"));

					EC_H(OpenMessageNonModal(
						m_lpMDB,
						m_lpMAPISession,
						m_lpFolder,
						m_lpContentsTableListCtrl,
						iNewItem,
						lpNewMessage,
						EXCHIVERB_OPEN,
						lpRect));
				}
				else
				{
					ErrDialog(__FILE__,__LINE__,IDS_EDACTIVATENEXT);
				}
			}
			// We have to load the form from scratch
			else if (hRes == S_FALSE)
			{
				DebugPrintEx(DBGFormViewer,CLASS,_T("ActivateNext"),_T("Didn't get new persist from OnActivateNext\n"));
				// we're going to return S_FALSE, which will shut us down, so we can spin a whole new site
				// we don't need to clean up this site since the shutdown will do it for us
				// BTW - it might be more efficient to in-line this code and eliminate a GetProps call
				EC_H(OpenMessageNonModal(
					m_lpMDB,
					m_lpMAPISession,
					m_lpFolder,
					m_lpContentsTableListCtrl,
					iNewItem,
					lpNewMessage,
					EXCHIVERB_OPEN,
					lpRect));
			}
			else CHECKHRESMSG(hRes,IDS_ONACTIVENEXTFAILED);

			if (lpNewPersistMessage) lpNewPersistMessage->Release();
			MAPIFreeBuffer(lpspvaShow);
		}
		lpNewMessage->Release();
	}

	if (!bUsedCurrentSite) return S_FALSE;
	return hRes;
}

STDMETHODIMP CMyMAPIFormViewer::GetPrintSetup(ULONG ulFlags,
											   LPFORMPRINTSETUP FAR * /*lppFormPrintSetup*/)
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("GetPrintSetup"),_T("ulFlags = 0x%08X\n"),ulFlags);
	return MAPI_E_NO_SUPPORT;
}

STDMETHODIMP CMyMAPIFormViewer::GetSaveStream(ULONG FAR * /*pulFlags*/,
											   ULONG FAR * /*pulFormat*/,
											   LPSTREAM FAR * /*ppstm*/)
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("GetSaveStream"),_T("\n"));
	return MAPI_E_NO_SUPPORT;
}

STDMETHODIMP CMyMAPIFormViewer::GetViewStatus(LPULONG lpulStatus)
{
	HRESULT hRes = S_OK;

	DebugPrintEx(DBGFormViewer,CLASS,_T("GetViewStatus"),_T("\n"));
	*lpulStatus = NULL;
	*lpulStatus |= VCSTATUS_INTERACTIVE;

	if (m_lpContentsTableListCtrl)
	{
		if (-1 != m_lpContentsTableListCtrl->GetNextItem(m_iItem,LVNI_BELOW))
		{
			*lpulStatus |= VCSTATUS_NEXT;
		}

		if (-1 != m_lpContentsTableListCtrl->GetNextItem(m_iItem,LVNI_ABOVE))
		{
			*lpulStatus |= VCSTATUS_PREV;
		}
	}
	DebugPrintEx(DBGFormViewer,CLASS,_T("GetViewStatus"),_T("Returning 0x%X\n"),*lpulStatus);
	return hRes;
}

///////////////////////////////////////////////////////////////////////////////
// End IMAPIViewContext implementation
///////////////////////////////////////////////////////////////////////////////

HRESULT CMyMAPIFormViewer::GetNextMessage(
										   ULONG ulDir,
										   int* piNewItem,
										   ULONG* pulStatus,
										   LPMESSAGE* ppMessage)
{
	DebugPrintEx(DBGFormViewer,CLASS,_T("GetNextMessage"),_T("ulDir = 0x%X\n"),ulDir);
	HRESULT hRes = S_OK;

	*piNewItem = NULL;
	*pulStatus = NULL;
	*ppMessage = NULL;

	// Without a view list control, we can't do 'next'
	if (!m_lpContentsTableListCtrl || !m_lpMDB) return MAPI_E_INVALID_PARAMETER;
	if (m_iItem == -1) return S_FALSE;

	if (ulDir & VCDIR_NEXT)
	{
		*piNewItem = m_lpContentsTableListCtrl->GetNextItem(m_iItem,LVNI_BELOW);
	}
	else if (ulDir & VCDIR_PREV)
	{
		*piNewItem = m_lpContentsTableListCtrl->GetNextItem(m_iItem,LVNI_ABOVE);
	}
	if (*piNewItem == -1)
	{
		hRes = S_FALSE;
	}
	else
	{
		SortListData* lpData = NULL; // do not free
		lpData = (SortListData*) m_lpContentsTableListCtrl->GetItemData(*piNewItem);

		if (lpData)
		{
			LPSBinary lpEID = NULL;
			lpEID = lpData->data.Contents.lpEntryID;
			if (lpEID)
			{
				EC_H(CallOpenEntry(
					m_lpMDB,
					NULL,
					NULL,
					NULL,
					lpEID->cb,
					(LPENTRYID) lpEID->lpb,
					NULL,
					MAPI_BEST_ACCESS,
					NULL,
					(LPUNKNOWN*)ppMessage));

				EC_H(m_lpFolder->GetMessageStatus(
					lpEID->cb,
					(LPENTRYID)lpEID->lpb,
					0,
					pulStatus));
			}
		}
	}

	if (FAILED(hRes))
	{
		*piNewItem = -1;
		*pulStatus = NULL;
		if (*ppMessage)
		{
			(*ppMessage)->Release();
			*ppMessage = NULL;
		}
	}

	return hRes;
}
