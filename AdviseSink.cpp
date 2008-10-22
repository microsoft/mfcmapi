// AdviseSink.cpp: implementation of the CAdviseSink class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "AdviseSink.h"

static TCHAR* CLASS = _T("CAdviseSink");

CAdviseSink::CAdviseSink(HWND hWndParent, HTREEITEM hTreeParent)
{
	TRACE_CONSTRUCTOR(CLASS);
	m_cRef = 1;
	m_hWndParent = hWndParent;
	m_hTreeParent = hTreeParent;
};

CAdviseSink::~CAdviseSink()
{
	TRACE_DESTRUCTOR(CLASS);
};


STDMETHODIMP CAdviseSink::QueryInterface(REFIID riid,
											  LPVOID * ppvObj)
{
	*ppvObj = 0;
	if (riid == IID_IMAPIAdviseSink ||
		riid == IID_IUnknown)
	{
		*ppvObj = (LPVOID)this;
		AddRef();
		return S_OK;
	}
	return E_NOINTERFACE;
};

STDMETHODIMP_(ULONG) CAdviseSink::AddRef()
{
	LONG lCount = InterlockedIncrement(&m_cRef);
	TRACE_ADDREF(CLASS,lCount);
	return lCount;
};

STDMETHODIMP_(ULONG) CAdviseSink::Release()
{
	LONG lCount = InterlockedDecrement(&m_cRef);
	TRACE_RELEASE(CLASS,lCount);
	if (!lCount)  delete this;
	return lCount;
};

STDMETHODIMP_(ULONG) CAdviseSink::OnNotify (ULONG cNotify,
												 LPNOTIFICATION lpNotifications)
{
	HRESULT			hRes = S_OK;

	DebugPrintNotifications(DBGNotify, cNotify, lpNotifications);

	if (!m_hWndParent) return 0;

	for (ULONG i=0 ; i<cNotify ; i++)
	{
		hRes = S_OK;
		if (fnevNewMail == lpNotifications[i].ulEventType)
		{
			MessageBeep(MB_OK);
		}
		else if (fnevTableModified == lpNotifications[i].ulEventType)
		{
			switch(lpNotifications[i].info.tab.ulTableEvent)
			{
			case(TABLE_ERROR):
			case(TABLE_CHANGED):
			case(TABLE_RELOAD):
				EC_H((HRESULT)::SendMessage(m_hWndParent,WM_MFCMAPI_REFRESHTABLE,(WPARAM) m_hTreeParent,0));
				break;
			case(TABLE_ROW_ADDED):
				EC_H((HRESULT)::SendMessage(
					m_hWndParent,
					WM_MFCMAPI_ADDITEM,
					(WPARAM) &lpNotifications[i].info.tab,
					(LPARAM) m_hTreeParent));
				break;
			case(TABLE_ROW_DELETED):
				EC_H((HRESULT)::SendMessage(
					m_hWndParent,
					WM_MFCMAPI_DELETEITEM,
					(WPARAM) &lpNotifications[i].info.tab,
					(LPARAM) m_hTreeParent));
				break;
			case(TABLE_ROW_MODIFIED):
				EC_H((HRESULT)::SendMessage(
					m_hWndParent,
					WM_MFCMAPI_MODIFYITEM,
					(WPARAM) &lpNotifications[i].info.tab,
					(LPARAM) m_hTreeParent));
				break;
			}
		}
	}
	return 0;
}
