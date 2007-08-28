// AdviseSink.h: interface for the CAdviseSink class.
//
//////////////////////////////////////////////////////////////////////

#pragma once

class CAdviseSink : public IMAPIAdviseSink
{
	// Constructors and destructors
public :
	CAdviseSink  (HWND hWndParent, HTREEITEM hTreeParent);
	virtual ~CAdviseSink();

	STDMETHODIMP			QueryInterface(REFIID riid, LPVOID * ppvObj);
	STDMETHODIMP_(ULONG)	AddRef();
	STDMETHODIMP_(ULONG)	Release();

	STDMETHODIMP_(ULONG)	OnNotify (ULONG cNotify, LPNOTIFICATION lpNotifications);

private :
	LONG		m_cRef;
	HWND		m_hWndParent;
	HTREEITEM	m_hTreeParent;
};