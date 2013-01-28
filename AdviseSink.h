#pragma once
// AdviseSink.h: interface for the CAdviseSink class.

class CAdviseSink : public IMAPIAdviseSink
{
public:
	CAdviseSink  (_In_ HWND hWndParent, _In_opt_ HTREEITEM hTreeParent);
	virtual ~CAdviseSink();

	STDMETHODIMP QueryInterface(REFIID riid, LPVOID* ppvObj);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	STDMETHODIMP_(ULONG) OnNotify (ULONG cNotify, LPNOTIFICATION lpNotifications);

private:
	LONG		m_cRef;
	HWND		m_hWndParent;
	HTREEITEM	m_hTreeParent;
};