#pragma once
// AdviseSink.h: interface for the CAdviseSink class.

class CAdviseSink : public IMAPIAdviseSink
{
public:
	CAdviseSink  (_In_ HWND hWndParent, _In_opt_ HTREEITEM hTreeParent);
	virtual ~CAdviseSink();

	_Check_return_ STDMETHODIMP         QueryInterface(_In_ REFIID riid, _Deref_out_opt_ LPVOID* ppvObj);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	_Check_return_ STDMETHODIMP_(ULONG) OnNotify (ULONG cNotify, _In_ LPNOTIFICATION lpNotifications);

private:
	LONG		m_cRef;
	HWND		m_hWndParent;
	HTREEITEM	m_hTreeParent;
};