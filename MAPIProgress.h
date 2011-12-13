#pragma once
// MAPIProgress.h: interface for the CMAPIProgress

class CMAPIProgress : public IMAPIProgress
{
public:
	CMAPIProgress(_In_z_ LPCTSTR lpszContext, _In_ HWND hWnd);
	virtual ~CMAPIProgress();

private:
	// IUnknown
	_Check_return_ STDMETHODIMP QueryInterface(REFIID riid, _Deref_out_opt_ LPVOID * ppvObj);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	// IMAPIProgress
	_Check_return_ STDMETHODIMP Progress(ULONG ulValue, ULONG ulCount, ULONG ulTotal);
	_Check_return_ STDMETHODIMP GetFlags(_Inout_ ULONG* lpulFlags);
	_Check_return_ STDMETHODIMP GetMax(_Inout_ ULONG* lpulMax);
	_Check_return_ STDMETHODIMP GetMin(_Inout_ ULONG* lpulMin);
	_Check_return_ STDMETHODIMP SetLimits(_Inout_ ULONG* lpulMin, _Inout_ ULONG* lpulMax, _Inout_ ULONG* lpulFlags);

	void OutputState(_In_z_ LPCTSTR lpszFunction);

	LONG		m_cRef;
	ULONG		m_ulMin;
	ULONG		m_ulMax;
	ULONG		m_ulFlags;
	CString		m_szContext;
	HWND		m_hWnd;
};

_Check_return_ CMAPIProgress* GetMAPIProgress(_In_z_ LPCTSTR lpszContext, _In_ HWND hWnd);