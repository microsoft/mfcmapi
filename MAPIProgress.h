#pragma once
// MAPIProgress.h: interface for the CMAPIProgress

class CMAPIProgress : public IMAPIProgress
{
public:
	CMAPIProgress(wstring lpszContext, _In_ HWND hWnd);
	virtual ~CMAPIProgress();

private:
	// IUnknown
	STDMETHODIMP QueryInterface(REFIID riid, LPVOID * ppvObj);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	// IMAPIProgress
	_Check_return_ STDMETHODIMP Progress(ULONG ulValue, ULONG ulCount, ULONG ulTotal);
	STDMETHODIMP GetFlags(ULONG* lpulFlags);
	STDMETHODIMP GetMax(ULONG* lpulMax);
	STDMETHODIMP GetMin(ULONG* lpulMin);
	STDMETHODIMP SetLimits(ULONG* lpulMin, ULONG* lpulMax, ULONG* lpulFlags);

	void OutputState(wstring lpszFunction);

	LONG		m_cRef;
	ULONG		m_ulMin;
	ULONG		m_ulMax;
	ULONG		m_ulFlags;
	wstring		m_szContext;
	HWND		m_hWnd;
};

_Check_return_ CMAPIProgress* GetMAPIProgress(_In_z_ LPCTSTR lpszContext, _In_ HWND hWnd);