#pragma once
// MAPIProgress.h: interface for the CMAPIProgress

class CMAPIProgress : public IMAPIProgress
{
public:
	CMAPIProgress(LPCTSTR lpszContext, HWND hWnd);
	virtual ~CMAPIProgress();

private:
	// IUnknown
	STDMETHODIMP         QueryInterface(REFIID riid, LPVOID * ppvObj);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();

	// IMAPIProgress
	STDMETHODIMP Progress(ULONG ulValue, ULONG ulCount, ULONG ulTotal);
	STDMETHODIMP GetFlags(ULONG FAR* lpulFlags);
	STDMETHODIMP GetMax(ULONG FAR* lpulMax);
	STDMETHODIMP GetMin(ULONG FAR* lpulMin);
	STDMETHODIMP SetLimits(ULONG FAR* lpulMin, ULONG FAR* lpulMax, ULONG FAR* lpulFlags);

	void OutputState(LPTSTR lpszFunction);

	LONG		m_cRef;
	ULONG		m_ulMin;
	ULONG		m_ulMax;
	ULONG		m_ulFlags;
	CString		m_szContext;
	HWND		m_hWnd;
};

CMAPIProgress * GetMAPIProgress(LPTSTR lpszContext, HWND hWnd);