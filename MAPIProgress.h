// MAPIProgress.h: interface for the CMAPIProgress

#pragma once

class CMAPIProgress : public IMAPIProgress
{
	// Constructors and destructors
public :
	CMAPIProgress(LPCTSTR lpszContext, HWND hWnd);
	virtual ~CMAPIProgress();

	STDMETHODIMP			QueryInterface(REFIID riid, LPVOID * ppvObj);
	STDMETHODIMP_(ULONG)	AddRef();
	STDMETHODIMP_(ULONG)	Release();

	// IMAPIProgress

	STDMETHODIMP Progress(ULONG	ulValue, ULONG ulCount, ULONG ulTotal);
	STDMETHODIMP GetFlags(ULONG FAR* lpulFlags);
	STDMETHODIMP GetMax(ULONG FAR* lpulMax);
	STDMETHODIMP GetMin(ULONG FAR* lpulMin);
	STDMETHODIMP SetLimits(ULONG FAR* lpulMin, ULONG FAR* lpulMax, ULONG FAR* lpulFlags);

private :
	void OutputState(LPTSTR lpszFunction);

	LONG		m_cRef;
	ULONG		m_ulMin;
	ULONG		m_ulMax;
	ULONG		m_ulFlags;
	CString		m_szContext;
	HWND		m_hWnd;
};
//
//////////////////////////////////////////////////////////////////////

CMAPIProgress * GetMAPIProgress(LPTSTR lpszContext, HWND hWnd);