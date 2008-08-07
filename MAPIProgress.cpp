#include "stdafx.h"
#include "Error.h"

#include "MAPIProgress.h"
#include "MAPIFunctions.h"
#include "enums.h"

#ifdef _DEBUG
#undef THIS_FILE
static char THIS_FILE[]=__FILE__;
#define new DEBUG_NEW
#endif

static TCHAR* CLASS = _T("CMAPIProgress");

CMAPIProgress * GetMAPIProgress(LPTSTR lpszContext, HWND hWnd)
{
	if(RegKeys[regkeyUSE_IMAPIPROGRESS].ulCurDWORD)
	{
		CMAPIProgress * pProgress = new CMAPIProgress(lpszContext, hWnd);

		return pProgress;
	}

	return NULL;
}

CMAPIProgress::CMAPIProgress(LPCTSTR lpszContext, HWND hWnd)
{
	TRACE_CONSTRUCTOR(CLASS);

	m_cRef = 1;
	m_ulMin = 1;
	m_ulMax = 1000;
	m_ulFlags = MAPI_TOP_LEVEL;
	m_hWnd = hWnd;

	if(lpszContext)
	{
		m_szContext = lpszContext;
	}
	else
	{
		m_szContext.LoadString(IDS_NOCONTEXT);
	}

}

CMAPIProgress::~CMAPIProgress()
{
	TRACE_DESTRUCTOR(CLASS);
}


STDMETHODIMP CMAPIProgress::QueryInterface(REFIID riid,
											  LPVOID * ppvObj)
{
	*ppvObj = 0;
	if (riid == IID_IMAPIProgress ||
		riid == IID_IUnknown)
	{
		*ppvObj = (LPVOID)this;
		AddRef();
		return S_OK;
	}
	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) CMAPIProgress::AddRef()
{
	LONG lCount = InterlockedIncrement(&m_cRef);
	TRACE_ADDREF(CLASS,lCount);
	return lCount;
}

STDMETHODIMP_(ULONG) CMAPIProgress::Release()
{
	LONG lCount = InterlockedDecrement(&m_cRef);
	TRACE_RELEASE(CLASS,lCount);
	if (!lCount)  delete this;
	return lCount;
};

STDMETHODIMP CMAPIProgress::Progress(ULONG	ulValue, ULONG ulCount, ULONG ulTotal)
{
	DebugPrintEx(DBGGeneric, CLASS, _T("Progress"), _T("(%s) - ulValue = %d, ulCount = %d, ulTotal = %d\n"),
		m_szContext, ulValue, ulCount, ulTotal);

	OutputState(_T("Progress"));

	if (m_hWnd)
	{
		int iPercent = ::MulDiv(ulValue - m_ulMin, 100, m_ulMax - m_ulMin);
		CString szStatusText;

		szStatusText.FormatMessage(IDS_PERCENTLOADED, m_szContext, iPercent);
		(void)::SendMessage(m_hWnd, WM_MFCMAPI_UPDATESTATUSBAR, STATUSRIGHTPANE, (LPARAM)(LPCTSTR) szStatusText);
	}

	return S_OK;
}

STDMETHODIMP CMAPIProgress::GetFlags(ULONG FAR* lpulFlags)
{
	if(!lpulFlags)
	{
		return MAPI_E_INVALID_PARAMETER;
	}

	OutputState(_T("GetFlags"));

	*lpulFlags = m_ulFlags;
	return S_OK;
}

STDMETHODIMP CMAPIProgress::GetMax(ULONG FAR* lpulMax)
{
	if(!lpulMax)
		return MAPI_E_INVALID_PARAMETER;

	OutputState(_T("GetMax"));

	*lpulMax = m_ulMax;
	return S_OK;
}

STDMETHODIMP CMAPIProgress::GetMin(ULONG FAR* lpulMin)
{
	if(!lpulMin)
		return MAPI_E_INVALID_PARAMETER;

	OutputState(_T("GetMin"));

	*lpulMin = m_ulMin;
	return S_OK;
}

STDMETHODIMP CMAPIProgress::SetLimits(ULONG FAR* lpulMin, ULONG FAR* lpulMax, ULONG FAR* lpulFlags)
{
	OutputState(_T("SetLimits"));

	HRESULT hRes = S_OK;
	TCHAR szMin[16];
	TCHAR szMax[16];
	TCHAR szFlags[16];

	if(lpulMin)
	{
		EC_H(StringCchPrintf(szMin, CCH(szMin), _T("%d"), *lpulMin));// STRING_OK
	}
	else
	{
		EC_H(StringCchPrintf(szMin, CCH(szMin), _T("NULL")));// STRING_OK
	}

	if(lpulMax)
	{
		EC_H(StringCchPrintf(szMax, CCH(szMax), _T("%d"), *lpulMax));// STRING_OK
	}
	else
	{
		EC_H(StringCchPrintf(szMax, CCH(szMax), _T("NULL")));// STRING_OK
	}

	if(lpulFlags)
	{
		EC_H(StringCchPrintf(szFlags, CCH(szFlags), _T("0x%08X"), *lpulFlags));// STRING_OK
	}
	else
	{
		EC_H(StringCchPrintf(szFlags, CCH(szFlags), _T("NULL")));// STRING_OK
	}

	DebugPrintEx(DBGGeneric, CLASS, _T("SetLimits"), _T("(%s) - Passed Values: lpulMin = %s, lpulMax = %s, lpulFlags = %s\n"),
		m_szContext, szMin, szMax, szFlags);

	if(lpulMin)
		m_ulMin = *lpulMin;

	if(lpulMax)
		m_ulMax = *lpulMax;

	if(lpulFlags)
		m_ulFlags = *lpulFlags;

	OutputState(_T("SetLimits"));

	return S_OK;
}

void CMAPIProgress::OutputState(LPTSTR lpszFunction)
{
	DebugPrint(DBGGeneric,_T("%s::%s(%s) - Current Values: Min = %d, Max = %d, Flags = %d\n"),
			CLASS, lpszFunction, m_szContext, m_ulMin, m_ulMax, m_ulFlags);
}