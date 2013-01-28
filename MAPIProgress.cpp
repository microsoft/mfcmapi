#include "stdafx.h"
#include "MAPIProgress.h"
#include "MAPIFunctions.h"
#include "enums.h"

static TCHAR* CLASS = _T("CMAPIProgress");

#ifdef MRMAPI
_Check_return_ CMAPIProgress* GetMAPIProgress(_In_z_ LPCTSTR /*lpszContext*/, _In_ HWND /*hWnd*/)
{
	return NULL;
} // GetMAPIProgress
#endif

#ifndef MRMAPI
_Check_return_ CMAPIProgress* GetMAPIProgress(_In_z_ LPCTSTR lpszContext, _In_ HWND hWnd)
{
	if (RegKeys[regkeyUSE_IMAPIPROGRESS].ulCurDWORD)
	{
		CMAPIProgress * pProgress = new CMAPIProgress(lpszContext, hWnd);

		return pProgress;
	}

	return NULL;
} // GetMAPIProgress

CMAPIProgress::CMAPIProgress(_In_z_ LPCTSTR lpszContext, _In_ HWND hWnd)
{
	TRACE_CONSTRUCTOR(CLASS);

	m_cRef = 1;
	m_ulMin = 1;
	m_ulMax = 1000;
	m_ulFlags = MAPI_TOP_LEVEL;
	m_hWnd = hWnd;

	if (lpszContext)
	{
		m_szContext = lpszContext;
	}
	else
	{
		HRESULT hRes = S_OK;
		EC_B(m_szContext.LoadString(IDS_NOCONTEXT));
	}
} // CMAPIProgress::CMAPIProgress

CMAPIProgress::~CMAPIProgress()
{
	TRACE_DESTRUCTOR(CLASS);
} // CMAPIProgress::~CMAPIProgress

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
} // CMAPIProgress::QueryInterface

STDMETHODIMP_(ULONG) CMAPIProgress::AddRef()
{
	LONG lCount = InterlockedIncrement(&m_cRef);
	TRACE_ADDREF(CLASS,lCount);
	return lCount;
} // CMAPIProgress::AddRef

STDMETHODIMP_(ULONG) CMAPIProgress::Release()
{
	LONG lCount = InterlockedDecrement(&m_cRef);
	TRACE_RELEASE(CLASS,lCount);
	if (!lCount)  delete this;
	return lCount;
} // CMAPIProgress::Release

_Check_return_ STDMETHODIMP CMAPIProgress::Progress(ULONG ulValue, ULONG ulCount, ULONG ulTotal)
{
	DebugPrintEx(DBGGeneric, CLASS, _T("Progress"), _T("(%s) - ulValue = %u, ulCount = %u, ulTotal = %u\n"),
		(LPCTSTR) m_szContext, ulValue, ulCount, ulTotal);

	OutputState(_T("Progress"));

	if (m_hWnd)
	{
		int iPercent = ::MulDiv(ulValue - m_ulMin, 100, m_ulMax - m_ulMin);
		CString szStatusText;

		szStatusText.FormatMessage(IDS_PERCENTLOADED, m_szContext, iPercent);
		(void)::SendMessage(m_hWnd, WM_MFCMAPI_UPDATESTATUSBAR, STATUSINFOTEXT, (LPARAM)(LPCTSTR) szStatusText);
	}

	return S_OK;
} // CMAPIProgress::Progress

STDMETHODIMP CMAPIProgress::GetFlags(ULONG* lpulFlags)
{
	if (!lpulFlags)
	{
		return MAPI_E_INVALID_PARAMETER;
	}

	OutputState(_T("GetFlags"));

	*lpulFlags = m_ulFlags;
	return S_OK;
} // CMAPIProgress::GetFlags

STDMETHODIMP CMAPIProgress::GetMax(ULONG* lpulMax)
{
	if (!lpulMax)
		return MAPI_E_INVALID_PARAMETER;

	OutputState(_T("GetMax"));

	*lpulMax = m_ulMax;
	return S_OK;
} // CMAPIProgress::GetMax

STDMETHODIMP CMAPIProgress::GetMin(ULONG* lpulMin)
{
	if (!lpulMin)
		return MAPI_E_INVALID_PARAMETER;

	OutputState(_T("GetMin"));

	*lpulMin = m_ulMin;
	return S_OK;
} // CMAPIProgress::GetMin

STDMETHODIMP CMAPIProgress::SetLimits(ULONG* lpulMin, ULONG* lpulMax, ULONG* lpulFlags)
{
	OutputState(_T("SetLimits"));

	HRESULT hRes = S_OK;
	TCHAR szMin[16];
	TCHAR szMax[16];
	TCHAR szFlags[16];

	if (lpulMin)
	{
		EC_H(StringCchPrintf(szMin, _countof(szMin), _T("%u"), *lpulMin)); // STRING_OK
	}
	else
	{
		EC_H(StringCchPrintf(szMin, _countof(szMin), _T("NULL"))); // STRING_OK
	}

	if (lpulMax)
	{
		EC_H(StringCchPrintf(szMax, _countof(szMax), _T("%u"), *lpulMax)); // STRING_OK
	}
	else
	{
		EC_H(StringCchPrintf(szMax, _countof(szMax), _T("NULL"))); // STRING_OK
	}

	if (lpulFlags)
	{
		EC_H(StringCchPrintf(szFlags, _countof(szFlags), _T("0x%08X"), *lpulFlags)); // STRING_OK
	}
	else
	{
		EC_H(StringCchPrintf(szFlags, _countof(szFlags), _T("NULL"))); // STRING_OK
	}

	DebugPrintEx(DBGGeneric, CLASS, _T("SetLimits"), _T("(%s) - Passed Values: lpulMin = %s, lpulMax = %s, lpulFlags = %s\n"),
		(LPCTSTR) m_szContext, szMin, szMax, szFlags);

	if (lpulMin)
		m_ulMin = *lpulMin;

	if (lpulMax)
		m_ulMax = *lpulMax;

	if (lpulFlags)
		m_ulFlags = *lpulFlags;

	OutputState(_T("SetLimits"));

	return S_OK;
} // CMAPIProgress::SetLimits

void CMAPIProgress::OutputState(_In_z_ LPCTSTR lpszFunction)
{
	DebugPrint(DBGGeneric,_T("%s::%s(%s) - Current Values: Min = %u, Max = %u, Flags = %u\n"),
		CLASS, lpszFunction, (LPCTSTR) m_szContext, m_ulMin, m_ulMax, m_ulFlags);
} // CMAPIProgress::OutputState
#endif