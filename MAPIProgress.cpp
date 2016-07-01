#include "stdafx.h"
#include "MAPIProgress.h"
#include "MAPIFunctions.h"
#include "enums.h"
#include "BaseDialog.h"

static wstring CLASS = L"CMAPIProgress";

#ifdef MRMAPI
_Check_return_ CMAPIProgress* GetMAPIProgress(wstring /*lpszContext*/, _In_ HWND /*hWnd*/)
{
	return NULL;
}
#endif

#ifndef MRMAPI
_Check_return_ CMAPIProgress* GetMAPIProgress(wstring lpszContext, _In_ HWND hWnd)
{
	if (RegKeys[regkeyUSE_IMAPIPROGRESS].ulCurDWORD)
	{
		CMAPIProgress * pProgress = new CMAPIProgress(lpszContext, hWnd);

		return pProgress;
	}

	return NULL;
}

CMAPIProgress::CMAPIProgress(wstring lpszContext, _In_ HWND hWnd)
{
	TRACE_CONSTRUCTOR(CLASS);

	m_cRef = 1;
	m_ulMin = 1;
	m_ulMax = 1000;
	m_ulFlags = MAPI_TOP_LEVEL;
	m_hWnd = hWnd;

	if (!lpszContext.empty())
	{
		m_szContext = lpszContext;
	}
	else
	{
		m_szContext = loadstring(IDS_NOCONTEXT);
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
	TRACE_ADDREF(CLASS, lCount);
	return lCount;
}

STDMETHODIMP_(ULONG) CMAPIProgress::Release()
{
	LONG lCount = InterlockedDecrement(&m_cRef);
	TRACE_RELEASE(CLASS, lCount);
	if (!lCount) delete this;
	return lCount;
}

_Check_return_ STDMETHODIMP CMAPIProgress::Progress(ULONG ulValue, ULONG ulCount, ULONG ulTotal)
{
	DebugPrintEx(DBGGeneric, CLASS, L"Progress", L"(%ws) - ulValue = %u, ulCount = %u, ulTotal = %u\n",
		m_szContext.c_str(), ulValue, ulCount, ulTotal);

	OutputState(L"Progress");

	if (m_hWnd)
	{
		int iPercent = ::MulDiv(ulValue - m_ulMin, 100, m_ulMax - m_ulMin);
		CBaseDialog::UpdateStatus(m_hWnd, STATUSINFOTEXT, formatmessage(IDS_PERCENTLOADED, m_szContext.c_str(), iPercent));
	}

	return S_OK;
}

STDMETHODIMP CMAPIProgress::GetFlags(ULONG* lpulFlags)
{
	if (!lpulFlags)
	{
		return MAPI_E_INVALID_PARAMETER;
	}

	OutputState(L"GetFlags");

	*lpulFlags = m_ulFlags;
	return S_OK;
}

STDMETHODIMP CMAPIProgress::GetMax(ULONG* lpulMax)
{
	if (!lpulMax)
		return MAPI_E_INVALID_PARAMETER;

	OutputState(L"GetMax");

	*lpulMax = m_ulMax;
	return S_OK;
}

STDMETHODIMP CMAPIProgress::GetMin(ULONG* lpulMin)
{
	if (!lpulMin)
		return MAPI_E_INVALID_PARAMETER;

	OutputState(L"GetMin");

	*lpulMin = m_ulMin;
	return S_OK;
}

STDMETHODIMP CMAPIProgress::SetLimits(ULONG* lpulMin, ULONG* lpulMax, ULONG* lpulFlags)
{
	OutputState(L"SetLimits");

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

	DebugPrintEx(DBGGeneric, CLASS, L"SetLimits", L"(%ws) - Passed Values: lpulMin = %ws, lpulMax = %ws, lpulFlags = %ws\n",
		m_szContext.c_str(), LPCTSTRToWstring(szMin).c_str(), LPCTSTRToWstring(szMax).c_str(), LPCTSTRToWstring(szFlags).c_str());

	if (lpulMin)
		m_ulMin = *lpulMin;

	if (lpulMax)
		m_ulMax = *lpulMax;

	if (lpulFlags)
		m_ulFlags = *lpulFlags;

	OutputState(L"SetLimits");

	return S_OK;
}

void CMAPIProgress::OutputState(wstring lpszFunction)
{
	DebugPrint(DBGGeneric, L"%ws::%ws(%ws) - Current Values: Min = %u, Max = %u, Flags = %u\n",
		CLASS, lpszFunction.c_str(), m_szContext.c_str(), m_ulMin, m_ulMax, m_ulFlags);
}
#endif