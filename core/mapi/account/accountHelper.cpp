#include <core/stdafx.h>
#include <core/mapi/account/accountHelper.h>
#include <core/mapi/mapiFunctions.h>

CAccountHelper::CAccountHelper(LPCWSTR lpwszProfName, LPMAPISESSION lpSession)
{
	m_cRef = 1;
	m_lpSession = mapi::safe_cast<LPMAPISESSION>(lpSession);
	m_lpwszProfile = lpwszProfName;
}

CAccountHelper::~CAccountHelper()
{
	if (m_lpSession) m_lpSession->Release();
}

STDMETHODIMP CAccountHelper::QueryInterface(REFIID riid, LPVOID* ppvObj)
{
	if (!ppvObj) return MAPI_E_INVALID_PARAMETER;
	*ppvObj = nullptr;
	if (riid == IID_IOlkAccountHelper || riid == IID_IUnknown)
	{
		*ppvObj = static_cast<LPVOID>(this);
		AddRef();
		return S_OK;
	}

	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) CAccountHelper::AddRef() noexcept
{
	const auto lCount = InterlockedIncrement(&m_cRef);
	return lCount;
}

STDMETHODIMP_(ULONG) CAccountHelper::Release() noexcept
{
	const auto lCount = InterlockedDecrement(&m_cRef);
	if (!lCount) delete this;
	return lCount;
}

STDMETHODIMP CAccountHelper::GetIdentity(LPWSTR pwszIdentity, DWORD* pcch) noexcept
{
	if (!pcch || m_lpwszProfile.empty()) return E_INVALIDARG;

	if (m_lpwszProfile.length() > *pcch)
	{
		*pcch = DWORD{m_lpwszProfile.length()};
		return E_OUTOFMEMORY;
	}

	wcscpy_s(pwszIdentity, *pcch, m_lpwszProfile.c_str());
	*pcch = DWORD{m_lpwszProfile.length()};

	return S_OK;
}

STDMETHODIMP CAccountHelper::GetMapiSession(LPUNKNOWN* ppmsess)
{
	if (!ppmsess) return E_INVALIDARG;

	*ppmsess = mapi::safe_cast<LPUNKNOWN>(m_lpSession);
	return S_OK;
}