#pragma once

class CAccountHelper : public IOlkAccountHelper
{
public:
	CAccountHelper(LPCWSTR lpwszProfName, LPMAPISESSION lpSession);
	~CAccountHelper();

	// IUnknown
	STDMETHODIMP QueryInterface(REFIID riid, LPVOID* ppvObj) override;
	STDMETHODIMP_(ULONG) AddRef() noexcept override;
	STDMETHODIMP_(ULONG) Release() noexcept override;

	// IOlkAccountHelper
	STDMETHODIMP PlaceHolder1(LPVOID) noexcept override { return E_NOTIMPL; }

	STDMETHODIMP GetIdentity(LPWSTR pwszIdentity, DWORD* pcch) noexcept override;
	STDMETHODIMP GetMapiSession(LPUNKNOWN* ppmsess) override;
	STDMETHODIMP HandsOffSession() noexcept override { return E_NOTIMPL; }

private:
	LONG m_cRef{};
	LPMAPISESSION m_lpSession{};
	std::wstring m_lpwszProfile{};
};