#pragma once
// Interface for the CMapiObjects class.

namespace cache
{
	class CMapiObjects
	{
	public:
		CMapiObjects(_In_opt_ CMapiObjects* OldMapiObjects);
		virtual ~CMapiObjects();

		STDMETHODIMP_(ULONG) AddRef();
		STDMETHODIMP_(ULONG) Release();

		_Check_return_ LPADRBOOK GetAddrBook(bool bForceOpen);
		_Check_return_ LPMDB GetMDB() const;
		_Check_return_ LPMAPISESSION GetSession() const;
		_Check_return_ LPMAPISESSION LogonGetSession(_In_ HWND hWnd);

		void SetAddrBook(_In_opt_ LPADRBOOK lpAddrBook);
		void SetMDB(_In_opt_ LPMDB lppMDB);
		void MAPILogonEx(_In_ HWND hwnd, _In_opt_z_ LPTSTR szProfileName, ULONG ulFlags);
		void Logoff(_In_ HWND hwnd, ULONG ulFlags);

	private:
		LONG m_cRef;
		LPMDB m_lpMDB;
		LPMAPISESSION m_lpMAPISession;
		LPADRBOOK m_lpAddrBook;
	};
} // namespace cache