#pragma once
// Interface for the CMapiObjects class.

namespace cache
{
	class CMapiObjects
	{
	public:
		CMapiObjects(_In_opt_ std::shared_ptr<CMapiObjects> OldMapiObjects);
		~CMapiObjects();
		CMapiObjects(CMapiObjects const&) = delete;
		void operator=(CMapiObjects const&) = delete;
		CMapiObjects(CMapiObjects&&) = delete;
		CMapiObjects& operator=(CMapiObjects&&) = delete;

		_Check_return_ LPADRBOOK GetAddrBook(bool bForceOpen);
		_Check_return_ LPMDB GetMDB() const noexcept;
		_Check_return_ LPMAPISESSION GetSession() const noexcept;
		_Check_return_ LPMAPISESSION LogonGetSession(_In_ HWND hWnd);

		void SetAddrBook(_In_opt_ LPADRBOOK lpAddrBook);
		void SetMDB(_In_opt_ LPMDB lppMDB);
		void MAPILogonEx(_In_ HWND hwnd, _In_opt_z_ LPTSTR szProfileName, ULONG ulFlags);
		void Logoff(_In_ HWND hwnd, ULONG ulFlags);

	private:
		LPMDB m_lpMDB{};
		LPMAPISESSION m_lpMAPISession{};
		LPADRBOOK m_lpAddrBook{};
	};
} // namespace cache