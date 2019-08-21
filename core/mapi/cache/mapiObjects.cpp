// Implementation of the CMapiObjects class.

#include <core/stdafx.h>
#include <core/mapi/cache/mapiObjects.h>
#include <core/mapi/cache/globalCache.h>
#include <core/utility/output.h>
#include <core/utility/error.h>

namespace cache
{
	static std::wstring CLASS = L"CMapiObjects";
	// Pass an existing CMapiObjects to make a copy, pass NULL to create a new one from scratch
	CMapiObjects::CMapiObjects(_In_opt_ std::shared_ptr<CMapiObjects> oldMapiObjects)
	{
		TRACE_CONSTRUCTOR(CLASS);
		output::DebugPrintEx(output::DBGConDes, CLASS, CLASS, L"OldMapiObjects = %p\n", &oldMapiObjects);

		// If we were passed a valid object, make copies of its interfaces.
		if (oldMapiObjects)
		{
			m_lpMAPISession = oldMapiObjects->m_lpMAPISession;
			if (m_lpMAPISession) m_lpMAPISession->AddRef();

			m_lpMDB = oldMapiObjects->m_lpMDB;
			if (m_lpMDB) m_lpMDB->AddRef();

			m_lpAddrBook = oldMapiObjects->m_lpAddrBook;
			if (m_lpAddrBook) m_lpAddrBook->AddRef();
		}
	}

	CMapiObjects::~CMapiObjects()
	{
		TRACE_DESTRUCTOR(CLASS);
		if (m_lpAddrBook) m_lpAddrBook->Release();
		if (m_lpMDB) m_lpMDB->Release();
		if (m_lpMAPISession) m_lpMAPISession->Release();
	}

	void CMapiObjects::MAPILogonEx(_In_ HWND hwnd, _In_opt_z_ LPTSTR szProfileName, ULONG ulFlags)
	{
		output::DebugPrint(output::DBGGeneric, L"Logging on with MAPILogonEx, ulFlags = 0x%X\n", ulFlags);

		CGlobalCache::getInstance().MAPIInitialize(NULL);
		if (!CGlobalCache::getInstance().bMAPIInitialized()) return;

		if (m_lpMAPISession) m_lpMAPISession->Release();
		m_lpMAPISession = nullptr;

		EC_H_CANCEL_S(
			::MAPILogonEx(reinterpret_cast<ULONG_PTR>(hwnd), szProfileName, nullptr, ulFlags, &m_lpMAPISession));

		output::DebugPrint(output::DBGGeneric, L"\tm_lpMAPISession set to %p\n", m_lpMAPISession);
	}

	void CMapiObjects::Logoff(_In_ HWND hwnd, ULONG ulFlags)
	{
		output::DebugPrint(output::DBGGeneric, L"Logging off of %p, ulFlags = 0x%08X\n", m_lpMAPISession, ulFlags);

		if (m_lpMAPISession)
		{
			EC_MAPI_S(m_lpMAPISession->Logoff(reinterpret_cast<ULONG_PTR>(hwnd), ulFlags, NULL));
			m_lpMAPISession->Release();
			m_lpMAPISession = nullptr;
		}
	}

	_Check_return_ LPMAPISESSION CMapiObjects::GetSession() const { return m_lpMAPISession; }

	_Check_return_ LPMAPISESSION CMapiObjects::LogonGetSession(_In_ HWND hWnd)
	{
		if (m_lpMAPISession) return m_lpMAPISession;

		MAPILogonEx(hWnd, nullptr, MAPI_EXTENDED | MAPI_LOGON_UI | MAPI_NEW_SESSION);

		return m_lpMAPISession;
	}

	void CMapiObjects::SetMDB(_In_opt_ LPMDB lpMDB)
	{
		output::DebugPrintEx(output::DBGGeneric, CLASS, L"SetMDB", L"replacing %p with %p\n", m_lpMDB, lpMDB);
		if (m_lpMDB) m_lpMDB->Release();
		m_lpMDB = lpMDB;
		if (m_lpMDB) m_lpMDB->AddRef();
	}

	_Check_return_ LPMDB CMapiObjects::GetMDB() const { return m_lpMDB; }

	void CMapiObjects::SetAddrBook(_In_opt_ LPADRBOOK lpAddrBook)
	{
		output::DebugPrintEx(
			output::DBGGeneric, CLASS, L"SetAddrBook", L"replacing %p with %p\n", m_lpAddrBook, lpAddrBook);
		if (m_lpAddrBook) m_lpAddrBook->Release();
		m_lpAddrBook = lpAddrBook;
		if (m_lpAddrBook) m_lpAddrBook->AddRef();
	}

	_Check_return_ LPADRBOOK CMapiObjects::GetAddrBook(bool bForceOpen)
	{
		// if we haven't opened the address book yet and we have a session, open it now
		if (!m_lpAddrBook && m_lpMAPISession && bForceOpen)
		{
			EC_MAPI_S(m_lpMAPISession->OpenAddressBook(NULL, nullptr, NULL, &m_lpAddrBook));
		}

		return m_lpAddrBook;
	}
} // namespace cache