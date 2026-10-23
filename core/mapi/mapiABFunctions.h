// Stand alone MAPI Address Book functions
#pragma once

namespace mapi::ab
{
	bool IsMailboxSMTPAddress(const std::wstring& smtpAddress) noexcept;
	_Check_return_ HRESULT
	ResolveExchangeMailbox(_In_ LPADRBOOK addressBook, const std::wstring& smtpAddress, _Out_ std::wstring& mailboxDN);

	_Check_return_ HRESULT
	DisplayDetails(_In_ LPADRBOOK lpAddrBook, HWND hwnd, ULONG cbEntryID, _In_ LPENTRYID lpEntryID);

	_Check_return_ HRESULT AddOneOffAddress(
		_In_ LPMAPISESSION lpMAPISession,
		_In_ LPMESSAGE lpMessage,
		_In_ const std::wstring& szDisplayName,
		_In_ const std::wstring& szAddrType,
		_In_ const std::wstring& szEmailAddress,
		ULONG ulRecipientType);

	_Check_return_ HRESULT AddRecipient(
		_In_ LPMAPISESSION lpMAPISession,
		_In_ LPMESSAGE lpMessage,
		_In_ const std::wstring& szName,
		ULONG ulRecipientType);

	_Check_return_ LPSRestriction
	CreateANRRestriction(ULONG ulPropTag, _In_ const std::wstring& szString, _In_opt_ LPVOID lpParent);

	_Check_return_ LPMAPITABLE GetABContainerTable(_In_ LPADRBOOK lpAdrBook);

	_Check_return_ LPADRLIST AllocAdrList(ULONG ulNumProps);

	_Check_return_ HRESULT ManualResolve(
		_In_ LPMAPISESSION lpMAPISession,
		_In_ LPMESSAGE lpMessage,
		_In_ const std::wstring& szName,
		ULONG PropTagToCompare);

	_Check_return_ HRESULT SearchContentsTableForName(
		_In_ LPMAPITABLE pTable,
		_In_ const std::wstring& szName,
		ULONG PropTagToCompare,
		_Deref_out_opt_ LPSPropValue* lppPropsFound);

	_Check_return_ LPMAILUSER SelectUser(_In_ LPADRBOOK lpAdrBook, HWND hwnd, _Out_opt_ ULONG* lpulObjType);
} // namespace mapi::ab