// Stand alone MAPI Address Book functions
#pragma once

namespace mapi
{
	namespace ab
	{
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

		_Check_return_ HRESULT CreateANRRestriction(
			ULONG ulPropTag,
			_In_ const std::wstring& szString,
			_In_opt_ LPVOID lpParent,
			_Deref_out_opt_ LPSRestriction* lppRes);

		_Check_return_ HRESULT
		GetABContainerTable(_In_ LPADRBOOK lpAdrBook, _Deref_out_opt_ LPMAPITABLE* lpABContainerTable);

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
	}
}