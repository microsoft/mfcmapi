// MAPIABFunctions.h : Stand alone MAPI Address Book functions
#pragma once

_Check_return_ HRESULT AddOneOffAddress(
	_In_ LPMAPISESSION lpMAPISession,
	_In_ LPMESSAGE lpMessage,
	_In_ wstring& szDisplayName,
	_In_ wstring& szAddrType,
	_In_ wstring& szEmailAddress,
	ULONG ulRecipientType);

_Check_return_ HRESULT AddRecipient(
	_In_ LPMAPISESSION lpMAPISession,
	_In_ LPMESSAGE lpMessage,
	_In_z_ LPCTSTR szName,
	ULONG ulRecipientType);

_Check_return_ HRESULT CreateANRRestriction(ULONG ulPropTag,
	wstring szString,
	_In_opt_ LPVOID lpParent,
	_Deref_out_opt_ LPSRestriction* lppRes);

_Check_return_ HRESULT GetABContainerTable(_In_ LPADRBOOK lpAdrBook, _Deref_out_opt_ LPMAPITABLE* lpABContainerTable);

_Check_return_ HRESULT HrAllocAdrList(ULONG ulNumProps, _Deref_out_opt_ LPADRLIST* lpAdrList);

_Check_return_ HRESULT ManualResolve(
	_In_ LPMAPISESSION lpMAPISession,
	_In_ LPMESSAGE lpMessage,
	_In_ wstring& szName,
	ULONG PropTagToCompare);

_Check_return_ HRESULT SearchContentsTableForName(
	_In_ LPMAPITABLE pTable,
	_In_ wstring& szName,
	ULONG PropTagToCompare,
	_Deref_out_opt_ LPSPropValue *lppPropsFound);

_Check_return_ HRESULT SelectUser(_In_ LPADRBOOK lpAdrBook, HWND hwnd, _Out_opt_ ULONG* lpulObjType, _Deref_out_opt_ LPMAILUSER* lppMailUser);
