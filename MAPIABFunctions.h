// MAPIABFunctions.h : Stand alone MAPI Address Book functions

#pragma once

_Check_return_ HRESULT AddOneOffAddress(
										_In_ LPMAPISESSION lpMAPISession,
										_In_ LPMESSAGE lpMessage,
										_In_z_ LPCTSTR szDisplayName,
										_In_z_ LPCTSTR szAddrType,
										_In_z_ LPCTSTR szEmailAddress,
										ULONG ulRecipientType);

_Check_return_ HRESULT AddRecipient(
									_In_ LPMAPISESSION lpMAPISession,
									_In_ LPMESSAGE lpMessage,
									_In_z_ LPCTSTR szName,
									ULONG ulRecipientType);

_Check_return_ HRESULT CreateANRRestriction(ULONG ulPropTag,
											_In_z_ LPCTSTR szString,
											_In_opt_ LPVOID lpParent,
											_Deref_out_opt_ LPSRestriction* lppRes);

_Check_return_ HRESULT GetABContainerTable(_In_ LPADRBOOK lpAdrBook, _Deref_out_opt_ LPMAPITABLE* lpABContainerTable);

_Check_return_ HRESULT HrAllocAdrList(ULONG ulNumProps, _Deref_out_opt_ LPADRLIST* lpAdrList);

_Check_return_ HRESULT ManualResolve(
									 _In_ LPMAPISESSION lpMAPISession,
									 _In_ LPMESSAGE lpMessage,
									 _In_z_ LPCTSTR szName,
									 ULONG PropTagToCompare);

_Check_return_ HRESULT SearchContentsTableForName(
	_In_ LPMAPITABLE pTable,
	_In_z_ LPCTSTR szName,
	ULONG PropTagToCompare,
	_Deref_out_opt_ LPSPropValue *lppPropsFound);
