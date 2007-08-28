// MAPIABFunctions.h : Stand alone MAPI Address Book functions

#pragma once

HRESULT AddOneOffAddress(
						 LPMAPISESSION lpMAPISession,
						 LPMESSAGE lpMessage,
						 LPCTSTR szDisplayName,
						 LPCTSTR szAddrType,
						 LPCTSTR szEmailAddress,
						 ULONG ulRecipientType);

HRESULT AddRecipient(
					 LPMAPISESSION lpMAPISession,
					 LPMESSAGE lpMessage,
					 LPCTSTR szName,
					 ULONG ulRecipientType);

HRESULT CreateANRRestriction(ULONG ulPropTag,
							 LPCTSTR szString,
							 LPVOID lpParent,
							 LPSRestriction* lppRes);

HRESULT GetABContainerTable(LPADRBOOK lpAdrBook, LPMAPITABLE* lpABContainerTable);

HRESULT HrAllocAdrList(ULONG ulNumProps, LPADRLIST* lpAdrList);

HRESULT	ManualResolve(
					  LPMAPISESSION lpMAPISession,
					  LPMESSAGE lpMessage,
					  LPCTSTR szName,
					  ULONG PropTagToCompare);

HRESULT SearchContentsTableForName(
									LPMAPITABLE pTable,
									LPCTSTR szName,
									ULONG PropTagToCompare,
									LPSPropValue *lppPropsFound);
