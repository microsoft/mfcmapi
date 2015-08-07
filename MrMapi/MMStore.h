#pragma once
// MMStore.h : Store Properties for MrMAPI

HRESULT HrMAPIOpenStoreAndFolder(
	_In_ LPMAPISESSION lpMAPISession,
	_In_ ULONG ulFolder,
	_In_ wstring lpszFolderPath,
	_Out_opt_ LPMDB* lppMDB,
	_Deref_out_opt_ LPMAPIFOLDER* lppFolder);

void PrintObjectProperties(_In_z_ LPCTSTR szObjType, _In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag);

void DoStore(_In_ MYOPTIONS ProgOpts);

HRESULT OpenStore(_In_ LPMAPISESSION lpMAPISession, ULONG ulIndex, _Out_ LPMDB* lppMDB);