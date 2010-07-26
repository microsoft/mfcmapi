// ImportProcs.h : header file for loading imports from DLLs
//

#pragma once

extern HMODULE	hModMSMAPI;
extern HMODULE	hModMAPI;
extern LPWRAPCOMPRESSEDRTFSTREAMEX	pfnWrapEx;
extern LPLAUNCHWIZARDENTRY			pfnLaunchWizard;
extern LPEDITSECURITY				pfnEditSecurity;
extern LPSTGCREATESTORAGEEX			pfnStgCreateStorageEx;
extern LPOPENTHEMEDATA				pfnOpenThemeData;
extern LPCLOSETHEMEDATA				pfnCloseThemeData;
extern LPGETTHEMEMARGINS			pfnGetThemeMargins;
extern LPMIMEOLEGETCODEPAGECHARSET	pfnMimeOleGetCodePageCharset;
extern LPMSIPROVIDEQUALIFIEDCOMPONENT pfnMsiProvideQualifiedComponent;
extern LPMSIGETFILEVERSION			pfnMsiGetFileVersion;

_Check_return_ HMODULE LoadFromSystemDir(_In_z_ LPTSTR szDLLName);

_Check_return_ HMODULE MyLoadLibrary(_In_z_ LPCTSTR lpszLibFileName);
void LoadRichEd();
void ImportProcs();

void GetMAPIPath(_In_opt_z_ LPCTSTR szClient, _Inout_z_count_(cchMAPIPath) LPTSTR szMAPIPath, ULONG cchMAPIPath);
void AutoLoadMAPI();
void UnloadMAPI();
void LoadMAPIFuncs(HMODULE hMod);

_Check_return_ STDAPI HrCopyRestriction(
						 _In_ LPSRestriction lpResSrc, // source restriction ptr
						 _In_opt_ LPVOID lpObject, // ptr to existing MAPI buffer
						 _In_ LPSRestriction FAR * lppResDest // dest restriction buffer ptr
						 );

_Check_return_ HRESULT HrCopyRestrictionArray(
							   _In_ LPSRestriction lpResSrc, // source restriction
							   _In_ LPVOID lpObject, // ptr to existing MAPI buffer
							   ULONG cRes, // # elements in array
							   _In_count_(cRes) LPSRestriction lpResDest // destination restriction
							   );

_Check_return_ STDMETHODIMP MyOpenStreamOnFile(_In_ LPALLOCATEBUFFER lpAllocateBuffer,
								_In_ LPFREEBUFFER lpFreeBuffer,
								ULONG ulFlags,
								_In_z_ LPCWSTR lpszFileName,
								_In_opt_z_ LPCWSTR /*lpszPrefix*/,
								_Out_ LPSTREAM FAR * lppStream);

void WINAPI MyHeapSetInformation(_In_opt_ HANDLE HeapHandle,
								 _In_ HEAP_INFORMATION_CLASS HeapInformationClass,
								 _In_opt_count_(HeapInformationLength) PVOID HeapInformation,
								 _In_ SIZE_T HeapInformationLength);