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

HMODULE LoadFromSystemDir(LPTSTR szDLLName);

HMODULE MyLoadLibrary(LPCTSTR lpLibFileName);
void LoadRichEd();
void ImportProcs();

void GetMAPIPath(LPTSTR szClient, LPTSTR szMAPIPath, ULONG cchMAPIPath);
void AutoLoadMAPI();
void UnloadMAPI();
void LoadMAPIFuncs(HMODULE hMod);

STDAPI HrCopyRestriction(
						 LPSRestriction lpResSrc, // source restriction ptr
						 LPVOID lpObject, // ptr to existing MAPI buffer
						 LPSRestriction FAR * lppResDest // dest restriction buffer ptr
						 );

HRESULT HrCopyRestrictionArray(
							   LPSRestriction lpResSrc, // source restriction
							   LPVOID lpObject, // ptr to existing MAPI buffer
							   ULONG cRes, // # elements in array
							   LPSRestriction lpResDest // destination restriction
							   );

STDMETHODIMP MyOpenStreamOnFile(LPALLOCATEBUFFER lpAllocateBuffer,
								LPFREEBUFFER lpFreeBuffer,
								ULONG ulFlags,
								__in LPCWSTR lpszFileName,
								__in LPCWSTR /*lpszPrefix*/,
								LPSTREAM FAR * lppStream);