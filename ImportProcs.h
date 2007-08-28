// ImportProcs.h : header file for loading imports from DLLs
//

#pragma once

#include "stdafx.h"
#include "MAPIWz.h"

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

BOOL GetComponentPath(
					  LPSTR szComponent,
					  LPTSTR szQualifier,
					  TCHAR* szDllPath,
					  DWORD cchDLLPath);

HMODULE LoadFromSystemDir(LPTSTR szDLLName);

HMODULE MyLoadLibrary(LPCTSTR lpLibFileName);
void LoadRichEd();
void LoadAclUI();
void LoadStg();
void LoadThemeUI();
void LoadMimeOLE();

void AutoLoadMAPI();
void UnloadMAPI();
void LoadMAPIFuncs(HMODULE hMod);
