#pragma once

#define VC_EXTRALEAN // Exclude rarely-used stuff from Windows headers

#pragma warning(push)
#pragma warning(disable : 4995)
#include <cstdio>
#include <cstring>
#include <cwchar>
#pragma warning(pop)

// Common CRT headers
#include <list>
#include <algorithm>
#include <locale>
#include <sstream>
#include <iterator>
#include <functional>
#include <map>
#include <utility>
#include <vector>
#include <cassert>
#include <stack>
#include <deque>
#include <thread>

// Speed up our string conversions for output
#ifdef MRMAPI
#define _CRT_DISABLE_PERFCRIT_LOCKS
#endif

#include <sal.h>
// A bug in annotations in shobjidl.h forces us to disable 6387 to include afxwin.h
#pragma warning(push)
#pragma warning(disable : 6387)
#include <afxwin.h> // MFC core and standard components
#pragma warning(pop)
#include <afxcmn.h> // MFC support for Windows Common Controls

#pragma warning(push)
#pragma warning(disable : 4091)
#include <ShlObj.h>
#pragma warning(pop)

// Fix a build issue with a few versions of the MAPI headers
#if !defined(FREEBUFFER_DEFINED)
typedef ULONG(STDAPICALLTYPE FREEBUFFER)(LPVOID lpBuffer);
#define FREEBUFFER_DEFINED
#endif

#include <MAPIX.h>
#include <MAPIUtil.h>
#include <MAPIForm.h>
#include <MAPIWz.h>
#include <MAPIHook.h>
#include <MSPST.h>

#include <EdkMdb.h>
#include <ExchForm.h>
#include <EMSABTAG.H>
#include <IMessage.h>
#include <EdkGuid.h>
#include <TNEF.h>
#include <MAPIAux.h>

#include <AclUI.h>
#include <Uxtheme.h>

// there's an odd conflict with mimeole.h and richedit.h - this should fix it
#ifdef UNICODE
#undef CHARFORMAT
#endif
#include <mimeole.h>
#ifdef UNICODE
#undef CHARFORMAT
#define CHARFORMAT CHARFORMATW
#endif

#include <core/res/Resource.h> // main symbols

#include <core/utility/error.h>

// Custom messages - used to ensure actions occur on the right threads.

// Used by CAdviseSink:
#define WM_MFCMAPI_ADDITEM (WM_APP + 1)
#define WM_MFCMAPI_DELETEITEM (WM_APP + 2)
#define WM_MFCMAPI_MODIFYITEM (WM_APP + 3)
#define WM_MFCMAPI_REFRESHTABLE (WM_APP + 4)

// Used by DwThreadFuncLoadTable
#define WM_MFCMAPI_THREADADDITEM (WM_APP + 5)
#define WM_MFCMAPI_UPDATESTATUSBAR (WM_APP + 6)
#define WM_MFCMAPI_CLEARSINGLEMAPIPROPLIST (WM_APP + 7)

// Used by CSingleMAPIPropListCtrl and CSortHeader
#define WM_MFCMAPI_SAVECOLUMNORDERHEADER (WM_APP + 10)
#define WM_MFCMAPI_SAVECOLUMNORDERLIST (WM_APP + 11)

// Used by CContentsTableDlg
#define WM_MFCMAPI_RESETCOLUMNS (WM_APP + 12)