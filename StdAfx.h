#pragma once

// clang-format off
#pragma warning(disable : 26426) // Warning C26426 Global initializer calls a non-constexpr (i.22)
#pragma warning(disable : 26446) // Warning C26446 Prefer to use gsl::at() instead of unchecked subscript operator (bounds.4).
#pragma warning(disable : 26481) // Warning C26481 Don't use pointer arithmetic. Use span instead (bounds.1).
#pragma warning(disable : 26485) // Warning C26485 Expression '': No array to pointer decay (bounds.3).
#pragma warning(disable : 26487) // Warning C26487 Don't return a pointer '' that may be invalid (lifetime.4).
// clang-format on

#define VC_EXTRALEAN // Exclude rarely-used stuff from Windows headers

#pragma warning(push)
#pragma warning(disable : 4995) // Warning C4995 'function': name was marked as #pragma deprecated
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
#pragma warning( \
	disable : 6387) // Warning C6387 'argument' may be 'value': this does not adhere to the specification for the function 'function name': Lines: x, y
#include <afxwin.h> // MFC core and standard components
#pragma warning(pop)
#include <afxcmn.h> // MFC support for Windows Common Controls

#pragma warning(push)
#pragma warning(disable : 4091) // Warning C4091 'keyword' : ignored on left of 'type' when no variable is declared
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
#include <actMgmt.h>
#include <IMSCapabilities.h>

// For import procs
#include <AclUI.h>
#include <Uxtheme.h>
#include <dpapi.h>

// there's an odd conflict with mimeole.h and richedit.h - this should fix it
#ifdef UNICODE
#undef CHARFORMAT
#endif
#pragma warning(push)
#pragma warning(disable : 28251) // Warning C28251 Inconsistent annotation for function: this instance has an error
#include <mimeole.h>
#pragma warning(pop)
#ifdef UNICODE
#undef CHARFORMAT
#define CHARFORMAT CHARFORMATW
#endif

#include <core/res/Resource.h> // main symbols

#include <core/utility/error.h>

// Custom messages - used to ensure actions occur on the right threads.

// Used by OnNotify:
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

namespace cache
{
	constexpr ULONG ulNoMatch = 0xffffffff;
}