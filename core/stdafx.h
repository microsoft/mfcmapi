#pragma once

#define VC_EXTRALEAN // Exclude rarely-used stuff from Windows headers

#pragma warning(push)
#pragma warning(disable : 4995)
#include <cstdio>
#include <cstring>
#include <cwchar>
#pragma warning(pop)

#include <list>
#include <algorithm>

// Speed up our string conversions for output
//#ifdef MRMAPI
//#define _CRT_DISABLE_PERFCRIT_LOCKS
//#endif

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
