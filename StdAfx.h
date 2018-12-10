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
#include <EMSAbTag.h>
#include <IMessage.h>
#include <edkguid.h>
#include <tnef.h>
#include <mapiaux.h>

#include <aclui.h>
#include <uxtheme.h>

// there's an odd conflict with mimeole.h and richedit.h - this should fix it
#ifdef UNICODE
#undef CHARFORMAT
#endif
#include <mimeole.h>
#ifdef UNICODE
#undef CHARFORMAT
#define CHARFORMAT CHARFORMATW
#endif

#include <Res/Resource.h> // main symbols

#include <IO/MFCOutput.h>
#include <IO/Registry.h>
#include <IO/Error.h>

#include <MFCMAPI.h>

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

// Definitions for WrapCompressedRTFStreamEx in param for WrapCompressedRTFStreamEX
// http://msdn2.microsoft.com/en-us/library/bb905293.aspx
struct RTF_WCSINFO
{
	ULONG size; // Size of the structure
	ULONG ulFlags;
	/****** MAPI_MODIFY ((ULONG) 0x00000001) above */
	/****** STORE_UNCOMPRESSED_RTF ((ULONG) 0x00008000) above */
	/****** MAPI_NATIVE_BODY ((ULONG) 0x00010000) mapidefs.h Only used for reading*/
	ULONG ulInCodePage; // Codepage of the message, used when passing MAPI_NATIVE_BODY, ignored otherwise
	ULONG ulOutCodePage; // Codepage of the Returned Stream, used when passing MAPI_NATIVE_BODY, ignored otherwise
};

// out param type information for WrapCompressedRTFStreamEX
// http://msdn2.microsoft.com/en-us/library/bb905294.aspx
struct RTF_WCSRETINFO
{
	ULONG size; // Size of the structure
	ULONG ulStreamFlags;
	/****** MAPI_NATIVE_BODY_TYPE_RTF ((ULONG) 0x00000001) mapidefs.h */
	/****** MAPI_NATIVE_BODY_TYPE_HTML ((ULONG) 0x00000002) mapidefs.h */
	/****** MAPI_NATIVE_BODY_TYPE_PLAINTEXT ((ULONG) 0x00000004) mapidefs.h */
};

// http://msdn.microsoft.com/en-us/library/office/dn433223.aspx
#pragma pack(4)
struct CONTAB_ENTRYID
{
	BYTE abFlags[4];
	MAPIUID muid;
	ULONG ulVersion;
	ULONG ulType;
	ULONG ulIndex;
	ULONG cbeid;
	BYTE abeid[1];
};
typedef CONTAB_ENTRYID* LPCONTAB_ENTRYID;
#pragma pack()

// http://msdn.microsoft.com/en-us/library/office/dn433221.aspx
#pragma pack(4)
struct DIR_ENTRYID
{
	BYTE abFlags[4];
	MAPIUID muid;
	ULONG ulVersion;
	ULONG ulType;
	MAPIUID muidID;
};
typedef DIR_ENTRYID* LPDIR_ENTRYID;
#pragma pack()

// http://msdn2.microsoft.com/en-us/library/bb820951.aspx
#define MAPI_IPROXYSTOREOBJECT_METHODS(IPURE) \
	MAPIMETHOD(PlaceHolder1) \
	() IPURE; \
	MAPIMETHOD(UnwrapNoRef) \
	(LPVOID * ppvObject) IPURE; \
	MAPIMETHOD(PlaceHolder2) \
	() IPURE;

// clang-format off
DECLARE_MAPI_INTERFACE_(IProxyStoreObject, IUnknown)
{
	BEGIN_INTERFACE
		MAPI_IUNKNOWN_METHODS(PURE)
		MAPI_IPROXYSTOREOBJECT_METHODS(PURE)
};
// clang-format on

#ifndef MAPI_IMAPICLIENTSHUTDOWN_METHODS
// http://blogs.msdn.com/stephen_griffin/archive/2009/03/03/fastest-shutdown-in-the-west.aspx
DECLARE_MAPI_INTERFACE_PTR(IMAPIClientShutdown, LPMAPICLIENTSHUTDOWN);
#define MAPI_IMAPICLIENTSHUTDOWN_METHODS(IPURE) \
	MAPIMETHOD(QueryFastShutdown) \
	(THIS) IPURE; \
	MAPIMETHOD(NotifyProcessShutdown) \
	(THIS) IPURE; \
	MAPIMETHOD(DoFastShutdown) \
	(THIS) IPURE;

// clang-format off
DECLARE_MAPI_INTERFACE_(IMAPIClientShutdown, IUnknown)
{
	BEGIN_INTERFACE
		MAPI_IUNKNOWN_METHODS(PURE)
		MAPI_IMAPICLIENTSHUTDOWN_METHODS(PURE)
};
// clang-format on
#define _IID_IMAPIClientShutdown_MISSING_IN_HEADER
#endif // MAPI_IMAPICLIENTSHUTDOWN_METHODS

#ifndef MAPI_IMAPIPROVIDERSHUTDOWN_METHODS
/* IMAPIProviderShutdown Interface --------------------------------------- */
DECLARE_MAPI_INTERFACE_PTR(IMAPIProviderShutdown, LPMAPIPROVIDERSHUTDOWN);

#define MAPI_IMAPIPROVIDERSHUTDOWN_METHODS(IPURE) \
	MAPIMETHOD(QueryFastShutdown) \
	(THIS) IPURE; \
	MAPIMETHOD(NotifyProcessShutdown) \
	(THIS) IPURE; \
	MAPIMETHOD(DoFastShutdown) \
	(THIS) IPURE;

// clang-format off
DECLARE_MAPI_INTERFACE_(IMAPIProviderShutdown, IUnknown)
{
	BEGIN_INTERFACE
		MAPI_IUNKNOWN_METHODS(PURE)
		MAPI_IMAPIPROVIDERSHUTDOWN_METHODS(PURE)
};
// clang-format on
#endif // MAPI_IMAPIPROVIDERSHUTDOWN_METHODS

// for CompareStrings
static DWORD g_lcid = MAKELCID(MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US), SORT_DEFAULT);

// In case we are compiling against an older version of headers

// http://msdn2.microsoft.com/en-us/library/bb820933.aspx
#define MAPI_IATTACHMENTSECURITY_METHODS(IPURE) \
	MAPIMETHOD(IsAttachmentBlocked) \
	(LPCWSTR pwszFileName, BOOL * pfBlocked) IPURE;

// clang-format off
DECLARE_MAPI_INTERFACE_(IAttachmentSecurity, IUnknown)
{
	BEGIN_INTERFACE
		MAPI_IUNKNOWN_METHODS(PURE)
		MAPI_IATTACHMENTSECURITY_METHODS(PURE)
};
// clang-format on

/* Indexing notifications (used for FTE related communications) */
/* Shares EXTENDED_NOTIFICATION to pass structures below, */
/* but NOTIFICATION type will be fnevIndexing */

// Sometimes IExchangeManageStore5 is in edkmdb.h, sometimes it isn't
#ifndef EXCHANGE_IEXCHANGEMANAGESTORE5_METHODS
#define USES_IID_IExchangeManageStore5

/*------------------------------------------------------------------------
*
* 'IExchangeManageStore5' Interface Declaration
*
* Used for store management functions.
*
*-----------------------------------------------------------------------*/

#define EXCHANGE_IEXCHANGEMANAGESTORE5_METHODS(IPURE) \
	MAPIMETHOD(GetMailboxTableEx) \
	(THIS_ LPSTR lpszServerName, LPGUID lpguidMdb, LPMAPITABLE * lppTable, ULONG ulFlags, UINT uOffset) IPURE; \
	MAPIMETHOD(GetPublicFolderTableEx) \
	(THIS_ LPSTR lpszServerName, LPGUID lpguidMdb, LPMAPITABLE * lppTable, ULONG ulFlags, UINT uOffset) IPURE;

#undef INTERFACE
#define INTERFACE IExchangeManageStore5
// clang-format off
DECLARE_MAPI_INTERFACE_(IExchangeManageStore5, IUnknown)
{
	MAPI_IUNKNOWN_METHODS(PURE)
		EXCHANGE_IEXCHANGEMANAGESTORE_METHODS(PURE)
		EXCHANGE_IEXCHANGEMANAGESTORE2_METHODS(PURE)
		EXCHANGE_IEXCHANGEMANAGESTORE3_METHODS(PURE)
		EXCHANGE_IEXCHANGEMANAGESTORE4_METHODS(PURE)
		EXCHANGE_IEXCHANGEMANAGESTORE5_METHODS(PURE)
};
// clang-format on
#undef IMPL
#define IMPL

DECLARE_MAPI_INTERFACE_PTR(IExchangeManageStore5, LPEXCHANGEMANAGESTORE5);
#endif // #ifndef EXCHANGE_IEXCHANGEMANAGESTORE5_METHODS

#define EXCHANGE_IEXCHANGEMANAGESTOREEX_METHODS(IPURE) \
	MAPIMETHOD(CreateStoreEntryID2) \
	(THIS_ ULONG cValues, LPSPropValue lpPropArray, ULONG ulFlags, ULONG * lpcbEntryID, LPENTRYID * lppEntryID) IPURE;

#undef INTERFACE
#define INTERFACE IExchangeManageStoreEx
// clang-format off
DECLARE_MAPI_INTERFACE_(IExchangeManageStoreEx, IUnknown)
{
	MAPI_IUNKNOWN_METHODS(PURE)
		EXCHANGE_IEXCHANGEMANAGESTORE_METHODS(PURE)
		EXCHANGE_IEXCHANGEMANAGESTOREEX_METHODS(PURE)
};
// clang-format on
#undef IMPL
#define IMPL

DECLARE_MAPI_INTERFACE_PTR(IExchangeManageStoreEx, LPEXCHANGEMANAGESTOREEX);

#define CbNewROWLIST(_centries) (offsetof(ROWLIST, aEntries) + (_centries) * sizeof(ROWENTRY))
#define MAXNewROWLIST ((ULONG_MAX - offsetof(ROWLIST, aEntries)) / sizeof(ROWENTRY))
#define MAXMessageClassArray ((ULONG_MAX - offsetof(SMessageClassArray, aMessageClass)) / sizeof(LPCSTR))
#define MAXNewADRLIST ((ULONG_MAX - offsetof(ADRLIST, aEntries)) / sizeof(ADRENTRY))