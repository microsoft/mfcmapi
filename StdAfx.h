// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//  are changed infrequently
//

#pragma once

#define VC_EXTRALEAN	// Exclude rarely-used stuff from Windows headers

#include <sal.h>
#include <afxwin.h>		// MFC core and standard components
#include <afxext.h>		// MFC extensions
#include <afxcmn.h>		// MFC support for Windows Common Controls

// Safe String handling header
#include <strsafe.h>

// Fix a build issue with a few versions of the MAPI headers
#if !defined(FREEBUFFER_DEFINED)
typedef ULONG (STDAPICALLTYPE FREEBUFFER)(
	LPVOID			lpBuffer
	);
#define FREEBUFFER_DEFINED
#endif

#include <MapiX.h>
#include <MapiUtil.h>
#include <MAPIform.h>
#include <MAPIWz.h>
#include <MAPIHook.h>
#include <MSPST.h>

#include <edkmdb.h>
#include <exchform.h>
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

// Some trees use bldver.h to maintain some definitions. Do not remove.
#include "bldver.h"
#include "resource.h"		// main symbols

#include "MFCOutput.h"
#include "registry.h"
#include "error.h"

#include "MFCMAPI.h"

struct TagNames
{
	ULONG ulMatchingTableColumn;
	UINT uidName;
};

struct _ContentsData
{
	LPSBinary		lpEntryID; // Allocated with MAPIAllocateMore
	LPSBinary		lpLongtermID; // Allocated with MAPIAllocateMore
	LPSBinary		lpInstanceKey; // Allocated with MAPIAllocateMore
	LPSBinary		lpServiceUID; // Allocated with MAPIAllocateMore
	LPSBinary		lpProviderUID; // Allocated with MAPIAllocateMore
	TCHAR*			szDN; // Allocated with MAPIAllocateMore
	CHAR*			szProfileDisplayName; // Allocated with MAPIAllocateMore
	ULONG			ulAttachNum;
	ULONG			ulAttachMethod;
	ULONG			ulRowID; // for recipients
	ULONG			ulRowType; // PR_ROW_TYPE
};

struct _PropListData
{
	ULONG		ulPropTag;
};

struct _MVPropData
{
	_PV val; // Allocated with MAPIAllocateMore
};

struct _TagData
{
	ULONG		ulPropTag;
};

struct _ResData
{
	LPSRestriction lpOldRes; // not allocated - just a pointer
	LPSRestriction lpNewRes; // Owned by an alloc parent - do not free
};

struct _CommentData
{
	LPSPropValue lpOldProp; // not allocated - just a pointer
	LPSPropValue lpNewProp; // Owned by an alloc parent - do not free
};

struct _BinaryData
{
	SBinary OldBin; // not allocated - just a pointer
	SBinary NewBin; // MAPIAllocateMore from m_lpNewEntryList
};

class CAdviseSink;

struct _NodeData
{
	LPSBinary			lpEntryID; // Allocated with MAPIAllocateMore
	LPSBinary			lpInstanceKey; // Allocated with MAPIAllocateMore
	LPMAPITABLE			lpHierarchyTable; // Object - free with Release
	CAdviseSink*		lpAdviseSink; // Object - free with Release
	ULONG_PTR			ulAdviseConnection;
	ULONG				bSubfolders; // this is intentionally ULONG instead of bool
	ULONG				ulContainerFlags;
};

struct SortListData
{
	TCHAR*				szSortText; // Allocated with MAPIAllocateBuffer
	ULARGE_INTEGER		ulSortValue;
	ULONG				ulSortDataType;
	union
	{
		_ContentsData	Contents; // SORTLIST_CONTENTS
		_PropListData	Prop; // SORTLIST_PROP
		_MVPropData		MV; // SORTLIST_MVPROP
		_TagData		Tag; // SORTLIST_TAGARRAY
		_ResData		Res; // SORTLIST_RES
		_CommentData	Comment; // SORTLIST_COMMENT
		_BinaryData		Binary; // SORTLIST_BINARY
		_NodeData		Node; // SORTLIST_TREENODE

	} data;
	ULONG			cSourceProps;
	LPSPropValue	lpSourceProps; // Stolen from lpsRowData in CContentsTableListCtrl::BuildDataItem - free with MAPIFreeBuffer
	bool			bItemFullyLoaded;
};

// Macros to assist in OnInitMenu
#define CHECK(state) ((state)?MF_CHECKED:MF_UNCHECKED)
#define DIM(state) ((state)?MF_ENABLED:MF_GRAYED)
#define DIMMSOK(iNumSelected) ((iNumSelected>=1)?MF_ENABLED:MF_GRAYED)
#define DIMMSNOK(iNumSelected) ((iNumSelected==1)?MF_ENABLED:MF_GRAYED)

// Various flags gleaned from product documentation and KB articles
// http://msdn2.microsoft.com/en-us/library/ms526744.aspx
#define STORE_HTML_OK			((ULONG) 0x00010000)
#define STORE_ANSI_OK			((ULONG) 0x00020000)
#define STORE_LOCALSTORE		((ULONG) 0x00080000)

// http://msdn2.microsoft.com/en-us/library/ms531462.aspx
#define ATT_INVISIBLE_IN_HTML	((ULONG) 0x00000001)
#define ATT_INVISIBLE_IN_RTF	((ULONG) 0x00000002)
#define ATT_MHTML_REF			((ULONG) 0x00000004)

// http://msdn2.microsoft.com/en-us/library/ms527629.aspx
#define	MSGFLAG_ORIGIN_X400		((ULONG) 0x00001000)
#define	MSGFLAG_ORIGIN_INTERNET	((ULONG) 0x00002000)
#define MSGFLAG_ORIGIN_MISC_EXT ((ULONG) 0x00008000)
#define MSGFLAG_OUTLOOK_NON_EMS_XP ((ULONG) 0x00010000)

// http://msdn2.microsoft.com/en-us/library/ms528848.aspx
#define	MSGSTATUS_DRAFT			((ULONG) 0x00000100)
#define	MSGSTATUS_ANSWERED		((ULONG) 0x00000200)

#define ENCODING_PREFERENCE			((ULONG) 0x00020000)
#define ENCODING_MIME				((ULONG) 0x00040000)
#define BODY_ENCODING_HTML			((ULONG) 0x00080000)
#define BODY_ENCODING_TEXT_AND_HTML	((ULONG) 0x00100000)
#define MAC_ATTACH_ENCODING_UUENCODE	((ULONG) 0x00200000)
#define MAC_ATTACH_ENCODING_APPLESINGLE	((ULONG) 0x00400000)
#define MAC_ATTACH_ENCODING_APPLEDOUBLE	((ULONG) 0x00600000)

// Custom messages - used to ensure actions occur on the right threads.

// Used by CAdviseSink:
#define WM_MFCMAPI_ADDITEM					WM_APP+1
#define WM_MFCMAPI_DELETEITEM				WM_APP+2
#define WM_MFCMAPI_MODIFYITEM				WM_APP+3
#define WM_MFCMAPI_REFRESHTABLE				WM_APP+4

// Used by DwThreadFuncLoadTable
#define WM_MFCMAPI_THREADADDITEM			WM_APP+5
#define WM_MFCMAPI_UPDATESTATUSBAR			WM_APP+6
#define WM_MFCMAPI_CLEARSINGLEMAPIPROPLIST	WM_APP+7

// Used by CSingleMAPIPropListCtrl and CSortHeader
#define WM_MFCMAPI_SAVECOLUMNORDERHEADER	WM_APP+10
#define WM_MFCMAPI_SAVECOLUMNORDERLIST		WM_APP+11

// Definitions for WrapCompressedRTFStreamEx in param for WrapCompressedRTFStreamEX
// http://msdn2.microsoft.com/en-us/library/bb905293.aspx
struct RTF_WCSINFO {
	ULONG		size;		// Size of the structure
	ULONG		ulFlags;
	/****** MAPI_MODIFY					((ULONG) 0x00000001) above */
	/****** STORE_UNCOMPRESSED_RTF		((ULONG) 0x00008000) above */
	/****** MAPI_NATIVE_BODY			((ULONG) 0x00010000) mapidefs.h   Only used for reading*/
	ULONG		ulInCodePage;	// Codepage of the message, used when passing MAPI_NATIVE_BODY, ignored otherwise
	ULONG		ulOutCodePage;	// Codepage of the Returned Stream, used when passing MAPI_NATIVE_BODY, ignored otherwise
};

// out param type information for WrapCompressedRTFStreamEX
// http://msdn2.microsoft.com/en-us/library/bb905294.aspx
struct RTF_WCSRETINFO {
	ULONG       size;	// Size of the structure
	ULONG		ulStreamFlags;
	/****** MAPI_NATIVE_BODY_TYPE_RTF			((ULONG) 0x00000001) mapidefs.h */
	/****** MAPI_NATIVE_BODY_TYPE_HTML			((ULONG) 0x00000002) mapidefs.h */
	/****** MAPI_NATIVE_BODY_TYPE_PLAINTEXT		((ULONG) 0x00000004) mapidefs.h */
};

#define MAPI_NATIVE_BODY					0x00010000

/* out param type infomation for WrapCompressedRTFStreamEx */
#define MAPI_NATIVE_BODY_TYPE_RTF			0x00000001
#define MAPI_NATIVE_BODY_TYPE_HTML			0x00000002
#define MAPI_NATIVE_BODY_TYPE_PLAINTEXT		0x00000004

// For EditSecurity
typedef bool (STDAPICALLTYPE EDITSECURITY)
(
 HWND hwndOwner,
 LPSECURITYINFO psi
 );
typedef EDITSECURITY* LPEDITSECURITY;

// For StgCreateStorageEx
typedef HRESULT (STDAPICALLTYPE STGCREATESTORAGEEX)
(
 IN const WCHAR* pwcsName,
 IN  DWORD grfMode,
 IN  DWORD stgfmt,              // enum
 IN  DWORD grfAttrs,             // reserved
 IN  STGOPTIONS * pStgOptions,
 IN  void * reserved,
 IN  REFIID riid,
 OUT void ** ppObjectOpen);
typedef STGCREATESTORAGEEX* LPSTGCREATESTORAGEEX;

// For Themes
typedef HTHEME (STDMETHODCALLTYPE OPENTHEMEDATA)
(
 HWND hwnd,
 LPCWSTR pszClassList);
typedef OPENTHEMEDATA* LPOPENTHEMEDATA;

typedef HTHEME (STDMETHODCALLTYPE CLOSETHEMEDATA)
(
 HTHEME hTheme);
typedef CLOSETHEMEDATA* LPCLOSETHEMEDATA;

typedef HRESULT (STDMETHODCALLTYPE GETTHEMEMARGINS)
(
 HTHEME hTheme,
 OPTIONAL HDC hdc,
 int iPartId,
 int iStateId,
 int iPropId,
 OPTIONAL RECT *prc,
 OUT MARGINS *pMargins);
typedef GETTHEMEMARGINS* LPGETTHEMEMARGINS;

typedef HRESULT (STDMETHODCALLTYPE MSIPROVIDEQUALIFIEDCOMPONENT)
(
 LPCTSTR     szCategory,
 LPCTSTR     szQualifier,
 DWORD       dwInstallMode,
 LPTSTR      lpPathBuf,
 LPDWORD     pcchPathBuf
 );
typedef MSIPROVIDEQUALIFIEDCOMPONENT* LPMSIPROVIDEQUALIFIEDCOMPONENT;

typedef HRESULT (STDMETHODCALLTYPE MSIGETFILEVERSION)
(
 LPCTSTR   szFilePath,
 LPTSTR    lpVersionBuf,
 LPDWORD   pcchVersionBuf,
 LPTSTR    lpLangBuf,
 LPDWORD   pcchLangBuf
 );
typedef MSIGETFILEVERSION* LPMSIGETFILEVERSION;

// http://msdn2.microsoft.com/en-us/library/bb820924.aspx
#pragma pack(4)
struct CONTAB_ENTRYID
{
	BYTE misc1[4];
	MAPIUID misc2;
	ULONG misc3;
	ULONG misc4;
	ULONG misc5;
	// EntryID of contact in store.
	ULONG cbeid;
	BYTE abeid[1];
};
typedef CONTAB_ENTRYID* LPCONTAB_ENTRYID;
#pragma pack()

// http://msdn2.microsoft.com/en-us/library/bb820951.aspx
#define MAPI_IPROXYSTOREOBJECT_METHODS(IPURE) \
	MAPIMETHOD(PlaceHolder1) \
	() IPURE; \
	MAPIMETHOD(UnwrapNoRef) \
	(LPVOID *ppvObject) IPURE; \
	MAPIMETHOD(PlaceHolder2) \
	() IPURE;

DECLARE_MAPI_INTERFACE_(IProxyStoreObject, IUnknown)
{
	BEGIN_INTERFACE
		MAPI_IUNKNOWN_METHODS(PURE)
		MAPI_IPROXYSTOREOBJECT_METHODS(PURE)
};

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

DECLARE_MAPI_INTERFACE_(IMAPIClientShutdown, IUnknown)
{
	BEGIN_INTERFACE
		MAPI_IUNKNOWN_METHODS(PURE)
		MAPI_IMAPICLIENTSHUTDOWN_METHODS(PURE)
};
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

DECLARE_MAPI_INTERFACE_(IMAPIProviderShutdown, IUnknown)
{
	BEGIN_INTERFACE
		MAPI_IUNKNOWN_METHODS(PURE)
		MAPI_IMAPIPROVIDERSHUTDOWN_METHODS(PURE)
};
#endif // MAPI_IMAPIPROVIDERSHUTDOWN_METHODS

// for CompareStrings
static DWORD g_lcid = MAKELCID(MAKELANGID(LANG_ENGLISH, SUBLANG_ENGLISH_US), SORT_DEFAULT);

// In case we are compiling against an older version of headers

#if !defined ACLTABLE_FREEBUSY
#define ACLTABLE_FREEBUSY		((ULONG) 0x00000002)
#endif // ACLTABLE_FREEBUSY

#if !defined frightsFreeBusySimple
#define frightsFreeBusySimple	0x0000800L
#endif // frightsFreeBusySimple

#if !defined frightsFreeBusyDetailed
#define frightsFreeBusyDetailed 0x0001000L
#endif // frightsFreeBusyDetailed

#if !defined fsdrightFreeBusySimple
#define fsdrightFreeBusySimple		0x00000001
#endif // fsdrightFreeBusySimple

#if !defined fsdrightFreeBusyDetailed
#define fsdrightFreeBusyDetailed	0x00000002
#endif // fsdrightFreeBusyDetailed

// http://msdn2.microsoft.com/en-us/library/bb820933.aspx
#define MAPI_IATTACHMENTSECURITY_METHODS(IPURE) \
	MAPIMETHOD(IsAttachmentBlocked) \
	(LPCWSTR pwszFileName, bool *pfBlocked) IPURE;

DECLARE_MAPI_INTERFACE_(IAttachmentSecurity, IUnknown)
{
	BEGIN_INTERFACE
		MAPI_IUNKNOWN_METHODS(PURE)
		MAPI_IATTACHMENTSECURITY_METHODS(PURE)
};

// http://msdn2.microsoft.com/en-us/library/bb820937.aspx
#define STORE_PUSHER_OK         ((ULONG) 0x00800000)

#define fnevIndexing		((ULONG) 0x00010000)

/* Indexing notifications (used for FTE related communications)		*/
/* Shares EXTENDED_NOTIFICATION to pass structures below,		*/
/* but NOTIFICATION type will be fnevIndexing				*/

// Stores that are pusher enabled (PR_STORE_SUPPORT_MASK contains STORE_PUSHER_OK)
// are required to send notifications regarding the process that is pushing.
#define INDEXING_SEARCH_OWNER		((ULONG) 0x00000001)

struct INDEX_SEARCH_PUSHER_PROCESS
{
	DWORD		dwPID;			// PID for process pushing
};

// http://blogs.msdn.com/stephen_griffin/archive/2006/05/11/595338.aspx
#define STORE_FULLTEXT_QUERY_OK ((ULONG) 0x02000000)
#define STORE_FILTER_SEARCH_OK  ((ULONG) 0x04000000)

// Will match prefix on words instead of the whole prop value
#define FL_PREFIX_ON_ANY_WORD	0x00000010

// Phrase match means the words have to be exactly matched and the
// sequence matters. This is different than FL_FULLSTRING because
// it doesn't require the whole property value to be the same. One
// term exactly matching a term in the property value is enough for
// a match even if there are more terms in the property.
#define FL_PHRASE_MATCH		0x00000020

// http://msdn2.microsoft.com/en-us/library/bb905283.aspx
#define dispidFormStorage		0x850F
#define dispidPageDirStream		0x8513
#define dispidFormPropStream	0x851B
#define dispidPropDefStream		0x8540
#define dispidScriptStream		0x8541
#define dispidCustomFlag		0x8542

#define INSP_ONEOFFFLAGS    0xD000000
#define INSP_PROPDEFINITION 0x2000000

// Sometimes IExchangeManageStore5 is in edkmdb.h, sometimes it isn't
#ifndef EXCHANGE_IEXCHANGEMANAGESTORE5_METHODS
#define USES_IID_IExchangeManageStore5

/*------------------------------------------------------------------------
*
*	'IExchangeManageStore5' Interface Declaration
*
*	Used for store management functions.
*
*-----------------------------------------------------------------------*/

#define EXCHANGE_IEXCHANGEMANAGESTORE5_METHODS(IPURE)					\
	MAPIMETHOD(GetMailboxTableEx)									\
	(THIS_	LPSTR						lpszServerName,				\
	LPGUID						lpguidMdb,					\
	LPMAPITABLE*			lppTable,					\
	ULONG						ulFlags,					\
	UINT						uOffset) IPURE;				\
	MAPIMETHOD(GetPublicFolderTableEx)								\
	(THIS_	LPSTR						lpszServerName,				\
	LPGUID						lpguidMdb,					\
	LPMAPITABLE*			lppTable,					\
	ULONG						ulFlags,					\
	UINT						uOffset) IPURE;				\

#undef  INTERFACE
#define INTERFACE  IExchangeManageStore5
DECLARE_MAPI_INTERFACE_(IExchangeManageStore5, IUnknown)
{
	MAPI_IUNKNOWN_METHODS(PURE)
		EXCHANGE_IEXCHANGEMANAGESTORE_METHODS(PURE)
		EXCHANGE_IEXCHANGEMANAGESTORE2_METHODS(PURE)
		EXCHANGE_IEXCHANGEMANAGESTORE3_METHODS(PURE)
		EXCHANGE_IEXCHANGEMANAGESTORE4_METHODS(PURE)
		EXCHANGE_IEXCHANGEMANAGESTORE5_METHODS(PURE)
};
#undef	IMPL
#define IMPL

DECLARE_MAPI_INTERFACE_PTR(IExchangeManageStore5, LPEXCHANGEMANAGESTORE5);
#endif // #ifndef EXCHANGE_IEXCHANGEMANAGESTORE5_METHODS

#define	CbNewROWLIST(_centries)	(offsetof(ROWLIST,aEntries) + \
	(_centries)*sizeof(ROWENTRY))
#define	MAXNewROWLIST (ULONG_MAX-offsetof(ROWLIST,aEntries))/sizeof(ROWENTRY)
#define MAXMessageClassArray (ULONG_MAX - offsetof(SMessageClassArray, aMessageClass))/sizeof(LPCSTR)
#define MAXNewADRLIST (ULONG_MAX - offsetof(ADRLIST, aEntries))/sizeof(ADRENTRY)

const WORD TZRULE_FLAG_RECUR_CURRENT_TZREG  = 0x0001; // see dispidApptTZDefRecur
const WORD TZRULE_FLAG_EFFECTIVE_TZREG      = 0x0002;

// http://blogs.msdn.com/stephen_griffin/archive/2007/03/19/mapi-and-exchange-2007.aspx
#define CONNECT_IGNORE_NO_PF					((ULONG)0x8000)

#define TABLE_SORT_CATEG_MAX ((ULONG) 0x00000004)
#define TABLE_SORT_CATEG_MIN ((ULONG) 0x00000008)