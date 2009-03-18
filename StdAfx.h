// stdafx.h : include file for standard system include files,
//  or project specific include files that are used frequently, but
//	  are changed infrequently
//

#pragma once

#define VC_EXTRALEAN	// Exclude rarely-used stuff from Windows headers

#ifndef WINVER		// Permit use of features specific to Windows 95 and Windows NT 4.0 or later.
#define WINVER 0x0500	// Change this to the appropriate value to target
#endif                     // Windows 98 and Windows 2000 or later.

#ifndef _WIN32_WINNT	// Permit use of features specific to Windows NT 4.0 or later.
#define _WIN32_WINNT 0x0400	// Change this to the appropriate value to target
#endif		         // Windows 98 and Windows 2000 or later.

#ifndef _WIN32_IE		// Permit use of features specific to Internet Explorer 4.0 or later.
#define _WIN32_IE 0x0500   // Change this to the appropriate value to target
#endif			// Internet Explorer 5.0 or later.

#include <afxwin.h>		// MFC core and standard components
#include <afxext.h>		// MFC extensions
#include <afxcmn.h>		// MFC support for Windows Common Controls

// Safe String handling header
// #define STRSAFE_NO_CB_FUNCTIONS
#include <strsafe.h>

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

// Sample tree uses bldver.h to maintain some definitions. Do not remove.
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
	ULONG			ulRowID; // for recipients
	ULONG			ulRowType; // PR_ROW_TYPE
};

struct _PropListData
{
	ULONG		ulPropTag;
};

struct _MVPropData
{
	union _PV	val; // Allocated with MAPIAllocateMore
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
	ULONG				bSubfolders; // this is intentionally ULONG instead of BOOL
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
	BOOL			bItemFullyLoaded;
};

// macro to get allocated string lengths
#define CCH(string) (sizeof((string))/sizeof(TCHAR))
#define CCHW(string) (sizeof((string))/sizeof(WCHAR))

// Macros to assist in OnInitMenu
#define CHECK(state) ((state)?MF_CHECKED:MF_UNCHECKED)
#define DIM(state) ((state)?MF_ENABLED:MF_GRAYED)
#define DIMMSOK(iNumSelected) ((iNumSelected>=1)?MF_ENABLED:MF_GRAYED)
#define DIMMSNOK(iNumSelected) ((iNumSelected==1)?MF_ENABLED:MF_GRAYED)

// Flags for cached/offline mode - See http://msdn2.microsoft.com/en-us/library/bb820947.aspx
// Used in OpenMsgStore
#define MDB_ONLINE ((ULONG) 0x00000100)

// Used in OpenEntry
#define MAPI_NO_CACHE ((ULONG) 0x00000200)

/* Flag to keep calls from redirecting in cached mode */
#define MAPI_CACHE_ONLY         ((ULONG) 0x00004000)

#define MAPI_BG_SESSION         0x00200000 /* Used for async profile access */
#define SPAMFILTER_ONSAVE       ((ULONG) 0x00000080)

// Various flags gleaned from product documentation and KB articles
// http://msdn2.microsoft.com/en-us/library/ms526744.aspx
#define STORE_HTML_OK			((ULONG) 0x00010000)
#define STORE_ANSI_OK			((ULONG) 0x00020000)
#define STORE_LOCALSTORE		((ULONG) 0x00080000)

// http://msdn2.microsoft.com/en-us/library/bb820947.aspx
#define STORE_ITEMPROC			((ULONG) 0x00200000)
#define ITEMPROC_FORCE			((ULONG) 0x00000800)
#define NON_EMS_XP_SAVE			((ULONG) 0x00001000)

// http://msdn2.microsoft.com/en-us/library/ms531462.aspx
#define ATT_INVISIBLE_IN_HTML	((ULONG) 0x00000001)
#define ATT_INVISIBLE_IN_RTF	((ULONG) 0x00000002)
#define ATT_MHTML_REF			((ULONG) 0x00000004)

// http://msdn2.microsoft.com/en-us/library/ms527629.aspx
#define	MSGFLAG_ORIGIN_X400		((ULONG) 0x00001000)
#define	MSGFLAG_ORIGIN_INTERNET	((ULONG) 0x00002000)
#define MSGFLAG_ORIGIN_MISC_EXT ((ULONG) 0x00008000)

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

// http://support.microsoft.com/kb/278321
#define ENCODEDONTKNOW 0
#define ENCODEMIME 1
#define ENCODEUUENCODE 2
#define ENCODEBINHEX 3

// Flags used in PR_ROH_FLAGS - http://support.microsoft.com/kb/898835
// Connect to my Exchange mailbox using HTTP
#define ROHFLAGS_USE_ROH                0x1
// Connect using SSL only
#define ROHFLAGS_SSL_ONLY               0x2
// Mutually authenticate the session when connecting with SSL
#define ROHFLAGS_MUTUAL_AUTH            0x4
// On fast networks, connect using HTTP first, then connect using TCP/IP
#define ROHFLAGS_HTTP_FIRST_ON_FAST     0x8
// On slow networks, connect using HTTP first, then connect using TCP/IP
#define ROHFLAGS_HTTP_FIRST_ON_SLOW     0x20

// Flags used in PR_ROH_PROXY_AUTH_SCHEME
// Basic Authentication
#define ROHAUTH_BASIC                   0x1
// NTLM Authentication
#define ROHAUTH_NTLM                    0x2

// http://support.microsoft.com/kb/194955
#define AG_MONTHS  0
#define AG_WEEKS   1
#define AG_DAYS    2

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
#define WM_MFCMAPI_CLEARABORT				WM_APP+8
#define WM_MFCMAPI_GETABORT					WM_APP+9

// Used by CSingleMAPIPropListCtrl and CSortHeader
#define WM_MFCMAPI_SAVECOLUMNORDERHEADER	WM_APP+10
#define WM_MFCMAPI_SAVECOLUMNORDERLIST		WM_APP+11

// Definitions for WrapCompressedRTFStreamEx in param for WrapCompressedRTFStreamEX
// http://msdn2.microsoft.com/en-us/library/bb905293.aspx
typedef struct {
	ULONG		size;		// Size of the structure
	ULONG		ulFlags;
		/****** MAPI_MODIFY 				((ULONG) 0x00000001) above */
		/****** STORE_UNCOMPRESSED_RTF		((ULONG) 0x00008000) above */
		/****** MAPI_NATIVE_BODY			((ULONG) 0x00010000) mapidefs.h   Only used for reading*/
	ULONG		ulInCodePage;	// Codepage of the message, used when passing MAPI_NATIVE_BODY, ignored otherwise
	ULONG		ulOutCodePage;	// Codepage of the Returned Stream, used when passing MAPI_NATIVE_BODY, ignored otherwise
} RTF_WCSINFO;

// out param type information for WrapCompressedRTFStreamEX
// http://msdn2.microsoft.com/en-us/library/bb905294.aspx
typedef struct {
	ULONG       size;	// Size of the structure
	ULONG		ulStreamFlags;
		/****** MAPI_NATIVE_BODY_TYPE_RTF			((ULONG) 0x00000001) mapidefs.h */
		/****** MAPI_NATIVE_BODY_TYPE_HTML			((ULONG) 0x00000002) mapidefs.h */
		/****** MAPI_NATIVE_BODY_TYPE_PLAINTEXT 	((ULONG) 0x00000004) mapidefs.h */
} RTF_WCSRETINFO;

// http://msdn2.microsoft.com/en-us/library/bb905291.aspx
typedef HRESULT (STDMETHODCALLTYPE WRAPCOMPRESSEDRTFSTREAMEX) (
						   IStream * pCompressedRTFStream, CONST RTF_WCSINFO * pWCSInfo,
							IStream ** ppUncompressedRTFStream, RTF_WCSRETINFO * pRetInfo);

typedef WRAPCOMPRESSEDRTFSTREAMEX *LPWRAPCOMPRESSEDRTFSTREAMEX;

#define MAPI_NATIVE_BODY					0x00010000

/* out param type infomation for WrapCompressedRTFStreamEx */
#define MAPI_NATIVE_BODY_TYPE_RTF  			0x00000001
#define MAPI_NATIVE_BODY_TYPE_HTML 			0x00000002
#define MAPI_NATIVE_BODY_TYPE_PLAINTEXT		0x00000004

// For GetComponentPath
typedef BOOL (STDAPICALLTYPE FGETCOMPONENTPATH)
	(LPSTR szComponent,
	LPSTR szQualifier,
	LPSTR szDllPath,
	DWORD cchBufferSize,
	BOOL fInstall);
typedef FGETCOMPONENTPATH FAR * LPFGETCOMPONENTPATH;

// For EditSecurity
typedef BOOL (STDAPICALLTYPE EDITSECURITY)
(
	HWND hwndOwner,
	LPSECURITYINFO psi
);
typedef EDITSECURITY FAR * LPEDITSECURITY;

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
typedef STGCREATESTORAGEEX FAR * LPSTGCREATESTORAGEEX;

// For Themes
typedef HTHEME (STDMETHODCALLTYPE OPENTHEMEDATA)
(
 HWND hwnd,
 LPCWSTR pszClassList);
typedef OPENTHEMEDATA FAR * LPOPENTHEMEDATA;

typedef HTHEME (STDMETHODCALLTYPE CLOSETHEMEDATA)
(
 HTHEME hTheme);
typedef CLOSETHEMEDATA FAR * LPCLOSETHEMEDATA;

typedef HRESULT (STDMETHODCALLTYPE GETTHEMEMARGINS)
(
 HTHEME hTheme,
 OPTIONAL HDC hdc,
 int iPartId,
 int iStateId,
 int iPropId,
 OPTIONAL RECT *prc,
 OUT MARGINS *pMargins);
typedef GETTHEMEMARGINS FAR * LPGETTHEMEMARGINS;

typedef HRESULT (STDMETHODCALLTYPE MIMEOLEGETCODEPAGECHARSET)
(
 CODEPAGEID cpiCodePage,
 CHARSETTYPE ctCsetType,
 LPHCHARSET phCharset
);
typedef MIMEOLEGETCODEPAGECHARSET FAR * LPMIMEOLEGETCODEPAGECHARSET;

// For all the MAPI funcs we use
typedef SCODE (STDMETHODCALLTYPE HRGETONEPROP)(
					LPMAPIPROP lpMapiProp,
					ULONG ulPropTag,
					LPSPropValue FAR *lppProp);
typedef HRGETONEPROP *LPHRGETONEPROP;

typedef void (STDMETHODCALLTYPE FREEPROWS)(
					LPSRowSet lpRows);
typedef FREEPROWS	*LPFREEPROWS;

typedef	LPSPropValue (STDMETHODCALLTYPE PPROPFINDPROP)(
					LPSPropValue lpPropArray, ULONG cValues,
					ULONG ulPropTag);
typedef PPROPFINDPROP *LPPPROPFINDPROP;

typedef SCODE (STDMETHODCALLTYPE SCDUPPROPSET)(
					int cValues, LPSPropValue lpPropArray,
					LPALLOCATEBUFFER lpAllocateBuffer, LPSPropValue FAR *lppPropArray);
typedef SCDUPPROPSET *LPSCDUPPROPSET;

typedef SCODE (STDMETHODCALLTYPE SCCOUNTPROPS)(
					int cValues, LPSPropValue lpPropArray, ULONG FAR *lpcb);
typedef SCCOUNTPROPS *LPSCCOUNTPROPS;

typedef SCODE (STDMETHODCALLTYPE SCCOPYPROPS)(
					int cValues, LPSPropValue lpPropArray, LPVOID lpvDst,
					ULONG FAR *lpcb);
typedef SCCOPYPROPS *LPSCCOPYPROPS;

typedef SCODE (STDMETHODCALLTYPE OPENIMSGONISTG)(
					LPMSGSESS		lpMsgSess,
					LPALLOCATEBUFFER lpAllocateBuffer,
					LPALLOCATEMORE 	lpAllocateMore,
					LPFREEBUFFER	lpFreeBuffer,
					LPMALLOC		lpMalloc,
					LPVOID			lpMapiSup,
					LPSTORAGE 		lpStg,
					MSGCALLRELEASE FAR *lpfMsgCallRelease,
					ULONG			ulCallerData,
					ULONG			ulFlags,
					LPMESSAGE		FAR *lppMsg );
typedef OPENIMSGONISTG *LPOPENIMSGONISTG;

typedef	LPMALLOC (STDMETHODCALLTYPE MAPIGETDEFAULTMALLOC)(
					void);
typedef MAPIGETDEFAULTMALLOC *LPMAPIGETDEFAULTMALLOC;

typedef	void (STDMETHODCALLTYPE CLOSEIMSGSESSION)(
					LPMSGSESS		lpMsgSess);
typedef CLOSEIMSGSESSION *LPCLOSEIMSGSESSION;

typedef	SCODE (STDMETHODCALLTYPE OPENIMSGSESSION)(
					LPMALLOC		lpMalloc,
					ULONG			ulFlags,
					LPMSGSESS FAR	*lppMsgSess);
typedef OPENIMSGSESSION *LPOPENIMSGSESSION;

typedef SCODE (STDMETHODCALLTYPE HRQUERYALLROWS)(
					LPMAPITABLE lpTable,
					LPSPropTagArray lpPropTags,
					LPSRestriction lpRestriction,
					LPSSortOrderSet lpSortOrderSet,
					LONG crowsMax,
					LPSRowSet FAR *lppRows);
typedef HRQUERYALLROWS *LPHRQUERYALLROWS;

typedef SCODE (STDMETHODCALLTYPE MAPIOPENFORMMGR)(
					LPMAPISESSION pSession, LPMAPIFORMMGR FAR * ppmgr);
typedef MAPIOPENFORMMGR *LPMAPIOPENFORMMGR;

typedef HRESULT (STDMETHODCALLTYPE RTFSYNC) (
	LPMESSAGE lpMessage, ULONG ulFlags, BOOL FAR * lpfMessageUpdated);
typedef RTFSYNC *LPRTFSYNC;

typedef SCODE (STDMETHODCALLTYPE HRSETONEPROP)(
					LPMAPIPROP lpMapiProp,
					LPSPropValue lpProp);
typedef HRSETONEPROP *LPHRSETONEPROP;

typedef void (STDMETHODCALLTYPE FREEPADRLIST)(
					LPADRLIST lpAdrList);
typedef FREEPADRLIST *LPFREEPADRLIST;

typedef SCODE (STDMETHODCALLTYPE PROPCOPYMORE)(
					LPSPropValue		lpSPropValueDest,
					LPSPropValue		lpSPropValueSrc,
					ALLOCATEMORE *	lpfAllocMore,
					LPVOID			lpvObject );
typedef PROPCOPYMORE *LPPROPCOPYMORE;

typedef HRESULT (STDMETHODCALLTYPE WRAPCOMPRESSEDRTFSTREAM) (
	LPSTREAM lpCompressedRTFStream, ULONG ulFlags, LPSTREAM FAR * lpUncompressedRTFStream);
typedef WRAPCOMPRESSEDRTFSTREAM *LPWRAPCOMPRESSEDRTFSTREAM;

typedef SCODE (STDMETHODCALLTYPE HRVALIDATEIPMSUBTREE)(
					LPMDB lpMDB, ULONG ulFlags,
					ULONG FAR *lpcValues, LPSPropValue FAR *lppValues,
					LPMAPIERROR FAR *lpperr);
typedef HRVALIDATEIPMSUBTREE *LPHRVALIDATEIPMSUBTREE;

typedef HRESULT (STDAPICALLTYPE MAPIOPENLOCALFORMCONTAINER)(LPMAPIFORMCONTAINER FAR * ppfcnt);
typedef MAPIOPENLOCALFORMCONTAINER *LPMAPIOPENLOCALFORMCONTAINER;

typedef SCODE (STDAPICALLTYPE HRDISPATCHNOTIFICATIONS)(ULONG ulFlags);
typedef HRDISPATCHNOTIFICATIONS *LPHRDISPATCHNOTIFICATIONS;

// http://msdn2.microsoft.com/en-us/library/bb820924.aspx
#pragma pack(4)
typedef struct _contab_entryid
{
BYTE misc1[4];
MAPIUID misc2;
ULONG misc3;
ULONG misc4;
ULONG misc5;
// EntryID of contact in store.
ULONG cbeid;
BYTE abeid[1];
} CONTAB_ENTRYID, *LPCONTAB_ENTRYID;
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

// http://blogs.msdn.com/stephen_griffin
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
		(LPCWSTR pwszFileName, BOOL *pfBlocked) IPURE;

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

typedef struct _INDEX_SEARCH_PUSHER_PROCESS
{
	DWORD		dwPID;			// PID for process pushing
} INDEX_SEARCH_PUSHER_PROCESS;

// http://blogs.msdn.com/stephen_griffin/archive/2006/05/11/595338.aspx
#define STORE_FULLTEXT_QUERY_OK ((ULONG) 0x02000000)
#define STORE_FILTER_SEARCH_OK  ((ULONG) 0x04000000)

// Will match prefix on words instead of the whole prop value
#define FL_PREFIX_ON_ANY_WORD 	0x00000010

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
#define dispidCustomFlag 		0x8542

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
				LPMAPITABLE FAR *			lppTable,					\
				ULONG						ulFlags,					\
				UINT						uOffset) IPURE;				\
	MAPIMETHOD(GetPublicFolderTableEx)								\
		(THIS_	LPSTR						lpszServerName,				\
				LPGUID						lpguidMdb,					\
				LPMAPITABLE FAR *			lppTable,					\
				ULONG						ulFlags,					\
				UINT						uOffset) IPURE;				\

#undef		 INTERFACE
#define		 INTERFACE  IExchangeManageStore5
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

// http://support.microsoft.com/kb/884671
#define STORE_UNICODE_OK		((ULONG) 0x00040000)

const WORD TZRULE_FLAG_RECUR_CURRENT_TZREG  = 0x0001; // see dispidApptTZDefRecur
const WORD TZRULE_FLAG_EFFECTIVE_TZREG      = 0x0002;

const ULONG	TZ_MAX_RULES		= 1024;
const BYTE	TZ_BIN_VERSION_MAJOR	= 0x02;
const BYTE	TZ_BIN_VERSION_MINOR	= 0x01;

// http://blogs.msdn.com/stephen_griffin/archive/2007/03/19/mapi-and-exchange-2007.aspx
#define CONNECT_IGNORE_NO_PF					((ULONG)0x8000)

// http://blogs.msdn.com/stephen_griffin/archive/2008/04/01/new-restriction-types-seen-in-wrapped-psts.aspx
#define RES_COUNT			((ULONG) 0x0000000B)
#define RES_ANNOTATION		((ULONG) 0x0000000C)