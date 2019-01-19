#pragma once

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
