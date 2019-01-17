#pragma once

#define VC_EXTRALEAN // Exclude rarely-used stuff from Windows headers

#include <Windows.h>

// Common CRT headers
#include <list>
#include <algorithm>
#include <locale>
#include <sstream>
#include <iterator>
#include <functional>

// All the MAPI headers
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

#include <core/res/Resource.h> // main symbols

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
