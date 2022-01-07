#ifndef IMSCAPABILITIESGUID_H
#ifdef INITGUID
#include <mapiguid.h>
#define IMSCAPABILITIESGUID_H
#endif /* INITGUID */

// {00020393-0000-0000-C000-000000000046}
#if !defined(INITGUID) || defined(USES_IID_IMSCapabilities)
DEFINE_OLEGUID(IID_IMSCapabilities, 0x00020393, 0, 0);
#endif

#endif /* IMSCAPABILITIESGUID_H */

#ifndef IMSCAPABILITIES_H
#define IMSCAPABILITIES_H

#if _MSC_VER > 1000
#pragma once
#endif

#ifdef __cplusplus
extern "C"
{
#endif

#ifndef BEGIN_INTERFACE
#define BEGIN_INTERFACE
#endif

/* Forward interface declarations */

DECLARE_MAPI_INTERFACE_PTR(IMSCapabilities, LPMSCAPABILITIES);

// Constants

/* Return values for IMSCapabilities::GetCapabilities */
#define MSCAP_RES_ANNOTATION ((ULONG) 0x00000001)
#define MSCAP_SECURE_FOLDER_HOMEPAGES ((ULONG) 0x00000020)

enum class MSCAP_SELECTOR
{
	MSCAP_SEL_RESERVED1 = 0,
	MSCAP_SEL_RESERVED2,
	MSCAP_SEL_FOLDER,
	MSCAP_SEL_RESERVED3,
	MSCAP_SEL_RESTRICTION,
};

// Interface declarations

/*------------------------------------------------------------------------
 *
 *	"IMSCapabilities" Interface Declaration
 *
 *	Provides information about what a store can support.
 *
 *-----------------------------------------------------------------------*/

#define EXCHANGE_IMSCAPABILITIES_METHODS(IPURE) \
	MAPIMETHOD_(ULONG, GetCapabilities) \
	(THIS_ MSCAP_SELECTOR mscapSelector) IPURE;

#undef INTERFACE
#define INTERFACE IMSCapabilities
DECLARE_MAPI_INTERFACE_(IMSCapabilities, IUnknown){MAPI_IUNKNOWN_METHODS(PURE) EXCHANGE_IMSCAPABILITIES_METHODS(PURE)};

#ifdef __cplusplus
} /*	extern "C" */
#endif

#endif /* IMSCAPABILITIES_H */
