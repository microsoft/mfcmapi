#pragma once
// Guid definitions and helper functions
#include <core/mapi/interfaces.h>

namespace guid
{
	std::wstring GUIDToString(_In_opt_ LPCGUID lpGUID);
	std::wstring GUIDToString(_In_ GUID guid);
	std::wstring GUIDToStringAndName(_In_opt_ LPCGUID lpGUID);
	std::wstring GUIDToStringAndName(_In_ GUID guid);
	GUID GUIDNameToGUID(_In_ const std::wstring& szGUID, bool bByteSwapped);
	_Check_return_ GUID StringToGUID(_In_ const std::wstring& szGUID);
	_Check_return_ GUID StringToGUID(_In_ const std::wstring& szGUID, bool bByteSwapped);

	// clang-format off
	DEFINE_GUID(CLSID_MailMessage, 0x00020D0B, 0x0000, 0x0000, 0xC0, 0x00, 0x0, 0x00, 0x0, 0x00, 0x00, 0x46);
	DEFINE_OLEGUID(PS_INTERNET_HEADERS, 0x00020386, 0, 0);

	// http://msdn2.microsoft.com/en-us/library/bb905283.aspx
	DEFINE_OLEGUID(PSETID_Appointment, MAKELONG(0x2000 + 0x2, 0x0006), 0, 0);
	DEFINE_OLEGUID(PSETID_Task, MAKELONG(0x2000 + 0x3, 0x0006), 0, 0);
	DEFINE_OLEGUID(PSETID_Address, MAKELONG(0x2000 + 0x4, 0x0006), 0, 0);
	DEFINE_OLEGUID(PSETID_Common, MAKELONG(0x2000 + 0x8, 0x0006), 0, 0);
	DEFINE_OLEGUID(PSETID_Log, MAKELONG(0x2000 + 0xA, 0x0006), 0, 0);
	DEFINE_GUID(PSETID_Meeting, MAKELONG(0xDA90, 0x6ED8), 0x450B, 0x101B, 0x98, 0xDA, 0x0, 0xAA, 0x0, 0x3F, 0x13, 0x05);
	// [MS-OXPROPS].pdf
	DEFINE_OLEGUID(PSETID_Note, MAKELONG(0x2000 + 0xE, 0x0006), 0, 0);
	DEFINE_OLEGUID(PSETID_Report, MAKELONG(0x2000 + 0x13, 0x0006), 0, 0);
	DEFINE_OLEGUID(PSETID_Remote, MAKELONG(0x2000 + 0x14, 0x0006), 0, 0);
	DEFINE_OLEGUID(PSETID_Sharing, MAKELONG(0x2000 + 64, 0x0006), 0, 0);
	DEFINE_OLEGUID(PSETID_PostRss, MAKELONG(0x2000 + 65, 0x0006), 0, 0);
	DEFINE_GUID(PSETID_UnifiedMessaging, 0x4442858E, 0xA9E3, 0x4E80, 0xB9, 0x00, 0x31, 0x7A, 0x21, 0x0C, 0xC1, 0x5B);
	DEFINE_GUID(PSETID_AirSync, 0x71035549, 0x0739, 0x4DCB, 0x91, 0x63, 0x00, 0xF0, 0x58, 0x0D, 0xBB, 0xDF);
	DEFINE_GUID(PSETID_CalendarAssistant, 0x11000E07, 0xB51B, 0x40D6, 0xAF, 0x21, 0xCA, 0xA8, 0x5E, 0xDA, 0xB1, 0xD0);
	DEFINE_GUID(PSETID_Attachment, 0x96357F7F, 0x59E1, 0x47D0, 0x99, 0xA7, 0x46, 0x51, 0x5C, 0x18, 0x3B, 0x54);

	// http://support.microsoft.com/kb/312900
	// 53BC2EC0-D953-11CD-9752-00AA004AE40E
	DEFINE_GUID(GUID_Dilkie, 0x53BC2EC0, 0xD953, 0x11CD, 0x97, 0x52, 0x00, 0xAA, 0x00, 0x4A, 0xE4, 0x0E);

	// http://msdn2.microsoft.com/en-us/library/bb820937.aspx
#ifndef DEFINE_PRXGUID
#define DEFINE_PRXGUID(_name, _l) \
	DEFINE_GUID(_name, (0x29F3AB10 + _l), 0x554D, 0x11D0, 0xA9, 0x7C, 0x00, 0xA0, 0xC9, 0x11, 0xF5, 0x0A)
#endif // !DEFINE_PRXGUID
	DEFINE_PRXGUID(IID_IProxyStoreObject, 0x00000000L);

#ifdef _IID_IMAPIClientShutdown_MISSING_IN_HEADER
	// http://blogs.msdn.com/stephen_griffin/archive/2009/03/03/fastest-shutdown-in-the-west.aspx
#if !defined(INITGUID) || defined(USES_IID_IMAPIClientShutdown)
	DEFINE_OLEGUID(IID_IMAPIClientShutdown, 0x00020397, 0, 0);
#endif
#endif // _IID_IMAPIClientShutdown_MISSING_IN_HEADER

	// Sometimes IExchangeManageStore5 is in edkmdb.h, sometimes it isn't
#ifdef USES_IID_IExchangeManageStore5
	DEFINE_GUID(IID_IExchangeManageStore5, 0x7907dd18, 0xf141, 0x4676, 0xb1, 0x02, 0x37, 0xc9, 0xd9, 0x36, 0x34, 0x30);
#endif

	// http://msdn2.microsoft.com/en-us/library/bb820933.aspx
	DEFINE_GUID(IID_IAttachmentSecurity, 0xB2533636, 0xC3F3, 0x416f, 0xBF, 0x04, 0xAE, 0xFE, 0x41, 0xAB, 0xAA, 0xE2);

	// Class Identifiers
	// {4e3a7680-b77a-11d0-9da5-00c04fd65685}
	DEFINE_GUID(CLSID_IConverterSession, 0x4e3a7680, 0xb77a, 0x11d0, 0x9d, 0xa5, 0x0, 0xc0, 0x4f, 0xd6, 0x56, 0x85);

	// Interface Identifiers
	// {4b401570-b77b-11d0-9da5-00c04fd65685}
	DEFINE_GUID(IID_IConverterSession, 0x4b401570, 0xb77b, 0x11d0, 0x9d, 0xa5, 0x0, 0xc0, 0x4f, 0xd6, 0x56, 0x85);

	DEFINE_OLEGUID(PSUNKNOWN, MAKELONG(0x2000 + 999, 0x0006), 0, 0);

	// [MS-OXCDATA].pdf
	// Exchange Private Store Provider
	// {20FA551B-66AA-CD11-9BC8-00AA002FC45A}
	DEFINE_GUID(g_muidStorePrivate, 0x20FA551B, 0x66AA, 0xCD11, 0x9B, 0xC8, 0x00, 0xAA, 0x00, 0x2F, 0xC4, 0x5A);

	// Exchange Public Store Provider
	// {1002831C-66AA-CD11-9BC8-00AA002FC45A}
	DEFINE_GUID(g_muidStorePublic, 0x1002831C, 0x66AA, 0xCD11, 0x9B, 0xC8, 0x00, 0xAA, 0x00, 0x2F, 0xC4, 0x5A);

	// Contact Provider
	// {0AAA42FE-C718-101A-E885-0B651C240000}
	DEFINE_GUID(muidContabDLL, 0x0AAA42FE, 0xC718, 0x101A, 0xe8, 0x85, 0x0B, 0x65, 0x1C, 0x24, 0x00, 0x00);

	// Exchange Public Folder Store Provider
	// {9073441A-66AA-CD11-9BC8-00AA002FC45A}
	DEFINE_GUID(pbLongTermNonPrivateGuid, 0x9073441A, 0x66AA, 0xCD11, 0x9b, 0xc8, 0x00, 0xaa, 0x00, 0x2f, 0xc4, 0x5a);

	// One Off Entry Provider
	// {A41F2B81-A3BE-1910-9D6E-00DD010F5402}
	DEFINE_GUID(muidOOP, 0xA41F2B81, 0xA3BE, 0x1910, 0x9d, 0x6e, 0x00, 0xdd, 0x01, 0x0f, 0x54, 0x02);

	// MAPI Wrapped Message Store Provider
	// {10BBA138-E505-1A10-A1BB-08002B2A56C2}
	DEFINE_GUID(muidStoreWrap, 0x10BBA138, 0xE505, 0x1A10, 0xa1, 0xbb, 0x08, 0x00, 0x2b, 0x2a, 0x56, 0xc2);

	// Exchange Address Book Provider
	// {C840A7DC-42C0-1A10-B4B9-08002B2FE182}
	DEFINE_GUID(muidEMSAB, 0xC840A7DC, 0x42C0, 0x1A10, 0xB4, 0xB9, 0x08, 0x00, 0x2B, 0x2F, 0xE1, 0x82);

	// Contact Address Book Wrapped Entry ID
	// {D3AD91C0-9D51-11CF-A4A9-00AA0047FAA4}
	DEFINE_GUID(WAB_GUID, 0xD3AD91C0, 0x9D51, 0x11CF, 0xA4, 0xA9, 0x00, 0xAA, 0x00, 0x47, 0xFA, 0xA4);

	// http://blogs.msdn.com/b/stephen_griffin/archive/2010/09/13/you-chose-wisely.aspx
	// Capone profile section
	// {00020D0A-0000-0000-C000-000000000046}
	DEFINE_OLEGUID(IID_CAPONE_PROF, 0x00020d0a, 0, 0);

	// IMAPISync
	DEFINE_GUID(IID_IMAPISync, 0x5024a385, 0x2d44, 0x486a, 0x81, 0xa8, 0x8f, 0xe, 0xcb, 0x60, 0x71, 0xdd);

	// IMAPISyncProgressCallback
	DEFINE_GUID(IID_IMAPISyncProgressCallback, 0x5024a386, 0x2d44, 0x486a, 0x81, 0xa8, 0x8f, 0xe, 0xcb, 0x60, 0x71, 0xdd);

	// IMAPISecureMessage
	DEFINE_GUID(IID_IMAPISecureMessage, 0x253cc320, 0xeab6, 0x11d0, 0x82, 0x22, 0, 0x60, 0x97, 0x93, 0x87, 0xea);

	// IMAPIGetSession
	DEFINE_GUID(IID_IMAPIGetSession, 0x614ab435, 0x491d, 0x4f5b, 0xa8, 0xb4, 0x60, 0xeb, 0x3, 0x10, 0x30, 0xc6);

	// {23239608-685D-4732-9C55-4C95CB4E8E33}
	DEFINE_GUID(PSETID_XmlExtractedEntities, 0x23239608, 0x685D, 0x4732, 0x9C, 0x55, 0x4C, 0x95, 0xCB, 0x4E, 0x8E, 0x33);

	DEFINE_OLEGUID(CLSID_MailFolder, 0x00067800, 0, 0);
	DEFINE_OLEGUID(CLSID_ContactFolder, 0x00067801, 0, 0);
	DEFINE_OLEGUID(CLSID_CalendarFolder, 0x00067802, 0, 0);
	DEFINE_OLEGUID(CLSID_TaskFolder, 0x00067803, 0, 0);
	DEFINE_OLEGUID(CLSID_NoteFolder, 0x00067804, 0, 0);
	DEFINE_OLEGUID(CLSID_JournalFolder, 0x00067808, 0, 0);
	// clang-format on
} // namespace guid