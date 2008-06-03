// guids.h: where our guids are defined

#include "stdafx.h"

#ifndef MFCMAPIGUID_INCLUDED
#define MFCMAPIGUID_INCLUDED

DEFINE_GUID(CLSID_MailMessage,
			0x00020D0B,
			0x0000, 0x0000, 0xC0, 0x00, 0x0, 0x00, 0x0, 0x00, 0x00, 0x46);
DEFINE_OLEGUID(PS_INTERNET_HEADERS,	0x00020386, 0, 0);

// http://msdn2.microsoft.com/en-us/library/bb905283.aspx
DEFINE_OLEGUID(PSETID_Appointment,  MAKELONG(0x2000+(0x2),0x0006),0,0);
DEFINE_OLEGUID(PSETID_Task,         MAKELONG(0x2000+(0x3),0x0006),0,0);
DEFINE_OLEGUID(PSETID_Address,      MAKELONG(0x2000+(0x4),0x0006),0,0);
DEFINE_OLEGUID(PSETID_Common,       MAKELONG(0x2000+(0x8),0x0006),0,0);
DEFINE_OLEGUID(PSETID_Log,          MAKELONG(0x2000+(0xA),0x0006),0,0);
DEFINE_GUID(PSETID_Meeting,
			MAKELONG(0xDA90, 0x6ED8),
			0x450B, 0x101B,	0x98, 0xDA, 0x0, 0xAA, 0x0, 0x3F, 0x13, 0x05);
// [MS-OXPROPS].pdf
DEFINE_OLEGUID(PSETID_Note,          MAKELONG(0x2000+(0xE),0x0006),0,0);
DEFINE_OLEGUID(PSETID_Sharing,       MAKELONG(0x2000+(64),0x0006),0,0);
DEFINE_OLEGUID(PSETID_PostRss,       MAKELONG(0x2000+(65),0x0006),0,0);
DEFINE_GUID(PSETID_UnifiedMessaging,
			0x4442858E,
			0xA9E3,
			0x4E80,
			0xB9,0x00,0x31,0x7A,0x21,0x0C,0xC1,0x5B);
DEFINE_GUID(PSETID_AirSync, 0x71035549, 0x0739, 0x4DCB, 0x91, 0x63, 0x00, 0xF0, 0x58, 0x0D, 0xBB, 0xDF);

// http://support.microsoft.com/kb/312900
// 53BC2EC0-D953-11CD-9752-00AA004AE40E
DEFINE_GUID(GUID_Dilkie,
			0x53BC2EC0,
			0xD953, 0x11CD,	0x97, 0x52, 0x00, 0xAA, 0x00, 0x4A, 0xE4, 0x0E);

// http://msdn2.microsoft.com/en-us/library/bb820947.aspx
DEFINE_OLEGUID(IID_IMessageRaw, 0x0002038A, 0, 0);

// http://msdn2.microsoft.com/en-us/library/bb820937.aspx
#ifndef DEFINE_PRXGUID
#define DEFINE_PRXGUID(_name, _l)								\
DEFINE_GUID(_name, (0x29F3AB10 + _l), 0x554D, 0x11D0, 0xA9,		\
				0x7C, 0x00, 0xA0, 0xC9, 0x11, 0xF5, 0x0A)
#endif // !DEFINE_PRXGUID
DEFINE_PRXGUID(IID_IProxyStoreObject, 0x00000000L);

// Sometimes IExchangeManageStore5 is in edkmdb.h, sometimes it isn't
#ifdef USES_IID_IExchangeManageStore5
DEFINE_GUID(IID_IExchangeManageStore5,0x7907dd18, 0xf141, 0x4676, 0xb1, 0x02, 0x37, 0xc9, 0xd9, 0x36, 0x34, 0x30);
#endif

// http://msdn2.microsoft.com/en-us/library/bb820933.aspx
DEFINE_GUID(IID_IAttachmentSecurity,
			0xB2533636,
			0xC3F3, 0x416f, 0xBF, 0x04, 0xAE, 0xFE, 0x41, 0xAB, 0xAA, 0xE2);

//Class Identifiers
// {4e3a7680-b77a-11d0-9da5-00c04fd65685}
DEFINE_GUID(CLSID_IConverterSession, 0x4e3a7680, 0xb77a, 0x11d0, 0x9d, 0xa5, 0x0, 0xc0, 0x4f, 0xd6, 0x56, 0x85);

//Interface Identifiers
// {4b401570-b77b-11d0-9da5-00c04fd65685}
DEFINE_GUID(IID_IConverterSession, 0x4b401570, 0xb77b, 0x11d0, 0x9d, 0xa5, 0x0, 0xc0, 0x4f, 0xd6, 0x56, 0x85);

#endif //MFCMAPIGUID_INCLUDED
