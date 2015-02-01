#include "stdafx.h"
#include "..\stdafx.h"
#include "EntryIdStruct.h"
#include "SmartView.h"
#include "..\String.h"
#include "..\Guids.h"
#include "..\Interpretprop2.h"
#include "..\ParseProperty.h"
#include "..\ExtraPropTags.h"

EntryIdStruct::EntryIdStruct(ULONG cbBin, _In_count_(cbBin) LPBYTE lpBin) : SmartViewParser(cbBin, lpBin)
{
	memset(m_abFlags, 0, sizeof(m_abFlags));
	memset(m_ProviderUID, 0, sizeof(m_ProviderUID));
	m_ObjectType = eidtUnknown;
	m_ProviderData = { 0 };
}

EntryIdStruct::~EntryIdStruct()
{
	switch (m_ObjectType)
	{
	case eidtOneOff:
		if (MAPI_UNICODE & m_ProviderData.OneOffRecipientObject.Bitmask)
		{
			delete[] m_ProviderData.OneOffRecipientObject.Strings.Unicode.DisplayName;
			delete[] m_ProviderData.OneOffRecipientObject.Strings.Unicode.AddressType;
			delete[] m_ProviderData.OneOffRecipientObject.Strings.Unicode.EmailAddress;
		}
		else
		{
			delete[] m_ProviderData.OneOffRecipientObject.Strings.ANSI.DisplayName;
			delete[] m_ProviderData.OneOffRecipientObject.Strings.ANSI.AddressType;
			delete[] m_ProviderData.OneOffRecipientObject.Strings.ANSI.EmailAddress;
		}
		break;
	case eidtAddressBook:
		delete[] m_ProviderData.AddressBookObject.X500DN;
		break;
	case eidtContact:
		delete m_ProviderData.ContactAddressBookObject.lpEntryID;
		break;
	case eidtMessageDatabase:
		delete[] m_ProviderData.MessageDatabaseObject.DLLFileName;
		delete[] m_ProviderData.MessageDatabaseObject.ServerShortname;
		delete[] m_ProviderData.MessageDatabaseObject.MailboxDN;
		delete[] m_ProviderData.MessageDatabaseObject.v2DN;
		delete[] m_ProviderData.MessageDatabaseObject.v2FQDN;
		break;
	case eidtWAB:
		delete m_ProviderData.WAB.lpEntryID;
	}
}

void EntryIdStruct::Parse()
{
	m_Parser.GetBYTE(&m_abFlags[0]);
	m_Parser.GetBYTE(&m_abFlags[1]);
	m_Parser.GetBYTE(&m_abFlags[2]);
	m_Parser.GetBYTE(&m_abFlags[3]);
	m_Parser.GetBYTESNoAlloc(sizeof(m_ProviderUID), sizeof(m_ProviderUID), (LPBYTE)&m_ProviderUID);

	// One Off Recipient
	if (!memcmp(m_ProviderUID, &muidOOP, sizeof(GUID)))
	{
		m_ObjectType = eidtOneOff;
	}
	// Address Book Recipient
	else if (!memcmp(m_ProviderUID, &muidEMSAB, sizeof(GUID)))
	{
		m_ObjectType = eidtAddressBook;
	}
	// Contact Address Book / Personal Distribution List (PDL)
	else if (!memcmp(m_ProviderUID, &muidContabDLL, sizeof(GUID)))
	{
		m_ObjectType = eidtContact;
	}
	// message store objects
	else if (!memcmp(m_ProviderUID, &muidStoreWrap, sizeof(GUID)))
	{
		m_ObjectType = eidtMessageDatabase;
	}
	else if (!memcmp(m_ProviderUID, &WAB_GUID, sizeof(GUID)))
	{
		m_ObjectType = eidtWAB;
	}
	// We can recognize Exchange message store folder and message entry IDs by their size
	else
	{
		size_t ulRemainingBytes = m_Parser.RemainingBytes();

		if (sizeof(WORD) + sizeof(GUID) + 6 * sizeof(BYTE) + 2 * sizeof(BYTE) == ulRemainingBytes)
		{
			m_ObjectType = eidtFolder;
		}
		else if (sizeof(WORD) + 2 * sizeof(GUID) + 12 * sizeof(BYTE) + 4 * sizeof(BYTE) == ulRemainingBytes)
		{
			m_ObjectType = eidtMessage;
		}
	}

	if (eidtUnknown != m_ObjectType)
	{
		switch (m_ObjectType)
		{
			// One Off Recipient
		case eidtOneOff:
			m_Parser.GetDWORD(&m_ProviderData.OneOffRecipientObject.Bitmask);
			if (MAPI_UNICODE & m_ProviderData.OneOffRecipientObject.Bitmask)
			{
				m_Parser.GetStringW(&m_ProviderData.OneOffRecipientObject.Strings.Unicode.DisplayName);
				m_Parser.GetStringW(&m_ProviderData.OneOffRecipientObject.Strings.Unicode.AddressType);
				m_Parser.GetStringW(&m_ProviderData.OneOffRecipientObject.Strings.Unicode.EmailAddress);
			}
			else
			{
				m_Parser.GetStringA(&m_ProviderData.OneOffRecipientObject.Strings.ANSI.DisplayName);
				m_Parser.GetStringA(&m_ProviderData.OneOffRecipientObject.Strings.ANSI.AddressType);
				m_Parser.GetStringA(&m_ProviderData.OneOffRecipientObject.Strings.ANSI.EmailAddress);
			}
			break;
			// Address Book Recipient
		case eidtAddressBook:
			m_Parser.GetDWORD(&m_ProviderData.AddressBookObject.Version);
			m_Parser.GetDWORD(&m_ProviderData.AddressBookObject.Type);
			m_Parser.GetStringA(&m_ProviderData.AddressBookObject.X500DN);
			break;
			// Contact Address Book / Personal Distribution List (PDL)
		case eidtContact:
		{
			m_Parser.GetDWORD(&m_ProviderData.ContactAddressBookObject.Version);
			m_Parser.GetDWORD(&m_ProviderData.ContactAddressBookObject.Type);

			if (CONTAB_CONTAINER == m_ProviderData.ContactAddressBookObject.Type)
			{
				m_Parser.GetBYTESNoAlloc(
					sizeof(m_ProviderData.ContactAddressBookObject.muidID),
					sizeof(m_ProviderData.ContactAddressBookObject.muidID),
					(LPBYTE)&m_ProviderData.ContactAddressBookObject.muidID);
			}
			else // Assume we've got some variation on the user/distlist format
			{
				m_Parser.GetDWORD(&m_ProviderData.ContactAddressBookObject.Index);
				m_Parser.GetDWORD(&m_ProviderData.ContactAddressBookObject.EntryIDCount);
			}

			// Read the wrapped entry ID from the remaining data
			size_t cbRemainingBytes = m_Parser.RemainingBytes();

			// If we already got a size, use it, else we just read the rest of the structure
			if (0 != m_ProviderData.ContactAddressBookObject.EntryIDCount &&
				m_ProviderData.ContactAddressBookObject.EntryIDCount < cbRemainingBytes)
			{
				cbRemainingBytes = m_ProviderData.ContactAddressBookObject.EntryIDCount;
			}

			m_ProviderData.ContactAddressBookObject.lpEntryID = new EntryIdStruct(
				(ULONG)cbRemainingBytes,
				m_Parser.GetCurrentAddress());
			if (m_ProviderData.ContactAddressBookObject.lpEntryID)
			{
				m_ProviderData.ContactAddressBookObject.lpEntryID->DisableJunkParsing();
				m_ProviderData.ContactAddressBookObject.lpEntryID->EnsureParsed();
				m_Parser.Advance(m_ProviderData.ContactAddressBookObject.lpEntryID->GetCurrentOffset());
			}
		}
		break;
		case eidtWAB:
		{
			m_ObjectType = eidtWAB;

			m_Parser.GetBYTE(&m_ProviderData.WAB.Type);

			m_ProviderData.WAB.lpEntryID = new EntryIdStruct(
				(ULONG)m_Parser.RemainingBytes(),
				m_Parser.GetCurrentAddress());
			if (m_ProviderData.WAB.lpEntryID)
			{
				m_ProviderData.WAB.lpEntryID->DisableJunkParsing();
				m_ProviderData.WAB.lpEntryID->EnsureParsed();
				m_Parser.Advance(m_ProviderData.WAB.lpEntryID->GetCurrentOffset());
			}
		}
		break;
		// message store objects
		case eidtMessageDatabase:
			m_Parser.GetBYTE(&m_ProviderData.MessageDatabaseObject.Version);
			m_Parser.GetBYTE(&m_ProviderData.MessageDatabaseObject.Flag);
			m_Parser.GetStringA(&m_ProviderData.MessageDatabaseObject.DLLFileName);
			m_ProviderData.MessageDatabaseObject.bIsExchange = false;

			// We only know how to parse emsmdb.dll's wrapped entry IDs
			if (m_ProviderData.MessageDatabaseObject.DLLFileName &&
				CSTR_EQUAL == CompareStringA(LOCALE_INVARIANT,
				NORM_IGNORECASE,
				m_ProviderData.MessageDatabaseObject.DLLFileName,
				-1,
				"emsmdb.dll", // STRING_OK
				-1))
			{
				m_ProviderData.MessageDatabaseObject.bIsExchange = true;
				size_t cbRead = m_Parser.GetCurrentOffset();
				// Advance to the next multiple of 4
				m_Parser.Advance(3 - ((cbRead + 3) % 4));
				m_Parser.GetDWORD(&m_ProviderData.MessageDatabaseObject.WrappedFlags);
				m_Parser.GetBYTESNoAlloc(sizeof(m_ProviderData.MessageDatabaseObject.WrappedProviderUID),
					sizeof(m_ProviderData.MessageDatabaseObject.WrappedProviderUID),
					(LPBYTE)&m_ProviderData.MessageDatabaseObject.WrappedProviderUID);
				m_Parser.GetDWORD(&m_ProviderData.MessageDatabaseObject.WrappedType);
				m_Parser.GetStringA(&m_ProviderData.MessageDatabaseObject.ServerShortname);

				m_ProviderData.MessageDatabaseObject.bV2 = false;

				// Test if we have a magic value. Some PF EIDs also have a mailbox DN and we need to accomodate them
				if (m_ProviderData.MessageDatabaseObject.WrappedType & OPENSTORE_PUBLIC)
				{
					cbRead = m_Parser.GetCurrentOffset();
					m_Parser.GetDWORD(&m_ProviderData.MessageDatabaseObject.v2.ulMagic);
					if (MDB_STORE_EID_V2_MAGIC == m_ProviderData.MessageDatabaseObject.v2.ulMagic)
					{
						m_ProviderData.MessageDatabaseObject.bV2 = true;
					}

					m_Parser.SetCurrentOffset(cbRead);
				}

				// Either we're not a PF eid, or this PF EID wasn't followed directly by a magic value
				if (!(m_ProviderData.MessageDatabaseObject.WrappedType & OPENSTORE_PUBLIC) ||
					!m_ProviderData.MessageDatabaseObject.bV2)
				{
					m_Parser.GetStringA(&m_ProviderData.MessageDatabaseObject.MailboxDN);
				}

				if (m_Parser.RemainingBytes() >= sizeof(MDB_STORE_EID_V2) + sizeof(WCHAR))
				{
					m_ProviderData.MessageDatabaseObject.bV2 = true;
					m_Parser.GetDWORD(&m_ProviderData.MessageDatabaseObject.v2.ulMagic);
					m_Parser.GetDWORD(&m_ProviderData.MessageDatabaseObject.v2.ulSize);
					m_Parser.GetDWORD(&m_ProviderData.MessageDatabaseObject.v2.ulVersion);
					m_Parser.GetDWORD(&m_ProviderData.MessageDatabaseObject.v2.ulOffsetDN);
					m_Parser.GetDWORD(&m_ProviderData.MessageDatabaseObject.v2.ulOffsetFQDN);
					if (m_ProviderData.MessageDatabaseObject.v2.ulOffsetDN)
					{
						m_Parser.GetStringA(&m_ProviderData.MessageDatabaseObject.v2DN);
					}

					if (m_ProviderData.MessageDatabaseObject.v2.ulOffsetFQDN)
					{
						m_Parser.GetStringW(&m_ProviderData.MessageDatabaseObject.v2FQDN);
					}

					m_Parser.GetBYTESNoAlloc(sizeof(m_ProviderData.MessageDatabaseObject.v2Reserved),
						sizeof(m_ProviderData.MessageDatabaseObject.v2Reserved),
						(LPBYTE)&m_ProviderData.MessageDatabaseObject.v2Reserved);
				}
			}
			break;
			// Exchange message store folder
		case eidtFolder:
			m_Parser.GetWORD(&m_ProviderData.FolderOrMessage.Type);
			m_Parser.GetBYTESNoAlloc(sizeof(m_ProviderData.FolderOrMessage.Data.FolderObject.DatabaseGUID),
				sizeof(m_ProviderData.FolderOrMessage.Data.FolderObject.DatabaseGUID),
				(LPBYTE)&m_ProviderData.FolderOrMessage.Data.FolderObject.DatabaseGUID);
			m_Parser.GetBYTESNoAlloc(sizeof(m_ProviderData.FolderOrMessage.Data.FolderObject.GlobalCounter),
				sizeof(m_ProviderData.FolderOrMessage.Data.FolderObject.GlobalCounter),
				(LPBYTE)&m_ProviderData.FolderOrMessage.Data.FolderObject.GlobalCounter);
			m_Parser.GetBYTESNoAlloc(sizeof(m_ProviderData.FolderOrMessage.Data.FolderObject.Pad),
				sizeof(m_ProviderData.FolderOrMessage.Data.FolderObject.Pad),
				(LPBYTE)&m_ProviderData.FolderOrMessage.Data.FolderObject.Pad);
			break;
			// Exchange message store message
		case eidtMessage:
			m_Parser.GetWORD(&m_ProviderData.FolderOrMessage.Type);
			m_Parser.GetBYTESNoAlloc(sizeof(m_ProviderData.FolderOrMessage.Data.MessageObject.FolderDatabaseGUID),
				sizeof(m_ProviderData.FolderOrMessage.Data.MessageObject.FolderDatabaseGUID),
				(LPBYTE)&m_ProviderData.FolderOrMessage.Data.MessageObject.FolderDatabaseGUID);
			m_Parser.GetBYTESNoAlloc(sizeof(m_ProviderData.FolderOrMessage.Data.MessageObject.FolderGlobalCounter),
				sizeof(m_ProviderData.FolderOrMessage.Data.MessageObject.FolderGlobalCounter),
				(LPBYTE)&m_ProviderData.FolderOrMessage.Data.MessageObject.FolderGlobalCounter);
			m_Parser.GetBYTESNoAlloc(sizeof(m_ProviderData.FolderOrMessage.Data.MessageObject.Pad1),
				sizeof(m_ProviderData.FolderOrMessage.Data.MessageObject.Pad1),
				(LPBYTE)&m_ProviderData.FolderOrMessage.Data.MessageObject.Pad1);
			m_Parser.GetBYTESNoAlloc(sizeof(m_ProviderData.FolderOrMessage.Data.MessageObject.MessageDatabaseGUID),
				sizeof(m_ProviderData.FolderOrMessage.Data.MessageObject.MessageDatabaseGUID),
				(LPBYTE)&m_ProviderData.FolderOrMessage.Data.MessageObject.MessageDatabaseGUID);
			m_Parser.GetBYTESNoAlloc(sizeof(m_ProviderData.FolderOrMessage.Data.MessageObject.MessageGlobalCounter),
				sizeof(m_ProviderData.FolderOrMessage.Data.MessageObject.MessageGlobalCounter),
				(LPBYTE)&m_ProviderData.FolderOrMessage.Data.MessageObject.MessageGlobalCounter);
			m_Parser.GetBYTESNoAlloc(sizeof(m_ProviderData.FolderOrMessage.Data.MessageObject.Pad2),
				sizeof(m_ProviderData.FolderOrMessage.Data.MessageObject.Pad2),
				(LPBYTE)&m_ProviderData.FolderOrMessage.Data.MessageObject.Pad2);
			break;
		}
	}

	// Check if we have an unidentified short term entry ID:
	if (eidtUnknown == m_ObjectType && m_abFlags[0] & MAPI_SHORTTERM)
		m_ObjectType = eidtShortTerm;
}

_Check_return_ wstring EntryIdStruct::ToStringInternal()
{
	wstring szEntryId;

	switch (m_ObjectType)
	{
	case eidtUnknown:
		szEntryId = formatmessage(IDS_ENTRYIDUNKNOWN);
		break;
	case eidtShortTerm:
		szEntryId = formatmessage(IDS_ENTRYIDSHORTTERM);
		break;
	case eidtFolder:
		szEntryId = formatmessage(IDS_ENTRYIDEXCHANGEFOLDER);
		break;
	case eidtMessage:
		szEntryId = formatmessage(IDS_ENTRYIDEXCHANGEMESSAGE);
		break;
	case eidtMessageDatabase:
		szEntryId = formatmessage(IDS_ENTRYIDMAPIMESSAGESTORE);
		break;
	case eidtOneOff:
		szEntryId = formatmessage(IDS_ENTRYIDMAPIONEOFF);
		break;
	case eidtAddressBook:
		szEntryId = formatmessage(IDS_ENTRYIDEXCHANGEADDRESS);
		break;
	case eidtContact:
		szEntryId = formatmessage(IDS_ENTRYIDCONTACTADDRESS);
		break;
	case eidtWAB:
		szEntryId = formatmessage(IDS_ENTRYIDWRAPPEDENTRYID);
		break;
	}

	if (0 == (m_abFlags[0] | m_abFlags[1] | m_abFlags[2] | m_abFlags[3]))
	{
		szEntryId += formatmessage(IDS_ENTRYIDHEADER1);
	}
	else if (0 == (m_abFlags[1] | m_abFlags[2] | m_abFlags[3]))
	{
		szEntryId += formatmessage(IDS_ENTRYIDHEADER2,
			m_abFlags[0], InterpretFlags(flagEntryId0, m_abFlags[0]).c_str());
	}
	else
	{
		szEntryId += formatmessage(IDS_ENTRYIDHEADER3,
			m_abFlags[0], InterpretFlags(flagEntryId0, m_abFlags[0]).c_str(),
			m_abFlags[1], InterpretFlags(flagEntryId1, m_abFlags[1]).c_str(),
			m_abFlags[2],
			m_abFlags[3]);
	}

	wstring szGUID = GUIDToStringAndName((LPGUID)&m_ProviderUID);
	szEntryId += formatmessage(IDS_ENTRYIDPROVIDERGUID, szGUID.c_str());

	if (eidtOneOff == m_ObjectType)
	{
		wstring szFlags = InterpretFlags(flagOneOffEntryId, m_ProviderData.OneOffRecipientObject.Bitmask);
		if (MAPI_UNICODE & m_ProviderData.OneOffRecipientObject.Bitmask)
		{
			szEntryId += formatmessage(IDS_ONEOFFENTRYIDFOOTERUNICODE,
				m_ProviderData.OneOffRecipientObject.Bitmask, szFlags.c_str(),
				m_ProviderData.OneOffRecipientObject.Strings.Unicode.DisplayName,
				m_ProviderData.OneOffRecipientObject.Strings.Unicode.AddressType,
				m_ProviderData.OneOffRecipientObject.Strings.Unicode.EmailAddress);
		}
		else
		{
			szEntryId += formatmessage(IDS_ONEOFFENTRYIDFOOTERANSI,
				m_ProviderData.OneOffRecipientObject.Bitmask, szFlags.c_str(),
				m_ProviderData.OneOffRecipientObject.Strings.ANSI.DisplayName,
				m_ProviderData.OneOffRecipientObject.Strings.ANSI.AddressType,
				m_ProviderData.OneOffRecipientObject.Strings.ANSI.EmailAddress);
		}
	}
	else if (eidtAddressBook == m_ObjectType)
	{
		wstring szVersion = InterpretFlags(flagExchangeABVersion, m_ProviderData.AddressBookObject.Version);
		wstring szType = InterpretNumberAsStringProp(m_ProviderData.AddressBookObject.Type, PR_DISPLAY_TYPE);

		szEntryId += formatmessage(IDS_ENTRYIDEXCHANGEADDRESSDATA,
			m_ProviderData.AddressBookObject.Version, szVersion.c_str(),
			m_ProviderData.AddressBookObject.Type, szType.c_str(),
			m_ProviderData.AddressBookObject.X500DN);
	}
	// Contact Address Book / Personal Distribution List (PDL)
	else if (eidtContact == m_ObjectType)
	{
		wstring szVersion = InterpretFlags(flagContabVersion, m_ProviderData.ContactAddressBookObject.Version);
		wstring szType = InterpretFlags(flagContabType, m_ProviderData.ContactAddressBookObject.Type);

		szEntryId += formatmessage(IDS_ENTRYIDCONTACTADDRESSDATAHEAD,
			m_ProviderData.ContactAddressBookObject.Version, szVersion.c_str(),
			m_ProviderData.ContactAddressBookObject.Type, szType.c_str());

		switch (m_ProviderData.ContactAddressBookObject.Type)
		{
		case CONTAB_USER:
		case CONTAB_DISTLIST:
			szEntryId += formatmessage(IDS_ENTRYIDCONTACTADDRESSDATAUSER,
				m_ProviderData.ContactAddressBookObject.Index, InterpretFlags(flagContabIndex, m_ProviderData.ContactAddressBookObject.Index).c_str(),
				m_ProviderData.ContactAddressBookObject.EntryIDCount);
			break;
		case CONTAB_ROOT:
		case CONTAB_CONTAINER:
		case CONTAB_SUBROOT:
			szGUID = GUIDToStringAndName((LPGUID)&m_ProviderData.ContactAddressBookObject.muidID);
			szEntryId += formatmessage(IDS_ENTRYIDCONTACTADDRESSDATACONTAINER, szGUID.c_str());
			break;
		default:
			break;
		}

		if (m_ProviderData.ContactAddressBookObject.lpEntryID)
		{
			szEntryId += m_ProviderData.ContactAddressBookObject.lpEntryID->ToString();
		}
	}
	else if (eidtWAB == m_ObjectType)
	{
		szEntryId += formatmessage(IDS_ENTRYIDWRAPPEDENTRYIDDATA, m_ProviderData.WAB.Type, InterpretFlags(flagWABEntryIDType, m_ProviderData.WAB.Type).c_str());

		if (m_ProviderData.WAB.lpEntryID)
		{
			szEntryId += m_ProviderData.WAB.lpEntryID->ToString();
		}
	}
	else if (eidtMessageDatabase == m_ObjectType)
	{
		wstring szVersion = InterpretFlags(flagMDBVersion, m_ProviderData.MessageDatabaseObject.Version);
		wstring szFlag = InterpretFlags(flagMDBFlag, m_ProviderData.MessageDatabaseObject.Flag);

		szEntryId += formatmessage(IDS_ENTRYIDMAPIMESSAGESTOREDATA,
			m_ProviderData.MessageDatabaseObject.Version, szVersion.c_str(),
			m_ProviderData.MessageDatabaseObject.Flag, szFlag.c_str(),
			m_ProviderData.MessageDatabaseObject.DLLFileName);
		if (m_ProviderData.MessageDatabaseObject.bIsExchange)
		{
			wstring szWrappedType = InterpretNumberAsStringProp(m_ProviderData.MessageDatabaseObject.WrappedType, PR_PROFILE_OPEN_FLAGS);

			szGUID = GUIDToStringAndName((LPGUID)&m_ProviderData.MessageDatabaseObject.WrappedProviderUID);
			szEntryId += formatmessage(IDS_ENTRYIDMAPIMESSAGESTOREEXCHANGEDATA,
				m_ProviderData.MessageDatabaseObject.WrappedFlags,
				szGUID.c_str(),
				m_ProviderData.MessageDatabaseObject.WrappedType, szWrappedType.c_str(),
				m_ProviderData.MessageDatabaseObject.ServerShortname,
				m_ProviderData.MessageDatabaseObject.MailboxDN ? m_ProviderData.MessageDatabaseObject.MailboxDN : ""); // STRING_OK
		}

		if (m_ProviderData.MessageDatabaseObject.bV2)
		{
			wstring szV2Magic = InterpretFlags(flagV2Magic, m_ProviderData.MessageDatabaseObject.v2.ulMagic);
			wstring szV2Version = InterpretFlags(flagV2Version, m_ProviderData.MessageDatabaseObject.v2.ulVersion);

			szEntryId += formatmessage(IDS_ENTRYIDMAPIMESSAGESTOREEXCHANGEDATAV2,
				m_ProviderData.MessageDatabaseObject.v2.ulMagic, szV2Magic.c_str(),
				m_ProviderData.MessageDatabaseObject.v2.ulSize,
				m_ProviderData.MessageDatabaseObject.v2.ulVersion, szV2Version.c_str(),
				m_ProviderData.MessageDatabaseObject.v2.ulOffsetDN,
				m_ProviderData.MessageDatabaseObject.v2.ulOffsetFQDN,
				m_ProviderData.MessageDatabaseObject.v2DN ? m_ProviderData.MessageDatabaseObject.v2DN : "", // STRING_OK
				m_ProviderData.MessageDatabaseObject.v2FQDN ? m_ProviderData.MessageDatabaseObject.v2FQDN : L""); // STRING_OK

			SBinary sBin = { 0 };
			sBin.cb = sizeof(m_ProviderData.MessageDatabaseObject.v2Reserved);
			sBin.lpb = m_ProviderData.MessageDatabaseObject.v2Reserved;
			szEntryId += BinToHexString(&sBin, true);
		}
	}
	else if (eidtFolder == m_ObjectType)
	{
		wstring szType = InterpretFlags(flagMessageDatabaseObjectType, m_ProviderData.FolderOrMessage.Type);

		wstring szDatabaseGUID = GUIDToStringAndName((LPGUID)&m_ProviderData.FolderOrMessage.Data.FolderObject.DatabaseGUID);

		SBinary sBinGlobalCounter = { 0 };
		sBinGlobalCounter.cb = sizeof(m_ProviderData.FolderOrMessage.Data.FolderObject.GlobalCounter);
		sBinGlobalCounter.lpb = m_ProviderData.FolderOrMessage.Data.FolderObject.GlobalCounter;

		SBinary sBinPad = { 0 };
		sBinPad.cb = sizeof(m_ProviderData.FolderOrMessage.Data.FolderObject.Pad);
		sBinPad.lpb = m_ProviderData.FolderOrMessage.Data.FolderObject.Pad;

		szEntryId += formatmessage(IDS_ENTRYIDEXCHANGEFOLDERDATA,
			m_ProviderData.FolderOrMessage.Type, szType.c_str(),
			szDatabaseGUID.c_str());
		szEntryId += BinToHexString(&sBinGlobalCounter, true);

		szEntryId += formatmessage(IDS_ENTRYIDEXCHANGEDATAPAD);
		szEntryId += BinToHexString(&sBinPad, true);
	}
	else if (eidtMessage == m_ObjectType)
	{
		wstring szType = InterpretFlags(flagMessageDatabaseObjectType, m_ProviderData.FolderOrMessage.Type);
		wstring szFolderDatabaseGUID = GUIDToStringAndName((LPGUID)&m_ProviderData.FolderOrMessage.Data.MessageObject.FolderDatabaseGUID);

		SBinary sBinFolderGlobalCounter = { 0 };
		sBinFolderGlobalCounter.cb = sizeof(m_ProviderData.FolderOrMessage.Data.MessageObject.FolderGlobalCounter);
		sBinFolderGlobalCounter.lpb = m_ProviderData.FolderOrMessage.Data.MessageObject.FolderGlobalCounter;

		SBinary sBinPad1 = { 0 };
		sBinPad1.cb = sizeof(m_ProviderData.FolderOrMessage.Data.MessageObject.Pad1);
		sBinPad1.lpb = m_ProviderData.FolderOrMessage.Data.MessageObject.Pad1;

		wstring szMessageDatabaseGUID = GUIDToStringAndName((LPGUID)&m_ProviderData.FolderOrMessage.Data.MessageObject.MessageDatabaseGUID);

		SBinary sBinMessageGlobalCounter = { 0 };
		sBinMessageGlobalCounter.cb = sizeof(m_ProviderData.FolderOrMessage.Data.MessageObject.MessageGlobalCounter);
		sBinMessageGlobalCounter.lpb = m_ProviderData.FolderOrMessage.Data.MessageObject.MessageGlobalCounter;

		SBinary sBinPad2 = { 0 };
		sBinPad2.cb = sizeof(m_ProviderData.FolderOrMessage.Data.MessageObject.Pad2);
		sBinPad2.lpb = m_ProviderData.FolderOrMessage.Data.MessageObject.Pad2;

		szEntryId += formatmessage(IDS_ENTRYIDEXCHANGEMESSAGEDATA,
			m_ProviderData.FolderOrMessage.Type, szType.c_str(),
			szFolderDatabaseGUID.c_str());
		szEntryId += BinToHexString(&sBinFolderGlobalCounter, true);

		szEntryId += formatmessage(IDS_ENTRYIDEXCHANGEDATAPADNUM, 1);
		szEntryId += BinToHexString(&sBinPad1, true);

		szEntryId += formatmessage(IDS_ENTRYIDEXCHANGEMESSAGEDATAGUID,
			szMessageDatabaseGUID.c_str());
		szEntryId += BinToHexString(&sBinMessageGlobalCounter, true);

		szEntryId += formatmessage(IDS_ENTRYIDEXCHANGEDATAPADNUM, 2);
		szEntryId += BinToHexString(&sBinPad2, true);
	}

	return szEntryId;
}