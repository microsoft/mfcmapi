#include "stdafx.h"
#include "EntryIdStruct.h"
#include "SmartView.h"
#include "String.h"
#include "Guids.h"
#include "Interpretprop2.h"
#include "ExtraPropTags.h"

EntryIdStruct::EntryIdStruct()
{
	memset(m_abFlags, 0, sizeof m_abFlags);
	memset(m_ProviderUID, 0, sizeof m_ProviderUID);
	m_ObjectType = eidtUnknown;
}

EntryIdStruct::~EntryIdStruct()
{
	switch (m_ObjectType)
	{
	case eidtOneOff:
		if (MAPI_UNICODE & m_OneOffRecipientObject.Bitmask)
		{
			delete[] m_OneOffRecipientObject.Unicode.DisplayName;
			delete[] m_OneOffRecipientObject.Unicode.AddressType;
			delete[] m_OneOffRecipientObject.Unicode.EmailAddress;
		}
		else
		{
			delete[] m_OneOffRecipientObject.ANSI.DisplayName;
			delete[] m_OneOffRecipientObject.ANSI.AddressType;
			delete[] m_OneOffRecipientObject.ANSI.EmailAddress;
		}
		break;
	case eidtAddressBook:
		delete[] m_AddressBookObject.X500DN;
		break;
	case eidtContact:
		delete m_ContactAddressBookObject.lpEntryID;
		break;
	case eidtMessageDatabase:
		delete[] m_MessageDatabaseObject.DLLFileName;
		delete[] m_MessageDatabaseObject.ServerShortname;
		delete[] m_MessageDatabaseObject.MailboxDN;
		delete[] m_MessageDatabaseObject.v2DN;
		delete[] m_MessageDatabaseObject.v2FQDN;
		delete[] m_MessageDatabaseObject.v3SmtpAddress;
		break;
	case eidtWAB:
		delete m_WAB.lpEntryID;
	}
}

void EntryIdStruct::Parse()
{
	m_Parser.GetBYTE(&m_abFlags[0]);
	m_Parser.GetBYTE(&m_abFlags[1]);
	m_Parser.GetBYTE(&m_abFlags[2]);
	m_Parser.GetBYTE(&m_abFlags[3]);
	m_Parser.GetBYTESNoAlloc(sizeof m_ProviderUID, sizeof m_ProviderUID, reinterpret_cast<LPBYTE>(&m_ProviderUID));

	// Ephemeral entry ID:
	if (m_abFlags[0] == EPHEMERAL)
	{
		m_ObjectType = eidtEphemeral;
	}
	// One Off Recipient
	else if (!memcmp(m_ProviderUID, &muidOOP, sizeof(GUID)))
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
		auto ulRemainingBytes = m_Parser.RemainingBytes();

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
			// Ephemeral Recipient
		case eidtEphemeral:
			m_Parser.GetDWORD(&m_EphemeralObject.Version);
			m_Parser.GetDWORD(&m_EphemeralObject.Type);
			break;
			// One Off Recipient
		case eidtOneOff:
			m_Parser.GetDWORD(&m_OneOffRecipientObject.Bitmask);
			if (MAPI_UNICODE & m_OneOffRecipientObject.Bitmask)
			{
				m_Parser.GetStringW(&m_OneOffRecipientObject.Unicode.DisplayName);
				m_Parser.GetStringW(&m_OneOffRecipientObject.Unicode.AddressType);
				m_Parser.GetStringW(&m_OneOffRecipientObject.Unicode.EmailAddress);
			}
			else
			{
				m_Parser.GetStringA(&m_OneOffRecipientObject.ANSI.DisplayName);
				m_Parser.GetStringA(&m_OneOffRecipientObject.ANSI.AddressType);
				m_Parser.GetStringA(&m_OneOffRecipientObject.ANSI.EmailAddress);
			}
			break;
			// Address Book Recipient
		case eidtAddressBook:
			m_Parser.GetDWORD(&m_AddressBookObject.Version);
			m_Parser.GetDWORD(&m_AddressBookObject.Type);
			m_Parser.GetStringA(&m_AddressBookObject.X500DN);
			break;
			// Contact Address Book / Personal Distribution List (PDL)
		case eidtContact:
		{
			m_Parser.GetDWORD(&m_ContactAddressBookObject.Version);
			m_Parser.GetDWORD(&m_ContactAddressBookObject.Type);

			if (CONTAB_CONTAINER == m_ContactAddressBookObject.Type)
			{
				m_Parser.GetBYTESNoAlloc(
					sizeof m_ContactAddressBookObject.muidID,
					sizeof m_ContactAddressBookObject.muidID,
					reinterpret_cast<LPBYTE>(&m_ContactAddressBookObject.muidID));
			}
			else // Assume we've got some variation on the user/distlist format
			{
				m_Parser.GetDWORD(&m_ContactAddressBookObject.Index);
				m_Parser.GetDWORD(&m_ContactAddressBookObject.EntryIDCount);
			}

			// Read the wrapped entry ID from the remaining data
			auto cbRemainingBytes = m_Parser.RemainingBytes();

			// If we already got a size, use it, else we just read the rest of the structure
			if (0 != m_ContactAddressBookObject.EntryIDCount &&
				m_ContactAddressBookObject.EntryIDCount < cbRemainingBytes)
			{
				cbRemainingBytes = m_ContactAddressBookObject.EntryIDCount;
			}

			m_ContactAddressBookObject.lpEntryID = new EntryIdStruct();
			if (m_ContactAddressBookObject.lpEntryID)
			{
				m_ContactAddressBookObject.lpEntryID->Init(
					static_cast<ULONG>(cbRemainingBytes),
					m_Parser.GetCurrentAddress());
				m_ContactAddressBookObject.lpEntryID->DisableJunkParsing();
				m_ContactAddressBookObject.lpEntryID->EnsureParsed();
				m_Parser.Advance(m_ContactAddressBookObject.lpEntryID->GetCurrentOffset());
			}
		}
		break;
		case eidtWAB:
		{
			m_ObjectType = eidtWAB;

			m_Parser.GetBYTE(&m_WAB.Type);

			m_WAB.lpEntryID = new EntryIdStruct();
			if (m_WAB.lpEntryID)
			{
				m_WAB.lpEntryID->Init(
					static_cast<ULONG>(m_Parser.RemainingBytes()),
					m_Parser.GetCurrentAddress());
				m_WAB.lpEntryID->DisableJunkParsing();
				m_WAB.lpEntryID->EnsureParsed();
				m_Parser.Advance(m_WAB.lpEntryID->GetCurrentOffset());
			}
		}
		break;
		// message store objects
		case eidtMessageDatabase:
			m_Parser.GetBYTE(&m_MessageDatabaseObject.Version);
			m_Parser.GetBYTE(&m_MessageDatabaseObject.Flag);
			m_Parser.GetStringA(&m_MessageDatabaseObject.DLLFileName);
			m_MessageDatabaseObject.bIsExchange = false;

			// We only know how to parse emsmdb.dll's wrapped entry IDs
			if (m_MessageDatabaseObject.DLLFileName &&
				CSTR_EQUAL == CompareStringA(LOCALE_INVARIANT,
					NORM_IGNORECASE,
					m_MessageDatabaseObject.DLLFileName,
					-1,
					"emsmdb.dll", // STRING_OK
					-1))
			{
				m_MessageDatabaseObject.bIsExchange = true;
				auto cbRead = m_Parser.GetCurrentOffset();
				// Advance to the next multiple of 4
				m_Parser.Advance(3 - (cbRead + 3) % 4);
				m_Parser.GetDWORD(&m_MessageDatabaseObject.WrappedFlags);
				m_Parser.GetBYTESNoAlloc(sizeof m_MessageDatabaseObject.WrappedProviderUID,
					sizeof m_MessageDatabaseObject.WrappedProviderUID,
					reinterpret_cast<LPBYTE>(&m_MessageDatabaseObject.WrappedProviderUID));
				m_Parser.GetDWORD(&m_MessageDatabaseObject.WrappedType);
				m_Parser.GetStringA(&m_MessageDatabaseObject.ServerShortname);

				m_MessageDatabaseObject.MagicVersion = MDB_STORE_EID_V1_VERSION;

				// Test if we have a magic value. Some PF EIDs also have a mailbox DN and we need to accomodate them
				if (m_MessageDatabaseObject.WrappedType & OPENSTORE_PUBLIC)
				{
					cbRead = m_Parser.GetCurrentOffset();
					m_Parser.GetDWORD(&m_MessageDatabaseObject.MagicVersion);
					m_Parser.SetCurrentOffset(cbRead);
				}

				// Either we're not a PF eid, or this PF EID wasn't followed directly by a magic value
				if (!(m_MessageDatabaseObject.WrappedType & OPENSTORE_PUBLIC) ||
					m_MessageDatabaseObject.MagicVersion != MDB_STORE_EID_V2_MAGIC &&
					m_MessageDatabaseObject.MagicVersion != MDB_STORE_EID_V3_MAGIC)
				{
					m_Parser.GetStringA(&m_MessageDatabaseObject.MailboxDN);
				}

				// Check again for a magic value
				cbRead = m_Parser.GetCurrentOffset();
				m_Parser.GetDWORD(&m_MessageDatabaseObject.MagicVersion);
				m_Parser.SetCurrentOffset(cbRead);

				switch (m_MessageDatabaseObject.MagicVersion)
				{
				case MDB_STORE_EID_V2_MAGIC:
					if (m_Parser.RemainingBytes() >= sizeof(MDB_STORE_EID_V2) + sizeof(WCHAR))
					{
						m_Parser.GetDWORD(&m_MessageDatabaseObject.v2.ulMagic);
						m_Parser.GetDWORD(&m_MessageDatabaseObject.v2.ulSize);
						m_Parser.GetDWORD(&m_MessageDatabaseObject.v2.ulVersion);
						m_Parser.GetDWORD(&m_MessageDatabaseObject.v2.ulOffsetDN);
						m_Parser.GetDWORD(&m_MessageDatabaseObject.v2.ulOffsetFQDN);
						if (m_MessageDatabaseObject.v2.ulOffsetDN)
						{
							m_Parser.GetStringA(&m_MessageDatabaseObject.v2DN);
						}

						if (m_MessageDatabaseObject.v2.ulOffsetFQDN)
						{
							m_Parser.GetStringW(&m_MessageDatabaseObject.v2FQDN);
						}

						m_Parser.GetBYTESNoAlloc(sizeof m_MessageDatabaseObject.v2Reserved,
							sizeof m_MessageDatabaseObject.v2Reserved,
							reinterpret_cast<LPBYTE>(&m_MessageDatabaseObject.v2Reserved));
					}
					break;
				case MDB_STORE_EID_V3_MAGIC:
					if (m_Parser.RemainingBytes() >= sizeof(MDB_STORE_EID_V3) + sizeof(WCHAR))
					{
						m_Parser.GetDWORD(&m_MessageDatabaseObject.v3.ulMagic);
						m_Parser.GetDWORD(&m_MessageDatabaseObject.v3.ulSize);
						m_Parser.GetDWORD(&m_MessageDatabaseObject.v3.ulVersion);
						m_Parser.GetDWORD(&m_MessageDatabaseObject.v3.ulOffsetSmtpAddress);
						if (m_MessageDatabaseObject.v3.ulOffsetSmtpAddress)
						{
							m_Parser.GetStringW(&m_MessageDatabaseObject.v3SmtpAddress);
						}

						m_Parser.GetBYTESNoAlloc(sizeof m_MessageDatabaseObject.v2Reserved,
							sizeof m_MessageDatabaseObject.v2Reserved,
							reinterpret_cast<LPBYTE>(&m_MessageDatabaseObject.v2Reserved));
					}
					break;
				}
			}
			break;
			// Exchange message store folder
		case eidtFolder:
			m_Parser.GetWORD(&m_FolderOrMessage.Type);
			m_Parser.GetBYTESNoAlloc(sizeof m_FolderOrMessage.FolderObject.DatabaseGUID,
				sizeof m_FolderOrMessage.FolderObject.DatabaseGUID,
				reinterpret_cast<LPBYTE>(&m_FolderOrMessage.FolderObject.DatabaseGUID));
			m_Parser.GetBYTESNoAlloc(sizeof m_FolderOrMessage.FolderObject.GlobalCounter,
				sizeof m_FolderOrMessage.FolderObject.GlobalCounter,
				reinterpret_cast<LPBYTE>(&m_FolderOrMessage.FolderObject.GlobalCounter));
			m_Parser.GetBYTESNoAlloc(sizeof m_FolderOrMessage.FolderObject.Pad,
				sizeof m_FolderOrMessage.FolderObject.Pad,
				reinterpret_cast<LPBYTE>(&m_FolderOrMessage.FolderObject.Pad));
			break;
			// Exchange message store message
		case eidtMessage:
			m_Parser.GetWORD(&m_FolderOrMessage.Type);
			m_Parser.GetBYTESNoAlloc(sizeof m_FolderOrMessage.MessageObject.FolderDatabaseGUID,
				sizeof m_FolderOrMessage.MessageObject.FolderDatabaseGUID,
				reinterpret_cast<LPBYTE>(&m_FolderOrMessage.MessageObject.FolderDatabaseGUID));
			m_Parser.GetBYTESNoAlloc(sizeof m_FolderOrMessage.MessageObject.FolderGlobalCounter,
				sizeof m_FolderOrMessage.MessageObject.FolderGlobalCounter,
				reinterpret_cast<LPBYTE>(&m_FolderOrMessage.MessageObject.FolderGlobalCounter));
			m_Parser.GetBYTESNoAlloc(sizeof m_FolderOrMessage.MessageObject.Pad1,
				sizeof m_FolderOrMessage.MessageObject.Pad1,
				reinterpret_cast<LPBYTE>(&m_FolderOrMessage.MessageObject.Pad1));
			m_Parser.GetBYTESNoAlloc(sizeof m_FolderOrMessage.MessageObject.MessageDatabaseGUID,
				sizeof m_FolderOrMessage.MessageObject.MessageDatabaseGUID,
				reinterpret_cast<LPBYTE>(&m_FolderOrMessage.MessageObject.MessageDatabaseGUID));
			m_Parser.GetBYTESNoAlloc(sizeof m_FolderOrMessage.MessageObject.MessageGlobalCounter,
				sizeof m_FolderOrMessage.MessageObject.MessageGlobalCounter,
				reinterpret_cast<LPBYTE>(&m_FolderOrMessage.MessageObject.MessageGlobalCounter));
			m_Parser.GetBYTESNoAlloc(sizeof m_FolderOrMessage.MessageObject.Pad2,
				sizeof m_FolderOrMessage.MessageObject.Pad2,
				reinterpret_cast<LPBYTE>(&m_FolderOrMessage.MessageObject.Pad2));
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
	case eidtEphemeral:
		szEntryId = formatmessage(IDS_ENTRYIDEPHEMERALADDRESS);
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

	auto szGUID = GUIDToStringAndName(reinterpret_cast<LPGUID>(&m_ProviderUID));
	szEntryId += formatmessage(IDS_ENTRYIDPROVIDERGUID, szGUID.c_str());

	if (eidtEphemeral == m_ObjectType)
	{
		auto szVersion = InterpretFlags(flagExchangeABVersion, m_EphemeralObject.Version);
		auto szType = InterpretNumberAsStringProp(m_EphemeralObject.Type, PR_DISPLAY_TYPE);

		szEntryId += formatmessage(IDS_ENTRYIDEPHEMERALADDRESSDATA,
			m_AddressBookObject.Version, szVersion.c_str(),
			m_AddressBookObject.Type, szType.c_str());
	}
	else if (eidtOneOff == m_ObjectType)
	{
		auto szFlags = InterpretFlags(flagOneOffEntryId, m_OneOffRecipientObject.Bitmask);
		if (MAPI_UNICODE & m_OneOffRecipientObject.Bitmask)
		{
			szEntryId += formatmessage(IDS_ONEOFFENTRYIDFOOTERUNICODE,
				m_OneOffRecipientObject.Bitmask, szFlags.c_str(),
				m_OneOffRecipientObject.Unicode.DisplayName,
				m_OneOffRecipientObject.Unicode.AddressType,
				m_OneOffRecipientObject.Unicode.EmailAddress);
		}
		else
		{
			szEntryId += formatmessage(IDS_ONEOFFENTRYIDFOOTERANSI,
				m_OneOffRecipientObject.Bitmask, szFlags.c_str(),
				m_OneOffRecipientObject.ANSI.DisplayName,
				m_OneOffRecipientObject.ANSI.AddressType,
				m_OneOffRecipientObject.ANSI.EmailAddress);
		}
	}
	else if (eidtAddressBook == m_ObjectType)
	{
		auto szVersion = InterpretFlags(flagExchangeABVersion, m_AddressBookObject.Version);
		auto szType = InterpretNumberAsStringProp(m_AddressBookObject.Type, PR_DISPLAY_TYPE);

		szEntryId += formatmessage(IDS_ENTRYIDEXCHANGEADDRESSDATA,
			m_AddressBookObject.Version, szVersion.c_str(),
			m_AddressBookObject.Type, szType.c_str(),
			m_AddressBookObject.X500DN);
	}
	// Contact Address Book / Personal Distribution List (PDL)
	else if (eidtContact == m_ObjectType)
	{
		auto szVersion = InterpretFlags(flagContabVersion, m_ContactAddressBookObject.Version);
		auto szType = InterpretFlags(flagContabType, m_ContactAddressBookObject.Type);

		szEntryId += formatmessage(IDS_ENTRYIDCONTACTADDRESSDATAHEAD,
			m_ContactAddressBookObject.Version, szVersion.c_str(),
			m_ContactAddressBookObject.Type, szType.c_str());

		switch (m_ContactAddressBookObject.Type)
		{
		case CONTAB_USER:
		case CONTAB_DISTLIST:
			szEntryId += formatmessage(IDS_ENTRYIDCONTACTADDRESSDATAUSER,
				m_ContactAddressBookObject.Index, InterpretFlags(flagContabIndex, m_ContactAddressBookObject.Index).c_str(),
				m_ContactAddressBookObject.EntryIDCount);
			break;
		case CONTAB_ROOT:
		case CONTAB_CONTAINER:
		case CONTAB_SUBROOT:
			szGUID = GUIDToStringAndName(reinterpret_cast<LPGUID>(&m_ContactAddressBookObject.muidID));
			szEntryId += formatmessage(IDS_ENTRYIDCONTACTADDRESSDATACONTAINER, szGUID.c_str());
			break;
		default:
			break;
		}

		if (m_ContactAddressBookObject.lpEntryID)
		{
			szEntryId += m_ContactAddressBookObject.lpEntryID->ToString();
		}
	}
	else if (eidtWAB == m_ObjectType)
	{
		szEntryId += formatmessage(IDS_ENTRYIDWRAPPEDENTRYIDDATA, m_WAB.Type, InterpretFlags(flagWABEntryIDType, m_WAB.Type).c_str());

		if (m_WAB.lpEntryID)
		{
			szEntryId += m_WAB.lpEntryID->ToString();
		}
	}
	else if (eidtMessageDatabase == m_ObjectType)
	{
		auto szVersion = InterpretFlags(flagMDBVersion, m_MessageDatabaseObject.Version);
		auto szFlag = InterpretFlags(flagMDBFlag, m_MessageDatabaseObject.Flag);

		szEntryId += formatmessage(IDS_ENTRYIDMAPIMESSAGESTOREDATA,
			m_MessageDatabaseObject.Version, szVersion.c_str(),
			m_MessageDatabaseObject.Flag, szFlag.c_str(),
			m_MessageDatabaseObject.DLLFileName);
		if (m_MessageDatabaseObject.bIsExchange)
		{
			auto szWrappedType = InterpretNumberAsStringProp(m_MessageDatabaseObject.WrappedType, PR_PROFILE_OPEN_FLAGS);

			szGUID = GUIDToStringAndName(reinterpret_cast<LPGUID>(&m_MessageDatabaseObject.WrappedProviderUID));
			szEntryId += formatmessage(IDS_ENTRYIDMAPIMESSAGESTOREEXCHANGEDATA,
				m_MessageDatabaseObject.WrappedFlags,
				szGUID.c_str(),
				m_MessageDatabaseObject.WrappedType, szWrappedType.c_str(),
				m_MessageDatabaseObject.ServerShortname,
				m_MessageDatabaseObject.MailboxDN ? m_MessageDatabaseObject.MailboxDN : ""); // STRING_OK
		}

		switch (m_MessageDatabaseObject.MagicVersion)
		{
		case MDB_STORE_EID_V2_MAGIC:
		{
			auto szV2Magic = InterpretFlags(flagEidMagic, m_MessageDatabaseObject.v2.ulMagic);
			auto szV2Version = InterpretFlags(flagEidVersion, m_MessageDatabaseObject.v2.ulVersion);

			szEntryId += formatmessage(IDS_ENTRYIDMAPIMESSAGESTOREEXCHANGEDATAV2,
				m_MessageDatabaseObject.v2.ulMagic, szV2Magic.c_str(),
				m_MessageDatabaseObject.v2.ulSize,
				m_MessageDatabaseObject.v2.ulVersion, szV2Version.c_str(),
				m_MessageDatabaseObject.v2.ulOffsetDN,
				m_MessageDatabaseObject.v2.ulOffsetFQDN,
				m_MessageDatabaseObject.v2DN ? m_MessageDatabaseObject.v2DN : "", // STRING_OK
				m_MessageDatabaseObject.v2FQDN ? m_MessageDatabaseObject.v2FQDN : L""); // STRING_OK

			SBinary sBin = { 0 };
			sBin.cb = sizeof m_MessageDatabaseObject.v2Reserved;
			sBin.lpb = m_MessageDatabaseObject.v2Reserved;
			szEntryId += BinToHexString(&sBin, true);
		}
		break;
		case MDB_STORE_EID_V3_MAGIC:
		{
			auto szV3Magic = InterpretFlags(flagEidMagic, m_MessageDatabaseObject.v3.ulMagic);
			auto szV3Version = InterpretFlags(flagEidVersion, m_MessageDatabaseObject.v3.ulVersion);

			szEntryId += formatmessage(IDS_ENTRYIDMAPIMESSAGESTOREEXCHANGEDATAV3,
				m_MessageDatabaseObject.v3.ulMagic, szV3Magic.c_str(),
				m_MessageDatabaseObject.v3.ulSize,
				m_MessageDatabaseObject.v3.ulVersion, szV3Version.c_str(),
				m_MessageDatabaseObject.v3.ulOffsetSmtpAddress,
				m_MessageDatabaseObject.v3SmtpAddress ? m_MessageDatabaseObject.v3SmtpAddress : L""); // STRING_OK

			SBinary sBin = { 0 };
			sBin.cb = sizeof m_MessageDatabaseObject.v2Reserved;
			sBin.lpb = m_MessageDatabaseObject.v2Reserved;
			szEntryId += BinToHexString(&sBin, true);
		}
		break;
		}
	}
	else if (eidtFolder == m_ObjectType)
	{
		auto szType = InterpretFlags(flagMessageDatabaseObjectType, m_FolderOrMessage.Type);

		auto szDatabaseGUID = GUIDToStringAndName(reinterpret_cast<LPGUID>(&m_FolderOrMessage.FolderObject.DatabaseGUID));

		SBinary sBinGlobalCounter = { 0 };
		sBinGlobalCounter.cb = sizeof m_FolderOrMessage.FolderObject.GlobalCounter;
		sBinGlobalCounter.lpb = m_FolderOrMessage.FolderObject.GlobalCounter;

		SBinary sBinPad = { 0 };
		sBinPad.cb = sizeof m_FolderOrMessage.FolderObject.Pad;
		sBinPad.lpb = m_FolderOrMessage.FolderObject.Pad;

		szEntryId += formatmessage(IDS_ENTRYIDEXCHANGEFOLDERDATA,
			m_FolderOrMessage.Type, szType.c_str(),
			szDatabaseGUID.c_str());
		szEntryId += BinToHexString(&sBinGlobalCounter, true);

		szEntryId += formatmessage(IDS_ENTRYIDEXCHANGEDATAPAD);
		szEntryId += BinToHexString(&sBinPad, true);
	}
	else if (eidtMessage == m_ObjectType)
	{
		auto szType = InterpretFlags(flagMessageDatabaseObjectType, m_FolderOrMessage.Type);
		auto szFolderDatabaseGUID = GUIDToStringAndName(reinterpret_cast<LPGUID>(&m_FolderOrMessage.MessageObject.FolderDatabaseGUID));

		SBinary sBinFolderGlobalCounter = { 0 };
		sBinFolderGlobalCounter.cb = sizeof m_FolderOrMessage.MessageObject.FolderGlobalCounter;
		sBinFolderGlobalCounter.lpb = m_FolderOrMessage.MessageObject.FolderGlobalCounter;

		SBinary sBinPad1 = { 0 };
		sBinPad1.cb = sizeof m_FolderOrMessage.MessageObject.Pad1;
		sBinPad1.lpb = m_FolderOrMessage.MessageObject.Pad1;

		auto szMessageDatabaseGUID = GUIDToStringAndName(reinterpret_cast<LPGUID>(&m_FolderOrMessage.MessageObject.MessageDatabaseGUID));

		SBinary sBinMessageGlobalCounter = { 0 };
		sBinMessageGlobalCounter.cb = sizeof m_FolderOrMessage.MessageObject.MessageGlobalCounter;
		sBinMessageGlobalCounter.lpb = m_FolderOrMessage.MessageObject.MessageGlobalCounter;

		SBinary sBinPad2 = { 0 };
		sBinPad2.cb = sizeof m_FolderOrMessage.MessageObject.Pad2;
		sBinPad2.lpb = m_FolderOrMessage.MessageObject.Pad2;

		szEntryId += formatmessage(IDS_ENTRYIDEXCHANGEMESSAGEDATA,
			m_FolderOrMessage.Type, szType.c_str(),
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