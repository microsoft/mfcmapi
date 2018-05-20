#include "StdAfx.h"
#include "EntryIdStruct.h"
#include <Interpret/SmartView/SmartView.h>
#include <Interpret/String.h>
#include <Interpret/Guids.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>

namespace smartview
{
	EntryIdStruct::EntryIdStruct() : m_FolderOrMessage(), m_MessageDatabaseObject(), m_AddressBookObject(),
		m_ContactAddressBookObject(), m_WAB()
	{
		m_ObjectType = eidtUnknown;
	}

	void EntryIdStruct::Parse()
	{
		m_abFlags.resize(4);
		m_abFlags[0] = m_Parser.Get<BYTE>();
		m_abFlags[1] = m_Parser.Get<BYTE>();
		m_abFlags[2] = m_Parser.Get<BYTE>();
		m_abFlags[3] = m_Parser.Get<BYTE>();
		m_ProviderUID = m_Parser.Get<GUID>();

		// Ephemeral entry ID:
		if (m_abFlags[0] == EPHEMERAL)
		{
			m_ObjectType = eidtEphemeral;
		}
		// One Off Recipient
		else if (!memcmp(&m_ProviderUID, &guid::muidOOP, sizeof(GUID)))
		{
			m_ObjectType = eidtOneOff;
		}
		// Address Book Recipient
		else if (!memcmp(&m_ProviderUID, &guid::muidEMSAB, sizeof(GUID)))
		{
			m_ObjectType = eidtAddressBook;
		}
		// Contact Address Book / Personal Distribution List (PDL)
		else if (!memcmp(&m_ProviderUID, &guid::muidContabDLL, sizeof(GUID)))
		{
			m_ObjectType = eidtContact;
		}
		// message store objects
		else if (!memcmp(&m_ProviderUID, &guid::muidStoreWrap, sizeof(GUID)))
		{
			m_ObjectType = eidtMessageDatabase;
		}
		else if (!memcmp(&m_ProviderUID, &guid::WAB_GUID, sizeof(GUID)))
		{
			m_ObjectType = eidtWAB;
		}
		// We can recognize Exchange message store folder and message entry IDs by their size
		else
		{
			const auto ulRemainingBytes = m_Parser.RemainingBytes();

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
				m_EphemeralObject.Version = m_Parser.Get<DWORD>();
				m_EphemeralObject.Type = m_Parser.Get<DWORD>();
				break;
				// One Off Recipient
			case eidtOneOff:
				m_OneOffRecipientObject.Bitmask = m_Parser.Get<DWORD>();
				if (MAPI_UNICODE & m_OneOffRecipientObject.Bitmask)
				{
					m_OneOffRecipientObject.Unicode.DisplayName = m_Parser.GetStringW();
					m_OneOffRecipientObject.Unicode.AddressType = m_Parser.GetStringW();
					m_OneOffRecipientObject.Unicode.EmailAddress = m_Parser.GetStringW();
				}
				else
				{
					m_OneOffRecipientObject.ANSI.DisplayName = m_Parser.GetStringA();
					m_OneOffRecipientObject.ANSI.AddressType = m_Parser.GetStringA();
					m_OneOffRecipientObject.ANSI.EmailAddress = m_Parser.GetStringA();
				}
				break;
				// Address Book Recipient
			case eidtAddressBook:
				m_AddressBookObject.Version = m_Parser.Get<DWORD>();
				m_AddressBookObject.Type = m_Parser.Get<DWORD>();
				m_AddressBookObject.X500DN = m_Parser.GetStringA();
				break;
				// Contact Address Book / Personal Distribution List (PDL)
			case eidtContact:
			{
				m_ContactAddressBookObject.Version = m_Parser.Get<DWORD>();
				m_ContactAddressBookObject.Type = m_Parser.Get<DWORD>();

				if (CONTAB_CONTAINER == m_ContactAddressBookObject.Type)
				{
					m_ContactAddressBookObject.muidID = m_Parser.Get<GUID>();
				}
				else // Assume we've got some variation on the user/distlist format
				{
					m_ContactAddressBookObject.Index = m_Parser.Get<DWORD>();
					m_ContactAddressBookObject.EntryIDCount = m_Parser.Get<DWORD>();
				}

				// Read the wrapped entry ID from the remaining data
				auto cbRemainingBytes = m_Parser.RemainingBytes();

				// If we already got a size, use it, else we just read the rest of the structure
				if (0 != m_ContactAddressBookObject.EntryIDCount &&
					m_ContactAddressBookObject.EntryIDCount < cbRemainingBytes)
				{
					cbRemainingBytes = m_ContactAddressBookObject.EntryIDCount;
				}

				EntryIdStruct entryIdStruct;
				entryIdStruct.Init(cbRemainingBytes, m_Parser.GetCurrentAddress());
				entryIdStruct.DisableJunkParsing();
				entryIdStruct.EnsureParsed();
				m_Parser.Advance(entryIdStruct.GetCurrentOffset());
				m_ContactAddressBookObject.lpEntryID.push_back(entryIdStruct);
			}
			break;
			case eidtWAB:
			{
				m_ObjectType = eidtWAB;

				m_WAB.Type = m_Parser.Get<BYTE>();

				EntryIdStruct entryIdStruct;
				entryIdStruct.Init(
					m_Parser.RemainingBytes(),
					m_Parser.GetCurrentAddress());
				entryIdStruct.DisableJunkParsing();
				entryIdStruct.EnsureParsed();
				m_Parser.Advance(entryIdStruct.GetCurrentOffset());
				m_WAB.lpEntryID.push_back(entryIdStruct);
			}
			break;
			// message store objects
			case eidtMessageDatabase:
				m_MessageDatabaseObject.Version = m_Parser.Get<BYTE>();
				m_MessageDatabaseObject.Flag = m_Parser.Get<BYTE>();
				m_MessageDatabaseObject.DLLFileName = m_Parser.GetStringA();
				m_MessageDatabaseObject.bIsExchange = false;

				// We only know how to parse emsmdb.dll's wrapped entry IDs
				if (!m_MessageDatabaseObject.DLLFileName.empty() &&
					CSTR_EQUAL == CompareStringA(LOCALE_INVARIANT,
						NORM_IGNORECASE,
						m_MessageDatabaseObject.DLLFileName.c_str(),
						-1,
						"emsmdb.dll", // STRING_OK
						-1))
				{
					m_MessageDatabaseObject.bIsExchange = true;
					auto cbRead = m_Parser.GetCurrentOffset();
					// Advance to the next multiple of 4
					m_Parser.Advance(3 - (cbRead + 3) % 4);
					m_MessageDatabaseObject.WrappedFlags = m_Parser.Get<DWORD>();
					m_MessageDatabaseObject.WrappedProviderUID = m_Parser.Get<GUID>();
					m_MessageDatabaseObject.WrappedType = m_Parser.Get<DWORD>();
					m_MessageDatabaseObject.ServerShortname = m_Parser.GetStringA();

					m_MessageDatabaseObject.MagicVersion = MDB_STORE_EID_V1_VERSION;

					// Test if we have a magic value. Some PF EIDs also have a mailbox DN and we need to accomodate them
					if (m_MessageDatabaseObject.WrappedType & OPENSTORE_PUBLIC)
					{
						cbRead = m_Parser.GetCurrentOffset();
						m_MessageDatabaseObject.MagicVersion = m_Parser.Get<DWORD>();
						m_Parser.SetCurrentOffset(cbRead);
					}

					// Either we're not a PF eid, or this PF EID wasn't followed directly by a magic value
					if (!(m_MessageDatabaseObject.WrappedType & OPENSTORE_PUBLIC) ||
						m_MessageDatabaseObject.MagicVersion != MDB_STORE_EID_V2_MAGIC &&
						m_MessageDatabaseObject.MagicVersion != MDB_STORE_EID_V3_MAGIC)
					{
						m_MessageDatabaseObject.MailboxDN = m_Parser.GetStringA();
					}

					// Check again for a magic value
					cbRead = m_Parser.GetCurrentOffset();
					m_MessageDatabaseObject.MagicVersion = m_Parser.Get<DWORD>();
					m_Parser.SetCurrentOffset(cbRead);

					switch (m_MessageDatabaseObject.MagicVersion)
					{
					case MDB_STORE_EID_V2_MAGIC:
						if (m_Parser.RemainingBytes() >= sizeof(MDB_STORE_EID_V2) + sizeof(WCHAR))
						{
							m_MessageDatabaseObject.v2.ulMagic = m_Parser.Get<DWORD>();
							m_MessageDatabaseObject.v2.ulSize = m_Parser.Get<DWORD>();
							m_MessageDatabaseObject.v2.ulVersion = m_Parser.Get<DWORD>();
							m_MessageDatabaseObject.v2.ulOffsetDN = m_Parser.Get<DWORD>();
							m_MessageDatabaseObject.v2.ulOffsetFQDN = m_Parser.Get<DWORD>();
							if (m_MessageDatabaseObject.v2.ulOffsetDN)
							{
								m_MessageDatabaseObject.v2DN = m_Parser.GetStringA();
							}

							if (m_MessageDatabaseObject.v2.ulOffsetFQDN)
							{
								m_MessageDatabaseObject.v2FQDN = m_Parser.GetStringW();
							}

							m_MessageDatabaseObject.v2Reserved = m_Parser.GetBYTES(2);
						}
						break;
					case MDB_STORE_EID_V3_MAGIC:
						if (m_Parser.RemainingBytes() >= sizeof(MDB_STORE_EID_V3) + sizeof(WCHAR))
						{
							m_MessageDatabaseObject.v3.ulMagic = m_Parser.Get<DWORD>();
							m_MessageDatabaseObject.v3.ulSize = m_Parser.Get<DWORD>();
							m_MessageDatabaseObject.v3.ulVersion = m_Parser.Get<DWORD>();
							m_MessageDatabaseObject.v3.ulOffsetSmtpAddress = m_Parser.Get<DWORD>();
							if (m_MessageDatabaseObject.v3.ulOffsetSmtpAddress)
							{
								m_MessageDatabaseObject.v3SmtpAddress = m_Parser.GetStringW();
							}

							m_MessageDatabaseObject.v2Reserved = m_Parser.GetBYTES(2);
						}
						break;
					}
				}
				break;
				// Exchange message store folder
			case eidtFolder:
				m_FolderOrMessage.Type = m_Parser.Get<WORD>();
				m_FolderOrMessage.FolderObject.DatabaseGUID = m_Parser.Get<GUID>();
				m_FolderOrMessage.FolderObject.GlobalCounter = m_Parser.GetBYTES(6);
				m_FolderOrMessage.FolderObject.Pad = m_Parser.GetBYTES(2);
				break;
				// Exchange message store message
			case eidtMessage:
				m_FolderOrMessage.Type = m_Parser.Get<WORD>();
				m_FolderOrMessage.MessageObject.FolderDatabaseGUID = m_Parser.Get<GUID>();
				m_FolderOrMessage.MessageObject.FolderGlobalCounter = m_Parser.GetBYTES(6);
				m_FolderOrMessage.MessageObject.Pad1 = m_Parser.GetBYTES(2);

				m_FolderOrMessage.MessageObject.MessageDatabaseGUID = m_Parser.Get<GUID>();
				m_FolderOrMessage.MessageObject.MessageGlobalCounter = m_Parser.GetBYTES(6);
				m_FolderOrMessage.MessageObject.Pad2 = m_Parser.GetBYTES(2);
				break;
			}
		}

		// Check if we have an unidentified short term entry ID:
		if (eidtUnknown == m_ObjectType && m_abFlags[0] & MAPI_SHORTTERM)
			m_ObjectType = eidtShortTerm;
	}

	_Check_return_ std::wstring EntryIdStruct::ToStringInternal()
	{
		std::wstring szEntryId;

		switch (m_ObjectType)
		{
		case eidtUnknown:
			szEntryId = strings::formatmessage(IDS_ENTRYIDUNKNOWN);
			break;
		case eidtEphemeral:
			szEntryId = strings::formatmessage(IDS_ENTRYIDEPHEMERALADDRESS);
			break;
		case eidtShortTerm:
			szEntryId = strings::formatmessage(IDS_ENTRYIDSHORTTERM);
			break;
		case eidtFolder:
			szEntryId = strings::formatmessage(IDS_ENTRYIDEXCHANGEFOLDER);
			break;
		case eidtMessage:
			szEntryId = strings::formatmessage(IDS_ENTRYIDEXCHANGEMESSAGE);
			break;
		case eidtMessageDatabase:
			szEntryId = strings::formatmessage(IDS_ENTRYIDMAPIMESSAGESTORE);
			break;
		case eidtOneOff:
			szEntryId = strings::formatmessage(IDS_ENTRYIDMAPIONEOFF);
			break;
		case eidtAddressBook:
			szEntryId = strings::formatmessage(IDS_ENTRYIDEXCHANGEADDRESS);
			break;
		case eidtContact:
			szEntryId = strings::formatmessage(IDS_ENTRYIDCONTACTADDRESS);
			break;
		case eidtWAB:
			szEntryId = strings::formatmessage(IDS_ENTRYIDWRAPPEDENTRYID);
			break;
		}

		if (0 == (m_abFlags[0] | m_abFlags[1] | m_abFlags[2] | m_abFlags[3]))
		{
			szEntryId += strings::formatmessage(IDS_ENTRYIDHEADER1);
		}
		else if (0 == (m_abFlags[1] | m_abFlags[2] | m_abFlags[3]))
		{
			szEntryId += strings::formatmessage(IDS_ENTRYIDHEADER2,
				m_abFlags[0], InterpretFlags(flagEntryId0, m_abFlags[0]).c_str());
		}
		else
		{
			szEntryId += strings::formatmessage(IDS_ENTRYIDHEADER3,
				m_abFlags[0], InterpretFlags(flagEntryId0, m_abFlags[0]).c_str(),
				m_abFlags[1], InterpretFlags(flagEntryId1, m_abFlags[1]).c_str(),
				m_abFlags[2],
				m_abFlags[3]);
		}

		auto szGUID = guid::GUIDToStringAndName(&m_ProviderUID);
		szEntryId += strings::formatmessage(IDS_ENTRYIDPROVIDERGUID, szGUID.c_str());

		if (eidtEphemeral == m_ObjectType)
		{
			auto szVersion = InterpretFlags(flagExchangeABVersion, m_EphemeralObject.Version);
			auto szType = InterpretNumberAsStringProp(m_EphemeralObject.Type, PR_DISPLAY_TYPE);

			szEntryId += strings::formatmessage(IDS_ENTRYIDEPHEMERALADDRESSDATA,
				m_EphemeralObject.Version, szVersion.c_str(),
				m_EphemeralObject.Type, szType.c_str());
		}
		else if (eidtOneOff == m_ObjectType)
		{
			auto szFlags = InterpretFlags(flagOneOffEntryId, m_OneOffRecipientObject.Bitmask);
			if (MAPI_UNICODE & m_OneOffRecipientObject.Bitmask)
			{
				szEntryId += strings::formatmessage(IDS_ONEOFFENTRYIDFOOTERUNICODE,
					m_OneOffRecipientObject.Bitmask, szFlags.c_str(),
					m_OneOffRecipientObject.Unicode.DisplayName.c_str(),
					m_OneOffRecipientObject.Unicode.AddressType.c_str(),
					m_OneOffRecipientObject.Unicode.EmailAddress.c_str());
			}
			else
			{
				szEntryId += strings::formatmessage(IDS_ONEOFFENTRYIDFOOTERANSI,
					m_OneOffRecipientObject.Bitmask, szFlags.c_str(),
					m_OneOffRecipientObject.ANSI.DisplayName.c_str(),
					m_OneOffRecipientObject.ANSI.AddressType.c_str(),
					m_OneOffRecipientObject.ANSI.EmailAddress.c_str());
			}
		}
		else if (eidtAddressBook == m_ObjectType)
		{
			auto szVersion = InterpretFlags(flagExchangeABVersion, m_AddressBookObject.Version);
			auto szType = InterpretNumberAsStringProp(m_AddressBookObject.Type, PR_DISPLAY_TYPE);

			szEntryId += strings::formatmessage(IDS_ENTRYIDEXCHANGEADDRESSDATA,
				m_AddressBookObject.Version, szVersion.c_str(),
				m_AddressBookObject.Type, szType.c_str(),
				m_AddressBookObject.X500DN.c_str());
		}
		// Contact Address Book / Personal Distribution List (PDL)
		else if (eidtContact == m_ObjectType)
		{
			auto szVersion = InterpretFlags(flagContabVersion, m_ContactAddressBookObject.Version);
			auto szType = InterpretFlags(flagContabType, m_ContactAddressBookObject.Type);

			szEntryId += strings::formatmessage(IDS_ENTRYIDCONTACTADDRESSDATAHEAD,
				m_ContactAddressBookObject.Version, szVersion.c_str(),
				m_ContactAddressBookObject.Type, szType.c_str());

			switch (m_ContactAddressBookObject.Type)
			{
			case CONTAB_USER:
			case CONTAB_DISTLIST:
				szEntryId += strings::formatmessage(IDS_ENTRYIDCONTACTADDRESSDATAUSER,
					m_ContactAddressBookObject.Index, InterpretFlags(flagContabIndex, m_ContactAddressBookObject.Index).c_str(),
					m_ContactAddressBookObject.EntryIDCount);
				break;
			case CONTAB_ROOT:
			case CONTAB_CONTAINER:
			case CONTAB_SUBROOT:
				szGUID = guid::GUIDToStringAndName(&m_ContactAddressBookObject.muidID);
				szEntryId += strings::formatmessage(IDS_ENTRYIDCONTACTADDRESSDATACONTAINER, szGUID.c_str());
				break;
			default:
				break;
			}

			for (auto entry : m_ContactAddressBookObject.lpEntryID)
			{
				szEntryId += entry.ToString();
			}
		}
		else if (eidtWAB == m_ObjectType)
		{
			szEntryId += strings::formatmessage(IDS_ENTRYIDWRAPPEDENTRYIDDATA, m_WAB.Type, InterpretFlags(flagWABEntryIDType, m_WAB.Type).c_str());

			for (auto& entry : m_WAB.lpEntryID)
			{
				szEntryId += entry.ToString();
			}
		}
		else if (eidtMessageDatabase == m_ObjectType)
		{
			auto szVersion = InterpretFlags(flagMDBVersion, m_MessageDatabaseObject.Version);
			auto szFlag = InterpretFlags(flagMDBFlag, m_MessageDatabaseObject.Flag);

			szEntryId += strings::formatmessage(IDS_ENTRYIDMAPIMESSAGESTOREDATA,
				m_MessageDatabaseObject.Version, szVersion.c_str(),
				m_MessageDatabaseObject.Flag, szFlag.c_str(),
				m_MessageDatabaseObject.DLLFileName.c_str());
			if (m_MessageDatabaseObject.bIsExchange)
			{
				auto szWrappedType = InterpretNumberAsStringProp(m_MessageDatabaseObject.WrappedType, PR_PROFILE_OPEN_FLAGS);

				szGUID = guid::GUIDToStringAndName(&m_MessageDatabaseObject.WrappedProviderUID);
				szEntryId += strings::formatmessage(IDS_ENTRYIDMAPIMESSAGESTOREEXCHANGEDATA,
					m_MessageDatabaseObject.WrappedFlags,
					szGUID.c_str(),
					m_MessageDatabaseObject.WrappedType, szWrappedType.c_str(),
					m_MessageDatabaseObject.ServerShortname.c_str(),
					m_MessageDatabaseObject.MailboxDN.c_str());
			}

			switch (m_MessageDatabaseObject.MagicVersion)
			{
			case MDB_STORE_EID_V2_MAGIC:
			{
				auto szV2Magic = InterpretFlags(flagEidMagic, m_MessageDatabaseObject.v2.ulMagic);
				auto szV2Version = InterpretFlags(flagEidVersion, m_MessageDatabaseObject.v2.ulVersion);

				szEntryId += strings::formatmessage(IDS_ENTRYIDMAPIMESSAGESTOREEXCHANGEDATAV2,
					m_MessageDatabaseObject.v2.ulMagic, szV2Magic.c_str(),
					m_MessageDatabaseObject.v2.ulSize,
					m_MessageDatabaseObject.v2.ulVersion, szV2Version.c_str(),
					m_MessageDatabaseObject.v2.ulOffsetDN,
					m_MessageDatabaseObject.v2.ulOffsetFQDN,
					m_MessageDatabaseObject.v2DN.c_str(),
					m_MessageDatabaseObject.v2FQDN.c_str());

				szEntryId += strings::BinToHexString(m_MessageDatabaseObject.v2Reserved, true);
			}
			break;
			case MDB_STORE_EID_V3_MAGIC:
			{
				auto szV3Magic = InterpretFlags(flagEidMagic, m_MessageDatabaseObject.v3.ulMagic);
				auto szV3Version = InterpretFlags(flagEidVersion, m_MessageDatabaseObject.v3.ulVersion);

				szEntryId += strings::formatmessage(IDS_ENTRYIDMAPIMESSAGESTOREEXCHANGEDATAV3,
					m_MessageDatabaseObject.v3.ulMagic, szV3Magic.c_str(),
					m_MessageDatabaseObject.v3.ulSize,
					m_MessageDatabaseObject.v3.ulVersion, szV3Version.c_str(),
					m_MessageDatabaseObject.v3.ulOffsetSmtpAddress,
					m_MessageDatabaseObject.v3SmtpAddress.c_str());

				szEntryId += strings::BinToHexString(m_MessageDatabaseObject.v2Reserved, true);
			}
			break;
			}
		}
		else if (eidtFolder == m_ObjectType)
		{
			auto szType = InterpretFlags(flagMessageDatabaseObjectType, m_FolderOrMessage.Type);

			auto szDatabaseGUID = guid::GUIDToStringAndName(&m_FolderOrMessage.FolderObject.DatabaseGUID);

			szEntryId += strings::formatmessage(IDS_ENTRYIDEXCHANGEFOLDERDATA,
				m_FolderOrMessage.Type, szType.c_str(),
				szDatabaseGUID.c_str());
			szEntryId += strings::BinToHexString(m_FolderOrMessage.FolderObject.GlobalCounter, true);

			szEntryId += strings::formatmessage(IDS_ENTRYIDEXCHANGEDATAPAD);
			szEntryId += strings::BinToHexString(m_FolderOrMessage.FolderObject.Pad, true);
		}
		else if (eidtMessage == m_ObjectType)
		{
			auto szType = InterpretFlags(flagMessageDatabaseObjectType, m_FolderOrMessage.Type);
			auto szFolderDatabaseGUID = guid::GUIDToStringAndName(&m_FolderOrMessage.MessageObject.FolderDatabaseGUID);
			auto szMessageDatabaseGUID = guid::GUIDToStringAndName(&m_FolderOrMessage.MessageObject.MessageDatabaseGUID);

			szEntryId += strings::formatmessage(IDS_ENTRYIDEXCHANGEMESSAGEDATA,
				m_FolderOrMessage.Type, szType.c_str(),
				szFolderDatabaseGUID.c_str());
			szEntryId += strings::BinToHexString(m_FolderOrMessage.MessageObject.FolderGlobalCounter, true);

			szEntryId += strings::formatmessage(IDS_ENTRYIDEXCHANGEDATAPADNUM, 1);
			szEntryId += strings::BinToHexString(m_FolderOrMessage.MessageObject.Pad1, true);

			szEntryId += strings::formatmessage(IDS_ENTRYIDEXCHANGEMESSAGEDATAGUID,
				szMessageDatabaseGUID.c_str());
			szEntryId += strings::BinToHexString(m_FolderOrMessage.MessageObject.MessageGlobalCounter, true);

			szEntryId += strings::formatmessage(IDS_ENTRYIDEXCHANGEDATAPADNUM, 2);
			szEntryId += strings::BinToHexString(m_FolderOrMessage.MessageObject.Pad2, true);
		}

		return szEntryId;
	}
}