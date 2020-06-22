#include <core/stdafx.h>
#include <core/smartview/EntryIdStruct.h>
#include <core/smartview/SmartView.h>
#include <core/interpret/guid.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	void EntryIdStruct::parse()
	{
		m_ObjectType = EIDStructType::unknown;
		if (parser->getSize() < 4) return;
		m_abFlags0 = blockT<byte>::parse(parser);
		m_abFlags1 = blockT<byte>::parse(parser);
		m_abFlags23 = blockBytes::parse(parser, 2);
		m_ProviderUID = blockT<GUID>::parse(parser);

		// Ephemeral entry ID:
		if (*m_abFlags0 == EPHEMERAL)
		{
			m_ObjectType = EIDStructType::ephemeral;
		}
		// One Off Recipient
		else if (*m_ProviderUID == guid::muidOOP)
		{
			m_ObjectType = EIDStructType::oneOff;
		}
		// Address Book Recipient
		else if (*m_ProviderUID == guid::muidEMSAB)
		{
			m_ObjectType = EIDStructType::addressBook;
		}
		// Contact Address Book / Personal Distribution List (PDL)
		else if (*m_ProviderUID == guid::muidContabDLL)
		{
			m_ObjectType = EIDStructType::contact;
		}
		// message store objects
		else if (*m_ProviderUID == guid::muidStoreWrap)
		{
			m_ObjectType = EIDStructType::messageDatabase;
		}
		else if (*m_ProviderUID == guid::WAB_GUID)
		{
			m_ObjectType = EIDStructType::WAB;
		}
		// We can recognize Exchange message store folder and message entry IDs by their size
		else
		{
			const auto ulRemainingBytes = parser->getSize();

			if (sizeof(WORD) + sizeof(GUID) + 6 * sizeof(BYTE) + 2 * sizeof(BYTE) == ulRemainingBytes)
			{
				m_ObjectType = EIDStructType::folder;
			}
			else if (sizeof(WORD) + 2 * sizeof(GUID) + 12 * sizeof(BYTE) + 4 * sizeof(BYTE) == ulRemainingBytes)
			{
				m_ObjectType = EIDStructType::message;
			}
		}

		if (EIDStructType::unknown != m_ObjectType)
		{
			switch (m_ObjectType)
			{
				// Ephemeral Recipient
			case EIDStructType::ephemeral:
				m_EphemeralObject.Version = blockT<DWORD>::parse(parser);
				m_EphemeralObject.Type = blockT<DWORD>::parse(parser);
				break;
				// One Off Recipient
			case EIDStructType::oneOff:
				m_OneOffRecipientObject.Bitmask = blockT<DWORD>::parse(parser);
				if (*m_OneOffRecipientObject.Bitmask & MAPI_UNICODE)
				{
					m_OneOffRecipientObject.Unicode.DisplayName = blockStringW::parse(parser);
					m_OneOffRecipientObject.Unicode.AddressType = blockStringW::parse(parser);
					m_OneOffRecipientObject.Unicode.EmailAddress = blockStringW::parse(parser);
				}
				else
				{
					m_OneOffRecipientObject.ANSI.DisplayName = blockStringA::parse(parser);
					m_OneOffRecipientObject.ANSI.AddressType = blockStringA::parse(parser);
					m_OneOffRecipientObject.ANSI.EmailAddress = blockStringA::parse(parser);
				}
				break;
				// Address Book Recipient
			case EIDStructType::addressBook:
				m_AddressBookObject.Version = blockT<DWORD>::parse(parser);
				m_AddressBookObject.Type = blockT<DWORD>::parse(parser);
				m_AddressBookObject.X500DN = blockStringA::parse(parser);
				break;
				// Contact Address Book / Personal Distribution List (PDL)
			case EIDStructType::contact:
			{
				m_ContactAddressBookObject.Version = blockT<DWORD>::parse(parser);
				m_ContactAddressBookObject.Type = blockT<DWORD>::parse(parser);

				if (*m_ContactAddressBookObject.Type == CONTAB_CONTAINER)
				{
					m_ContactAddressBookObject.muidID = blockT<GUID>::parse(parser);
				}
				else // Assume we've got some variation on the user/distlist format
				{
					m_ContactAddressBookObject.Index = blockT<DWORD>::parse(parser);
					m_ContactAddressBookObject.EntryIDCount = blockT<DWORD>::parse(parser);
				}

				// Read the wrapped entry ID from the remaining data
				auto cbRemainingBytes = parser->getSize();

				// If we already got a size, use it, else we just read the rest of the structure
				if (m_ContactAddressBookObject.EntryIDCount != 0 &&
					*m_ContactAddressBookObject.EntryIDCount < cbRemainingBytes)
				{
					cbRemainingBytes = *m_ContactAddressBookObject.EntryIDCount;
				}

				m_ContactAddressBookObject.lpEntryID = block::parse<EntryIdStruct>(parser, cbRemainingBytes, false);
			}
			break;
			case EIDStructType::WAB:
			{
				m_ObjectType = EIDStructType::WAB;

				m_WAB.Type = blockT<byte>::parse(parser);
				m_WAB.lpEntryID = block::parse<EntryIdStruct>(parser, false);
			}
			break;
			// message store objects
			case EIDStructType::messageDatabase:
				m_MessageDatabaseObject.Version = blockT<byte>::parse(parser);
				m_MessageDatabaseObject.Flag = blockT<byte>::parse(parser);
				m_MessageDatabaseObject.DLLFileName = blockStringA::parse(parser);
				m_MessageDatabaseObject.bIsExchange = false;

				// We only know how to parse emsmdb.dll's wrapped entry IDs
				if (!m_MessageDatabaseObject.DLLFileName->empty() &&
					CSTR_EQUAL == CompareStringA(
									  LOCALE_INVARIANT,
									  NORM_IGNORECASE,
									  m_MessageDatabaseObject.DLLFileName->c_str(),
									  -1,
									  "emsmdb.dll", // STRING_OK
									  -1))
				{
					m_MessageDatabaseObject.bIsExchange = true;
					auto cbRead = parser->getOffset();
					// Advance to the next multiple of 4
					parser->advance(3 - (cbRead + 3) % 4);
					m_MessageDatabaseObject.WrappedFlags = blockT<DWORD>::parse(parser);
					m_MessageDatabaseObject.WrappedProviderUID = blockT<GUID>::parse(parser);
					m_MessageDatabaseObject.WrappedType = blockT<DWORD>::parse(parser);
					m_MessageDatabaseObject.ServerShortname = blockStringA::parse(parser);

					// Test if we have a magic value. Some PF EIDs also have a mailbox DN and we need to accomodate them
					if (*m_MessageDatabaseObject.WrappedType & OPENSTORE_PUBLIC)
					{
						cbRead = parser->getOffset();
						m_MessageDatabaseObject.MagicVersion = blockT<DWORD>::parse(parser);
						parser->setOffset(cbRead);
					}
					else
					{
						m_MessageDatabaseObject.MagicVersion->setOffset(0);
						m_MessageDatabaseObject.MagicVersion->setSize(0);
						m_MessageDatabaseObject.MagicVersion->setData(MDB_STORE_EID_V1_VERSION);
					}

					// Either we're not a PF eid, or this PF EID wasn't followed directly by a magic value
					if (!(*m_MessageDatabaseObject.WrappedType & OPENSTORE_PUBLIC) ||
						*m_MessageDatabaseObject.MagicVersion != MDB_STORE_EID_V2_MAGIC &&
							*m_MessageDatabaseObject.MagicVersion != MDB_STORE_EID_V3_MAGIC)
					{
						m_MessageDatabaseObject.MailboxDN = blockStringA::parse(parser);
					}

					// Check again for a magic value
					cbRead = parser->getOffset();
					m_MessageDatabaseObject.MagicVersion = blockT<DWORD>::parse(parser);
					parser->setOffset(cbRead);

					switch (*m_MessageDatabaseObject.MagicVersion)
					{
					case MDB_STORE_EID_V2_MAGIC:
						if (parser->getSize() >= MDB_STORE_EID_V2::size + sizeof(WCHAR))
						{
							m_MessageDatabaseObject.v2.ulMagic = blockT<DWORD>::parse(parser);
							m_MessageDatabaseObject.v2.ulSize = blockT<DWORD>::parse(parser);
							m_MessageDatabaseObject.v2.ulVersion = blockT<DWORD>::parse(parser);
							m_MessageDatabaseObject.v2.ulOffsetDN = blockT<DWORD>::parse(parser);
							m_MessageDatabaseObject.v2.ulOffsetFQDN = blockT<DWORD>::parse(parser);
							if (*m_MessageDatabaseObject.v2.ulOffsetDN)
							{
								m_MessageDatabaseObject.v2DN = blockStringA::parse(parser);
							}

							if (*m_MessageDatabaseObject.v2.ulOffsetFQDN)
							{
								m_MessageDatabaseObject.v2FQDN = blockStringW::parse(parser);
							}

							m_MessageDatabaseObject.v2Reserved = blockBytes::parse(parser, 2);
						}
						break;
					case MDB_STORE_EID_V3_MAGIC:
						if (parser->getSize() >= MDB_STORE_EID_V3::size + sizeof(WCHAR))
						{
							m_MessageDatabaseObject.v3.ulMagic = blockT<DWORD>::parse(parser);
							m_MessageDatabaseObject.v3.ulSize = blockT<DWORD>::parse(parser);
							m_MessageDatabaseObject.v3.ulVersion = blockT<DWORD>::parse(parser);
							m_MessageDatabaseObject.v3.ulOffsetSmtpAddress = blockT<DWORD>::parse(parser);
							if (m_MessageDatabaseObject.v3.ulOffsetSmtpAddress)
							{
								m_MessageDatabaseObject.v3SmtpAddress = blockStringW::parse(parser);
							}

							m_MessageDatabaseObject.v2Reserved = blockBytes::parse(parser, 2);
						}
						break;
					}
				}
				break;
				// Exchange message store folder
			case EIDStructType::folder:
				m_FolderOrMessage.Type = blockT<WORD>::parse(parser);
				m_FolderOrMessage.FolderObject.DatabaseGUID = blockT<GUID>::parse(parser);
				m_FolderOrMessage.FolderObject.GlobalCounter = blockBytes::parse(parser, 6);
				m_FolderOrMessage.FolderObject.Pad = blockBytes::parse(parser, 2);
				break;
				// Exchange message store message
			case EIDStructType::message:
				m_FolderOrMessage.Type = blockT<WORD>::parse(parser);
				m_FolderOrMessage.MessageObject.FolderDatabaseGUID = blockT<GUID>::parse(parser);
				m_FolderOrMessage.MessageObject.FolderGlobalCounter = blockBytes::parse(parser, 6);
				m_FolderOrMessage.MessageObject.Pad1 = blockBytes::parse(parser, 2);

				m_FolderOrMessage.MessageObject.MessageDatabaseGUID = blockT<GUID>::parse(parser);
				m_FolderOrMessage.MessageObject.MessageGlobalCounter = blockBytes::parse(parser, 6);
				m_FolderOrMessage.MessageObject.Pad2 = blockBytes::parse(parser, 2);
				break;
			}
		}

		// Check if we have an unidentified short term entry ID:
		if (EIDStructType::unknown == m_ObjectType && *m_abFlags0 & MAPI_SHORTTERM)
			m_ObjectType = EIDStructType::shortTerm;
	}

	void EntryIdStruct::parseBlocks()
	{
		switch (m_ObjectType)
		{
		case EIDStructType::unknown:
			setText(L"Entry ID:");
			break;
		case EIDStructType::ephemeral:
			setText(L"Ephemeral Entry ID:");
			break;
		case EIDStructType::shortTerm:
			setText(L"Short Term Entry ID:");
			break;
		case EIDStructType::folder:
			setText(L"Exchange Folder Entry ID:");
			break;
		case EIDStructType::message:
			setText(L"Exchange Message Entry ID:");
			break;
		case EIDStructType::messageDatabase:
			setText(L"MAPI Message Store Entry ID:");
			break;
		case EIDStructType::oneOff:
			setText(L"MAPI One Off Recipient Entry ID:");
			break;
		case EIDStructType::addressBook:
			setText(L"Exchange Address Entry ID:");
			break;
		case EIDStructType::contact:
			setText(L"Contact Address Book / PDL Entry ID:");
			break;
		case EIDStructType::WAB:
			setText(L"Wrapped Entry ID:");
			break;
		}

		if (m_abFlags23->empty()) return;
		if (0 == (*m_abFlags0 | *m_abFlags1 | m_abFlags23->data()[0] | m_abFlags23->data()[1]))
		{
			addChild(create(4, m_abFlags0->getOffset(), L"abFlags = 0x00000000"));
		}
		else if (0 == (*m_abFlags1 | m_abFlags23->data()[0] | m_abFlags23->data()[1]))
		{
			addChild(
				m_abFlags0,
				L"abFlags[0] = 0x%1!02X!= %2!ws!",
				m_abFlags0->getData(),
				flags::InterpretFlags(flagEntryId0, *m_abFlags0).c_str());
			addChild(create(3, m_abFlags1->getOffset(), L"abFlags[1..3] = 0x000000"));
		}
		else
		{
			addChild(
				m_abFlags0,
				L"abFlags[0] = 0x%1!02X!= %2!ws!",
				m_abFlags0->getData(),
				flags::InterpretFlags(flagEntryId0, *m_abFlags0).c_str());
			addChild(
				m_abFlags1,
				L"abFlags[1] = 0x%1!02X!= %2!ws!",
				m_abFlags1->getData(),
				flags::InterpretFlags(flagEntryId1, *m_abFlags1).c_str());
			addChild(m_abFlags23, L"abFlags[2..3] = 0x%1!02X!%2!02X!", m_abFlags23->data()[0], m_abFlags23->data()[1]);
		}

		addChild(m_ProviderUID, L"Provider GUID = %1!ws!", guid::GUIDToStringAndName(*m_ProviderUID).c_str());

		if (EIDStructType::ephemeral == m_ObjectType)
		{
			auto szVersion = flags::InterpretFlags(flagExchangeABVersion, *m_EphemeralObject.Version);
			addChild(
				m_EphemeralObject.Version,
				L"Version = 0x%1!08X! = %2!ws!",
				m_EphemeralObject.Version->getData(),
				szVersion.c_str());

			auto szType = InterpretNumberAsStringProp(*m_EphemeralObject.Type, PR_DISPLAY_TYPE);
			addChild(
				m_EphemeralObject.Type,
				L"Type = 0x%1!08X! = %2!ws!",
				m_EphemeralObject.Type->getData(),
				szType.c_str());
		}
		else if (EIDStructType::oneOff == m_ObjectType)
		{
			auto szFlags = flags::InterpretFlags(flagOneOffEntryId, *m_OneOffRecipientObject.Bitmask);
			if (*m_OneOffRecipientObject.Bitmask & MAPI_UNICODE)
			{
				addChild(
					m_OneOffRecipientObject.Bitmask,
					L"dwBitmask: 0x%1!08X! = %2!ws!",
					m_OneOffRecipientObject.Bitmask->getData(),
					szFlags.c_str());
				addChild(
					m_OneOffRecipientObject.Unicode.DisplayName,
					L"szDisplayName = %1!ws!",
					m_OneOffRecipientObject.Unicode.DisplayName->c_str());
				addChild(
					m_OneOffRecipientObject.Unicode.AddressType,
					L"szAddressType = %1!ws!",
					m_OneOffRecipientObject.Unicode.AddressType->c_str());
				addChild(
					m_OneOffRecipientObject.Unicode.EmailAddress,
					L"szEmailAddress = %1!ws!",
					m_OneOffRecipientObject.Unicode.EmailAddress->c_str());
			}
			else
			{
				addChild(
					m_OneOffRecipientObject.Bitmask,
					L"dwBitmask: 0x%1!08X! = %2!ws!",
					m_OneOffRecipientObject.Bitmask->getData(),
					szFlags.c_str());
				addChild(
					m_OneOffRecipientObject.ANSI.DisplayName,
					L"szDisplayName = %1!hs!",
					m_OneOffRecipientObject.ANSI.DisplayName->c_str());
				addChild(
					m_OneOffRecipientObject.ANSI.AddressType,
					L"szAddressType = %1!hs!",
					m_OneOffRecipientObject.ANSI.AddressType->c_str());
				addChild(
					m_OneOffRecipientObject.ANSI.EmailAddress,
					L"szEmailAddress = %1!hs!",
					m_OneOffRecipientObject.ANSI.EmailAddress->c_str());
			}
		}
		else if (EIDStructType::addressBook == m_ObjectType)
		{
			auto szVersion = flags::InterpretFlags(flagExchangeABVersion, *m_AddressBookObject.Version);
			addChild(
				m_AddressBookObject.Version,
				L"Version = 0x%1!08X! = %2!ws!",
				m_AddressBookObject.Version->getData(),
				szVersion.c_str());
			auto szType = InterpretNumberAsStringProp(*m_AddressBookObject.Type, PR_DISPLAY_TYPE);
			addChild(
				m_AddressBookObject.Type,
				L"Type = 0x%1!08X! = %2!ws!",
				m_AddressBookObject.Type->getData(),
				szType.c_str());
			addChild(m_AddressBookObject.X500DN, L"X500DN = %1!hs!", m_AddressBookObject.X500DN->c_str());
		}
		// Contact Address Book / Personal Distribution List (PDL)
		else if (EIDStructType::contact == m_ObjectType)
		{
			auto szVersion = flags::InterpretFlags(flagContabVersion, *m_ContactAddressBookObject.Version);
			addChild(
				m_ContactAddressBookObject.Version,
				L"Version = 0x%1!08X! = %2!ws!",
				m_ContactAddressBookObject.Version->getData(),
				szVersion.c_str());

			auto szType = flags::InterpretFlags(flagContabType, *m_ContactAddressBookObject.Type);
			addChild(
				m_ContactAddressBookObject.Type,
				L"Type = 0x%1!08X! = %2!ws!",
				m_ContactAddressBookObject.Type->getData(),
				szType.c_str());

			switch (*m_ContactAddressBookObject.Type)
			{
			case CONTAB_USER:
			case CONTAB_DISTLIST:
				addChild(
					m_ContactAddressBookObject.Index,
					L"Index = 0x%1!08X! = %2!ws!",
					m_ContactAddressBookObject.Index->getData(),
					flags::InterpretFlags(flagContabIndex, *m_ContactAddressBookObject.Index).c_str());
				addChild(
					m_ContactAddressBookObject.EntryIDCount,
					L"EntryIDCount = 0x%1!08X!",
					m_ContactAddressBookObject.EntryIDCount->getData());
				break;
			case CONTAB_ROOT:
			case CONTAB_CONTAINER:
			case CONTAB_SUBROOT:
				addChild(
					m_ContactAddressBookObject.muidID,
					L"muidID = %1!ws!",
					guid::GUIDToStringAndName(*m_ContactAddressBookObject.muidID).c_str());
				break;
			default:
				break;
			}

			addChild(m_ContactAddressBookObject.lpEntryID);
		}
		else if (EIDStructType::WAB == m_ObjectType)
		{
			addChild(
				m_WAB.Type,
				L"Wrapped Entry Type = 0x%1!02X! = %2!ws!",
				m_WAB.Type->getData(),
				flags::InterpretFlags(flagWABEntryIDType, *m_WAB.Type).c_str());

			addChild(m_WAB.lpEntryID);
		}
		else if (EIDStructType::messageDatabase == m_ObjectType)
		{
			auto szVersion = flags::InterpretFlags(flagMDBVersion, *m_MessageDatabaseObject.Version);
			addChild(
				m_MessageDatabaseObject.Version,
				L"Version = 0x%1!02X! = %2!ws!",
				m_MessageDatabaseObject.Version->getData(),
				szVersion.c_str());

			auto szFlag = flags::InterpretFlags(flagMDBFlag, *m_MessageDatabaseObject.Flag);
			addChild(
				m_MessageDatabaseObject.Flag,
				L"Flag = 0x%1!02X! = %2!ws!",
				m_MessageDatabaseObject.Flag->getData(),
				szFlag.c_str());

			addChild(
				m_MessageDatabaseObject.DLLFileName,
				L"DLLFileName = %1!hs!",
				m_MessageDatabaseObject.DLLFileName->c_str());
			if (m_MessageDatabaseObject.bIsExchange)
			{
				auto szWrappedType =
					InterpretNumberAsStringProp(*m_MessageDatabaseObject.WrappedType, PR_PROFILE_OPEN_FLAGS);
				addChild(
					m_MessageDatabaseObject.WrappedFlags,
					L"Wrapped Flags = 0x%1!08X!",
					m_MessageDatabaseObject.WrappedFlags->getData());
				addChild(
					m_MessageDatabaseObject.WrappedProviderUID,
					L"WrappedProviderUID = %1!ws!",
					guid::GUIDToStringAndName(*m_MessageDatabaseObject.WrappedProviderUID).c_str());
				addChild(
					m_MessageDatabaseObject.WrappedType,
					L"WrappedType = 0x%1!08X! = %2!ws!",
					m_MessageDatabaseObject.WrappedType->getData(),
					szWrappedType.c_str());
				addChild(
					m_MessageDatabaseObject.ServerShortname,
					L"ServerShortname = %1!hs!",
					m_MessageDatabaseObject.ServerShortname->c_str());
				addChild(
					m_MessageDatabaseObject.MailboxDN,
					L"MailboxDN = %1!hs!",
					m_MessageDatabaseObject.MailboxDN->c_str());
			}

			switch (*m_MessageDatabaseObject.MagicVersion)
			{
			case MDB_STORE_EID_V2_MAGIC:
			{
				auto szV2Magic = flags::InterpretFlags(flagEidMagic, *m_MessageDatabaseObject.v2.ulMagic);
				addChild(
					m_MessageDatabaseObject.v2.ulMagic,
					L"Magic = 0x%1!08X! = %2!ws!",
					m_MessageDatabaseObject.v2.ulMagic->getData(),
					szV2Magic.c_str());
				addChild(
					m_MessageDatabaseObject.v2.ulSize,
					L"Size = 0x%1!08X! = %1!d!",
					m_MessageDatabaseObject.v2.ulSize->getData());
				auto szV2Version = flags::InterpretFlags(flagEidVersion, *m_MessageDatabaseObject.v2.ulVersion);
				addChild(
					m_MessageDatabaseObject.v2.ulVersion,
					L"Version = 0x%1!08X! = %2!ws!",
					m_MessageDatabaseObject.v2.ulVersion->getData(),
					szV2Version.c_str());
				addChild(
					m_MessageDatabaseObject.v2.ulOffsetDN,
					L"OffsetDN = 0x%1!08X!",
					m_MessageDatabaseObject.v2.ulOffsetDN->getData());
				addChild(
					m_MessageDatabaseObject.v2.ulOffsetFQDN,
					L"OffsetFQDN = 0x%1!08X!",
					m_MessageDatabaseObject.v2.ulOffsetFQDN->getData());
				addChild(m_MessageDatabaseObject.v2DN, L"DN = %1!hs!", m_MessageDatabaseObject.v2DN->c_str());
				addChild(m_MessageDatabaseObject.v2FQDN, L"FQDN = %1!ws!", m_MessageDatabaseObject.v2FQDN->c_str());

				addLabeledChild(L"Reserved Bytes =", m_MessageDatabaseObject.v2Reserved);
			}
			break;
			case MDB_STORE_EID_V3_MAGIC:
			{
				auto szV3Magic = flags::InterpretFlags(flagEidMagic, *m_MessageDatabaseObject.v3.ulMagic);
				addChild(
					m_MessageDatabaseObject.v3.ulMagic,
					L"Magic = 0x%1!08X! = %2!ws!",
					m_MessageDatabaseObject.v3.ulMagic->getData(),
					szV3Magic.c_str());
				addChild(
					m_MessageDatabaseObject.v3.ulSize,
					L"Size = 0x%1!08X! = %1!d!",
					m_MessageDatabaseObject.v3.ulSize->getData());
				auto szV3Version = flags::InterpretFlags(flagEidVersion, *m_MessageDatabaseObject.v3.ulVersion);
				addChild(
					m_MessageDatabaseObject.v3.ulVersion,
					L"Version = 0x%1!08X! = %2!ws!",
					m_MessageDatabaseObject.v3.ulVersion->getData(),
					szV3Version.c_str());
				addChild(
					m_MessageDatabaseObject.v3.ulOffsetSmtpAddress,
					L"OffsetSmtpAddress = 0x%1!08X!",
					m_MessageDatabaseObject.v3.ulOffsetSmtpAddress->getData());
				addChild(
					m_MessageDatabaseObject.v3SmtpAddress,
					L"SmtpAddress = %1!ws!",
					m_MessageDatabaseObject.v3SmtpAddress->c_str());

				addLabeledChild(L"Reserved Bytes =", m_MessageDatabaseObject.v2Reserved);
			}
			break;
			}
		}
		else if (EIDStructType::folder == m_ObjectType)
		{
			auto szType = flags::InterpretFlags(flagMessageDatabaseObjectType, *m_FolderOrMessage.Type);
			addChild(
				m_FolderOrMessage.Type,
				L"Folder Type = 0x%1!04X! = %2!ws!",
				m_FolderOrMessage.Type->getData(),
				szType.c_str());
			addChild(
				m_FolderOrMessage.FolderObject.DatabaseGUID,
				L"Database GUID = %1!ws!",
				guid::GUIDToStringAndName(*m_FolderOrMessage.FolderObject.DatabaseGUID).c_str());
			addLabeledChild(L"GlobalCounter =", m_FolderOrMessage.FolderObject.GlobalCounter);

			addLabeledChild(L"Pad =", m_FolderOrMessage.FolderObject.Pad);
		}
		else if (EIDStructType::message == m_ObjectType)
		{
			auto szType = flags::InterpretFlags(flagMessageDatabaseObjectType, *m_FolderOrMessage.Type);
			addChild(
				m_FolderOrMessage.Type,
				L"Message Type = 0x%1!04X! = %2!ws!",
				m_FolderOrMessage.Type->getData(),
				szType.c_str());
			addChild(
				m_FolderOrMessage.MessageObject.FolderDatabaseGUID,
				L"Folder Database GUID = %1!ws!",
				guid::GUIDToStringAndName(*m_FolderOrMessage.MessageObject.FolderDatabaseGUID).c_str());
			addLabeledChild(L"Folder GlobalCounter =", m_FolderOrMessage.MessageObject.FolderGlobalCounter);

			addLabeledChild(L"Pad1 =", m_FolderOrMessage.MessageObject.Pad1);

			addChild(
				m_FolderOrMessage.MessageObject.MessageDatabaseGUID,
				L"Message Database GUID = %1!ws!",
				guid::GUIDToStringAndName(*m_FolderOrMessage.MessageObject.MessageDatabaseGUID).c_str());
			addLabeledChild(L"Message GlobalCounter =", m_FolderOrMessage.MessageObject.MessageGlobalCounter);

			addLabeledChild(L"Pad2 =", m_FolderOrMessage.MessageObject.Pad2);
		}
	}
} // namespace smartview