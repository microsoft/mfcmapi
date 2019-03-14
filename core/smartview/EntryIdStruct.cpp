#include <core/stdafx.h>
#include <core/smartview/EntryIdStruct.h>
#include <core/smartview/SmartView.h>
#include <core/interpret/guid.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	void EntryIdStruct::Parse()
	{
		m_ObjectType = eidtUnknown;
		if (m_Parser->RemainingBytes() < 4) return;
		m_abFlags0.parse<byte>(m_Parser);
		m_abFlags1.parse<byte>(m_Parser);
		m_abFlags23.parse(m_Parser, 2);
		m_ProviderUID.parse<GUID>(m_Parser);

		// Ephemeral entry ID:
		if (m_abFlags0 == EPHEMERAL)
		{
			m_ObjectType = eidtEphemeral;
		}
		// One Off Recipient
		else if (m_ProviderUID == guid::muidOOP)
		{
			m_ObjectType = eidtOneOff;
		}
		// Address Book Recipient
		else if (m_ProviderUID == guid::muidEMSAB)
		{
			m_ObjectType = eidtAddressBook;
		}
		// Contact Address Book / Personal Distribution List (PDL)
		else if (m_ProviderUID == guid::muidContabDLL)
		{
			m_ObjectType = eidtContact;
		}
		// message store objects
		else if (m_ProviderUID == guid::muidStoreWrap)
		{
			m_ObjectType = eidtMessageDatabase;
		}
		else if (m_ProviderUID == guid::WAB_GUID)
		{
			m_ObjectType = eidtWAB;
		}
		// We can recognize Exchange message store folder and message entry IDs by their size
		else
		{
			const auto ulRemainingBytes = m_Parser->RemainingBytes();

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
				m_EphemeralObject.Version.parse<DWORD>(m_Parser);
				m_EphemeralObject.Type.parse<DWORD>(m_Parser);
				break;
				// One Off Recipient
			case eidtOneOff:
				m_OneOffRecipientObject.Bitmask.parse<DWORD>(m_Parser);
				if (MAPI_UNICODE & m_OneOffRecipientObject.Bitmask)
				{
					m_OneOffRecipientObject.Unicode.DisplayName.parse(m_Parser);
					m_OneOffRecipientObject.Unicode.AddressType.parse(m_Parser);
					m_OneOffRecipientObject.Unicode.EmailAddress.parse(m_Parser);
				}
				else
				{
					m_OneOffRecipientObject.ANSI.DisplayName.parse(m_Parser);
					m_OneOffRecipientObject.ANSI.AddressType.parse(m_Parser);
					m_OneOffRecipientObject.ANSI.EmailAddress.parse(m_Parser);
				}
				break;
				// Address Book Recipient
			case eidtAddressBook:
				m_AddressBookObject.Version.parse<DWORD>(m_Parser);
				m_AddressBookObject.Type.parse<DWORD>(m_Parser);
				m_AddressBookObject.X500DN.parse(m_Parser);
				break;
				// Contact Address Book / Personal Distribution List (PDL)
			case eidtContact:
			{
				m_ContactAddressBookObject.Version.parse<DWORD>(m_Parser);
				m_ContactAddressBookObject.Type.parse<DWORD>(m_Parser);

				if (CONTAB_CONTAINER == m_ContactAddressBookObject.Type)
				{
					m_ContactAddressBookObject.muidID.parse<GUID>(m_Parser);
				}
				else // Assume we've got some variation on the user/distlist format
				{
					m_ContactAddressBookObject.Index.parse<DWORD>(m_Parser);
					m_ContactAddressBookObject.EntryIDCount.parse<DWORD>(m_Parser);
				}

				// Read the wrapped entry ID from the remaining data
				auto cbRemainingBytes = m_Parser->RemainingBytes();

				// If we already got a size, use it, else we just read the rest of the structure
				if (m_ContactAddressBookObject.EntryIDCount != 0 &&
					m_ContactAddressBookObject.EntryIDCount < cbRemainingBytes)
				{
					cbRemainingBytes = m_ContactAddressBookObject.EntryIDCount;
				}

				m_ContactAddressBookObject.lpEntryID =
					std::make_shared<EntryIdStruct>(m_Parser, cbRemainingBytes, false);
			}
			break;
			case eidtWAB:
			{
				m_ObjectType = eidtWAB;

				m_WAB.Type.parse<BYTE>(m_Parser);
				m_WAB.lpEntryID = std::make_shared<EntryIdStruct>(m_Parser, false);
			}
			break;
			// message store objects
			case eidtMessageDatabase:
				m_MessageDatabaseObject.Version.parse<BYTE>(m_Parser);
				m_MessageDatabaseObject.Flag.parse<BYTE>(m_Parser);
				m_MessageDatabaseObject.DLLFileName.parse(m_Parser);
				m_MessageDatabaseObject.bIsExchange = false;

				// We only know how to parse emsmdb.dll's wrapped entry IDs
				if (!m_MessageDatabaseObject.DLLFileName.empty() &&
					CSTR_EQUAL == CompareStringA(
									  LOCALE_INVARIANT,
									  NORM_IGNORECASE,
									  m_MessageDatabaseObject.DLLFileName.c_str(),
									  -1,
									  "emsmdb.dll", // STRING_OK
									  -1))
				{
					m_MessageDatabaseObject.bIsExchange = true;
					auto cbRead = m_Parser->GetCurrentOffset();
					// Advance to the next multiple of 4
					m_Parser->advance(3 - (cbRead + 3) % 4);
					m_MessageDatabaseObject.WrappedFlags.parse<DWORD>(m_Parser);
					m_MessageDatabaseObject.WrappedProviderUID.parse<GUID>(m_Parser);
					m_MessageDatabaseObject.WrappedType.parse<DWORD>(m_Parser);
					m_MessageDatabaseObject.ServerShortname.parse(m_Parser);

					// Test if we have a magic value. Some PF EIDs also have a mailbox DN and we need to accomodate them
					if (m_MessageDatabaseObject.WrappedType & OPENSTORE_PUBLIC)
					{
						cbRead = m_Parser->GetCurrentOffset();
						m_MessageDatabaseObject.MagicVersion.parse<DWORD>(m_Parser);
						m_Parser->SetCurrentOffset(cbRead);
					}
					else
					{
						m_MessageDatabaseObject.MagicVersion.setOffset(0);
						m_MessageDatabaseObject.MagicVersion.setSize(0);
						m_MessageDatabaseObject.MagicVersion.setData(MDB_STORE_EID_V1_VERSION);
					}

					// Either we're not a PF eid, or this PF EID wasn't followed directly by a magic value
					if (!(m_MessageDatabaseObject.WrappedType & OPENSTORE_PUBLIC) ||
						m_MessageDatabaseObject.MagicVersion != MDB_STORE_EID_V2_MAGIC &&
							m_MessageDatabaseObject.MagicVersion != MDB_STORE_EID_V3_MAGIC)
					{
						m_MessageDatabaseObject.MailboxDN.parse(m_Parser);
					}

					// Check again for a magic value
					cbRead = m_Parser->GetCurrentOffset();
					m_MessageDatabaseObject.MagicVersion.parse<DWORD>(m_Parser);
					m_Parser->SetCurrentOffset(cbRead);

					switch (m_MessageDatabaseObject.MagicVersion)
					{
					case MDB_STORE_EID_V2_MAGIC:
						if (m_Parser->RemainingBytes() >= MDB_STORE_EID_V2::size + sizeof(WCHAR))
						{
							m_MessageDatabaseObject.v2.ulMagic.parse<DWORD>(m_Parser);
							m_MessageDatabaseObject.v2.ulSize.parse<DWORD>(m_Parser);
							m_MessageDatabaseObject.v2.ulVersion.parse<DWORD>(m_Parser);
							m_MessageDatabaseObject.v2.ulOffsetDN.parse<DWORD>(m_Parser);
							m_MessageDatabaseObject.v2.ulOffsetFQDN.parse<DWORD>(m_Parser);
							if (m_MessageDatabaseObject.v2.ulOffsetDN)
							{
								m_MessageDatabaseObject.v2DN.parse(m_Parser);
							}

							if (m_MessageDatabaseObject.v2.ulOffsetFQDN)
							{
								m_MessageDatabaseObject.v2FQDN.parse(m_Parser);
							}

							m_MessageDatabaseObject.v2Reserved.parse(m_Parser, 2);
						}
						break;
					case MDB_STORE_EID_V3_MAGIC:
						if (m_Parser->RemainingBytes() >= MDB_STORE_EID_V3::size + sizeof(WCHAR))
						{
							m_MessageDatabaseObject.v3.ulMagic.parse<DWORD>(m_Parser);
							m_MessageDatabaseObject.v3.ulSize.parse<DWORD>(m_Parser);
							m_MessageDatabaseObject.v3.ulVersion.parse<DWORD>(m_Parser);
							m_MessageDatabaseObject.v3.ulOffsetSmtpAddress.parse<DWORD>(m_Parser);
							if (m_MessageDatabaseObject.v3.ulOffsetSmtpAddress)
							{
								m_MessageDatabaseObject.v3SmtpAddress.parse(m_Parser);
							}

							m_MessageDatabaseObject.v2Reserved.parse(m_Parser, 2);
						}
						break;
					}
				}
				break;
				// Exchange message store folder
			case eidtFolder:
				m_FolderOrMessage.Type.parse<WORD>(m_Parser);
				m_FolderOrMessage.FolderObject.DatabaseGUID.parse<GUID>(m_Parser);
				m_FolderOrMessage.FolderObject.GlobalCounter.parse(m_Parser, 6);
				m_FolderOrMessage.FolderObject.Pad.parse(m_Parser, 2);
				break;
				// Exchange message store message
			case eidtMessage:
				m_FolderOrMessage.Type.parse<WORD>(m_Parser);
				m_FolderOrMessage.MessageObject.FolderDatabaseGUID.parse<GUID>(m_Parser);
				m_FolderOrMessage.MessageObject.FolderGlobalCounter.parse(m_Parser, 6);
				m_FolderOrMessage.MessageObject.Pad1.parse(m_Parser, 2);

				m_FolderOrMessage.MessageObject.MessageDatabaseGUID.parse<GUID>(m_Parser);
				m_FolderOrMessage.MessageObject.MessageGlobalCounter.parse(m_Parser, 6);
				m_FolderOrMessage.MessageObject.Pad2.parse(m_Parser, 2);
				break;
			}
		}

		// Check if we have an unidentified short term entry ID:
		if (eidtUnknown == m_ObjectType && m_abFlags0 & MAPI_SHORTTERM) m_ObjectType = eidtShortTerm;
	}

	void EntryIdStruct::ParseBlocks()
	{
		switch (m_ObjectType)
		{
		case eidtUnknown:
			setRoot(L"Entry ID:\r\n");
			break;
		case eidtEphemeral:
			setRoot(L"Ephemeral Entry ID:\r\n");
			break;
		case eidtShortTerm:
			setRoot(L"Short Term Entry ID:\r\n");
			break;
		case eidtFolder:
			setRoot(L"Exchange Folder Entry ID:\r\n");
			break;
		case eidtMessage:
			setRoot(L"Exchange Message Entry ID:\r\n");
			break;
		case eidtMessageDatabase:
			setRoot(L"MAPI Message Store Entry ID:\r\n");
			break;
		case eidtOneOff:
			setRoot(L"MAPI One Off Recipient Entry ID:\r\n");
			break;
		case eidtAddressBook:
			setRoot(L"Exchange Address Entry ID:\r\n");
			break;
		case eidtContact:
			setRoot(L"Contact Address Book / PDL Entry ID:\r\n");
			break;
		case eidtWAB:
			setRoot(L"Wrapped Entry ID:\r\n");
			break;
		}

		if (m_abFlags23.empty()) return;
		if (0 == (m_abFlags0 | m_abFlags1 | m_abFlags23[0] | m_abFlags23[1]))
		{
			auto tempBlock = std::make_shared<block>();
			tempBlock->setOffset(m_abFlags0.getOffset());
			tempBlock->setSize(4);
			addChild(tempBlock, L"abFlags = 0x00000000\r\n");
		}
		else if (0 == (m_abFlags1 | m_abFlags23[0] | m_abFlags23[1]))
		{
			addChild(
				m_abFlags0,
				L"abFlags[0] = 0x%1!02X!= %2!ws!\r\n",
				m_abFlags0.getData(),
				flags::InterpretFlags(flagEntryId0, m_abFlags0).c_str());
			auto tempBlock = std::make_shared<block>();
			tempBlock->setOffset(m_abFlags1.getOffset());
			tempBlock->setSize(3);
			addChild(tempBlock, L"abFlags[1..3] = 0x000000\r\n");
		}
		else
		{
			addChild(
				m_abFlags0,
				L"abFlags[0] = 0x%1!02X!= %2!ws!\r\n",
				m_abFlags0.getData(),
				flags::InterpretFlags(flagEntryId0, m_abFlags0).c_str());
			addChild(
				m_abFlags1,
				L"abFlags[1] = 0x%1!02X!= %2!ws!\r\n",
				m_abFlags1.getData(),
				flags::InterpretFlags(flagEntryId1, m_abFlags1).c_str());
			addChild(m_abFlags23, L"abFlags[2..3] = 0x%1!02X!%2!02X!\r\n", m_abFlags23[0], m_abFlags23[1]);
		}

		addChild(m_ProviderUID, L"Provider GUID = %1!ws!", guid::GUIDToStringAndName(m_ProviderUID).c_str());

		if (eidtEphemeral == m_ObjectType)
		{
			terminateBlock();

			auto szVersion = flags::InterpretFlags(flagExchangeABVersion, m_EphemeralObject.Version);
			addChild(
				m_EphemeralObject.Version,
				L"Version = 0x%1!08X! = %2!ws!\r\n",
				m_EphemeralObject.Version.getData(),
				szVersion.c_str());

			auto szType = InterpretNumberAsStringProp(m_EphemeralObject.Type, PR_DISPLAY_TYPE);
			addChild(
				m_EphemeralObject.Type, L"Type = 0x%1!08X! = %2!ws!", m_EphemeralObject.Type.getData(), szType.c_str());
		}
		else if (eidtOneOff == m_ObjectType)
		{
			auto szFlags = flags::InterpretFlags(flagOneOffEntryId, m_OneOffRecipientObject.Bitmask);
			if (MAPI_UNICODE & m_OneOffRecipientObject.Bitmask)
			{
				terminateBlock();
				addChild(
					m_OneOffRecipientObject.Bitmask,
					L"dwBitmask: 0x%1!08X! = %2!ws!\r\n",
					m_OneOffRecipientObject.Bitmask.getData(),
					szFlags.c_str());
				addChild(
					m_OneOffRecipientObject.Unicode.DisplayName,
					L"szDisplayName = %1!ws!\r\n",
					m_OneOffRecipientObject.Unicode.DisplayName.c_str());
				addChild(
					m_OneOffRecipientObject.Unicode.AddressType,
					L"szAddressType = %1!ws!\r\n",
					m_OneOffRecipientObject.Unicode.AddressType.c_str());
				addChild(
					m_OneOffRecipientObject.Unicode.EmailAddress,
					L"szEmailAddress = %1!ws!",
					m_OneOffRecipientObject.Unicode.EmailAddress.c_str());
			}
			else
			{
				terminateBlock();
				addChild(
					m_OneOffRecipientObject.Bitmask,
					L"dwBitmask: 0x%1!08X! = %2!ws!\r\n",
					m_OneOffRecipientObject.Bitmask.getData(),
					szFlags.c_str());
				addChild(
					m_OneOffRecipientObject.ANSI.DisplayName,
					L"szDisplayName = %1!hs!\r\n",
					m_OneOffRecipientObject.ANSI.DisplayName.c_str());
				addChild(
					m_OneOffRecipientObject.ANSI.AddressType,
					L"szAddressType = %1!hs!\r\n",
					m_OneOffRecipientObject.ANSI.AddressType.c_str());
				addChild(
					m_OneOffRecipientObject.ANSI.EmailAddress,
					L"szEmailAddress = %1!hs!",
					m_OneOffRecipientObject.ANSI.EmailAddress.c_str());
			}
		}
		else if (eidtAddressBook == m_ObjectType)
		{
			terminateBlock();
			auto szVersion = flags::InterpretFlags(flagExchangeABVersion, m_AddressBookObject.Version);
			addChild(
				m_AddressBookObject.Version,
				L"Version = 0x%1!08X! = %2!ws!\r\n",
				m_AddressBookObject.Version.getData(),
				szVersion.c_str());
			auto szType = InterpretNumberAsStringProp(m_AddressBookObject.Type, PR_DISPLAY_TYPE);
			addChild(
				m_AddressBookObject.Type,
				L"Type = 0x%1!08X! = %2!ws!\r\n",
				m_AddressBookObject.Type.getData(),
				szType.c_str());
			addChild(m_AddressBookObject.X500DN, L"X500DN = %1!hs!", m_AddressBookObject.X500DN.c_str());
		}
		// Contact Address Book / Personal Distribution List (PDL)
		else if (eidtContact == m_ObjectType)
		{
			terminateBlock();
			auto szVersion = flags::InterpretFlags(flagContabVersion, m_ContactAddressBookObject.Version);
			addChild(
				m_ContactAddressBookObject.Version,
				L"Version = 0x%1!08X! = %2!ws!\r\n",
				m_ContactAddressBookObject.Version.getData(),
				szVersion.c_str());

			auto szType = flags::InterpretFlags(flagContabType, m_ContactAddressBookObject.Type);
			addChild(
				m_ContactAddressBookObject.Type,
				L"Type = 0x%1!08X! = %2!ws!\r\n",
				m_ContactAddressBookObject.Type.getData(),
				szType.c_str());

			switch (m_ContactAddressBookObject.Type)
			{
			case CONTAB_USER:
			case CONTAB_DISTLIST:
				addChild(
					m_ContactAddressBookObject.Index,
					L"Index = 0x%1!08X! = %2!ws!\r\n",
					m_ContactAddressBookObject.Index.getData(),
					flags::InterpretFlags(flagContabIndex, m_ContactAddressBookObject.Index).c_str());
				addChild(
					m_ContactAddressBookObject.EntryIDCount,
					L"EntryIDCount = 0x%1!08X!\r\n",
					m_ContactAddressBookObject.EntryIDCount.getData());
				break;
			case CONTAB_ROOT:
			case CONTAB_CONTAINER:
			case CONTAB_SUBROOT:
				addChild(
					m_ContactAddressBookObject.muidID,
					L"muidID = %1!ws!\r\n",
					guid::GUIDToStringAndName(m_ContactAddressBookObject.muidID).c_str());
				break;
			default:
				break;
			}

			addChild(m_ContactAddressBookObject.lpEntryID->getBlock());
		}
		else if (eidtWAB == m_ObjectType)
		{
			terminateBlock();
			addChild(
				m_WAB.Type,
				L"Wrapped Entry Type = 0x%1!02X! = %2!ws!\r\n",
				m_WAB.Type.getData(),
				flags::InterpretFlags(flagWABEntryIDType, m_WAB.Type).c_str());

			addChild(m_WAB.lpEntryID->getBlock());
		}
		else if (eidtMessageDatabase == m_ObjectType)
		{
			terminateBlock();
			auto szVersion = flags::InterpretFlags(flagMDBVersion, m_MessageDatabaseObject.Version);
			addChild(
				m_MessageDatabaseObject.Version,
				L"Version = 0x%1!02X! = %2!ws!\r\n",
				m_MessageDatabaseObject.Version.getData(),
				szVersion.c_str());

			auto szFlag = flags::InterpretFlags(flagMDBFlag, m_MessageDatabaseObject.Flag);
			addChild(
				m_MessageDatabaseObject.Flag,
				L"Flag = 0x%1!02X! = %2!ws!\r\n",
				m_MessageDatabaseObject.Flag.getData(),
				szFlag.c_str());

			addChild(
				m_MessageDatabaseObject.DLLFileName,
				L"DLLFileName = %1!hs!",
				m_MessageDatabaseObject.DLLFileName.c_str());
			if (m_MessageDatabaseObject.bIsExchange)
			{
				terminateBlock();

				auto szWrappedType =
					InterpretNumberAsStringProp(m_MessageDatabaseObject.WrappedType, PR_PROFILE_OPEN_FLAGS);
				addChild(
					m_MessageDatabaseObject.WrappedFlags,
					L"Wrapped Flags = 0x%1!08X!\r\n",
					m_MessageDatabaseObject.WrappedFlags.getData());
				addChild(
					m_MessageDatabaseObject.WrappedProviderUID,
					L"WrappedProviderUID = %1!ws!\r\n",
					guid::GUIDToStringAndName(m_MessageDatabaseObject.WrappedProviderUID).c_str());
				addChild(
					m_MessageDatabaseObject.WrappedType,
					L"WrappedType = 0x%1!08X! = %2!ws!\r\n",
					m_MessageDatabaseObject.WrappedType.getData(),
					szWrappedType.c_str());
				addChild(
					m_MessageDatabaseObject.ServerShortname,
					L"ServerShortname = %1!hs!\r\n",
					m_MessageDatabaseObject.ServerShortname.c_str());
				addChild(
					m_MessageDatabaseObject.MailboxDN,
					L"MailboxDN = %1!hs!",
					m_MessageDatabaseObject.MailboxDN.c_str());
			}

			switch (m_MessageDatabaseObject.MagicVersion)
			{
			case MDB_STORE_EID_V2_MAGIC:
			{
				terminateBlock();
				auto szV2Magic = flags::InterpretFlags(flagEidMagic, m_MessageDatabaseObject.v2.ulMagic);
				addChild(
					m_MessageDatabaseObject.v2.ulMagic,
					L"Magic = 0x%1!08X! = %2!ws!\r",
					m_MessageDatabaseObject.v2.ulMagic.getData(),
					szV2Magic.c_str());
				addChild(
					m_MessageDatabaseObject.v2.ulSize,
					L"Size = 0x%1!08X! = %1!d!\r\n",
					m_MessageDatabaseObject.v2.ulSize.getData());
				auto szV2Version = flags::InterpretFlags(flagEidVersion, m_MessageDatabaseObject.v2.ulVersion);
				addChild(
					m_MessageDatabaseObject.v2.ulVersion,
					L"Version = 0x%1!08X! = %2!ws!\r\n",
					m_MessageDatabaseObject.v2.ulVersion.getData(),
					szV2Version.c_str());
				addChild(
					m_MessageDatabaseObject.v2.ulOffsetDN,
					L"OffsetDN = 0x%1!08X!\r\n",
					m_MessageDatabaseObject.v2.ulOffsetDN.getData());
				addChild(
					m_MessageDatabaseObject.v2.ulOffsetFQDN,
					L"OffsetFQDN = 0x%1!08X!\r\n",
					m_MessageDatabaseObject.v2.ulOffsetFQDN.getData());
				addChild(m_MessageDatabaseObject.v2DN, L"DN = %1!hs!\r\n", m_MessageDatabaseObject.v2DN.c_str());
				addChild(m_MessageDatabaseObject.v2FQDN, L"FQDN = %1!ws!\r\n", m_MessageDatabaseObject.v2FQDN.c_str());

				addHeader(L"Reserved Bytes = ");
				addChild(m_MessageDatabaseObject.v2Reserved);
			}
			break;
			case MDB_STORE_EID_V3_MAGIC:
			{
				terminateBlock();
				auto szV3Magic = flags::InterpretFlags(flagEidMagic, m_MessageDatabaseObject.v3.ulMagic);
				addChild(
					m_MessageDatabaseObject.v3.ulMagic,
					L"Magic = 0x%1!08X! = %2!ws!\r\n",
					m_MessageDatabaseObject.v3.ulMagic.getData(),
					szV3Magic.c_str());
				addChild(
					m_MessageDatabaseObject.v3.ulSize,
					L"Size = 0x%1!08X! = %1!d!\r\n",
					m_MessageDatabaseObject.v3.ulSize.getData());
				auto szV3Version = flags::InterpretFlags(flagEidVersion, m_MessageDatabaseObject.v3.ulVersion);
				addChild(
					m_MessageDatabaseObject.v3.ulVersion,
					L"Version = 0x%1!08X! = %2!ws!\r\n",
					m_MessageDatabaseObject.v3.ulVersion.getData(),
					szV3Version.c_str());
				addChild(
					m_MessageDatabaseObject.v3.ulOffsetSmtpAddress,
					L"OffsetSmtpAddress = 0x%1!08X!\r\n",
					m_MessageDatabaseObject.v3.ulOffsetSmtpAddress.getData());
				addChild(
					m_MessageDatabaseObject.v3SmtpAddress,
					L"SmtpAddress = %1!ws!\r\n",
					m_MessageDatabaseObject.v3SmtpAddress.c_str());
				addHeader(L"Reserved Bytes = ");

				addChild(m_MessageDatabaseObject.v2Reserved);
			}
			break;
			}
		}
		else if (eidtFolder == m_ObjectType)
		{
			terminateBlock();
			auto szType = flags::InterpretFlags(flagMessageDatabaseObjectType, m_FolderOrMessage.Type);
			addChild(
				m_FolderOrMessage.Type,
				L"Folder Type = 0x%1!04X! = %2!ws!\r\n",
				m_FolderOrMessage.Type.getData(),
				szType.c_str());
			addChild(
				m_FolderOrMessage.FolderObject.DatabaseGUID,
				L"Database GUID = %1!ws!\r\n",
				guid::GUIDToStringAndName(m_FolderOrMessage.FolderObject.DatabaseGUID).c_str());
			addHeader(L"GlobalCounter = ");
			addChild(m_FolderOrMessage.FolderObject.GlobalCounter);
			terminateBlock();

			addHeader(L"Pad = ");
			addChild(m_FolderOrMessage.FolderObject.Pad);
		}
		else if (eidtMessage == m_ObjectType)
		{
			terminateBlock();
			auto szType = flags::InterpretFlags(flagMessageDatabaseObjectType, m_FolderOrMessage.Type);
			addChild(
				m_FolderOrMessage.Type,
				L"Message Type = 0x%1!04X! = %2!ws!\r\n",
				m_FolderOrMessage.Type.getData(),
				szType.c_str());
			addChild(
				m_FolderOrMessage.MessageObject.FolderDatabaseGUID,
				L"Folder Database GUID = %1!ws!\r\n",
				guid::GUIDToStringAndName(m_FolderOrMessage.MessageObject.FolderDatabaseGUID).c_str());
			addHeader(L"Folder GlobalCounter = ");
			addChild(m_FolderOrMessage.MessageObject.FolderGlobalCounter);
			terminateBlock();

			addHeader(L"Pad1 = ");
			addChild(m_FolderOrMessage.MessageObject.Pad1);
			terminateBlock();

			addChild(
				m_FolderOrMessage.MessageObject.MessageDatabaseGUID,
				L"Message Database GUID = %1!ws!\r\n",
				guid::GUIDToStringAndName(m_FolderOrMessage.MessageObject.MessageDatabaseGUID).c_str());
			addHeader(L"Message GlobalCounter = ");
			addChild(m_FolderOrMessage.MessageObject.MessageGlobalCounter);
			terminateBlock();

			addHeader(L"Pad2 = ");
			addChild(m_FolderOrMessage.MessageObject.Pad2);
		}
	}
} // namespace smartview