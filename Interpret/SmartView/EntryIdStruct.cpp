#include <StdAfx.h>
#include <Interpret/SmartView/EntryIdStruct.h>
#include <Interpret/SmartView/SmartView.h>
#include <core/utility/strings.h>
#include <core/interpret/guid.h>
#include <Interpret/InterpretProp.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	void EntryIdStruct::Parse()
	{
		m_ObjectType = eidtUnknown;
		if (m_Parser.RemainingBytes() < 4) return;
		m_abFlags0 = m_Parser.Get<byte>();
		m_abFlags1 = m_Parser.Get<byte>();
		m_abFlags23 = m_Parser.GetBYTES(2);
		m_ProviderUID = m_Parser.Get<GUID>();

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
				if (m_ContactAddressBookObject.EntryIDCount != 0 &&
					m_ContactAddressBookObject.EntryIDCount < cbRemainingBytes)
				{
					cbRemainingBytes = m_ContactAddressBookObject.EntryIDCount;
				}

				EntryIdStruct entryIdStruct;
				entryIdStruct.parse(m_Parser, cbRemainingBytes, false);
				m_ContactAddressBookObject.lpEntryID.push_back(entryIdStruct);
			}
			break;
			case eidtWAB:
			{
				m_ObjectType = eidtWAB;

				m_WAB.Type = m_Parser.Get<BYTE>();

				EntryIdStruct entryIdStruct;
				entryIdStruct.parse(m_Parser, false);
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
					CSTR_EQUAL == CompareStringA(
									  LOCALE_INVARIANT,
									  NORM_IGNORECASE,
									  m_MessageDatabaseObject.DLLFileName.c_str(),
									  -1,
									  "emsmdb.dll", // STRING_OK
									  -1))
				{
					m_MessageDatabaseObject.bIsExchange = true;
					auto cbRead = m_Parser.GetCurrentOffset();
					// Advance to the next multiple of 4
					m_Parser.advance(3 - (cbRead + 3) % 4);
					m_MessageDatabaseObject.WrappedFlags = m_Parser.Get<DWORD>();
					m_MessageDatabaseObject.WrappedProviderUID = m_Parser.Get<GUID>();
					m_MessageDatabaseObject.WrappedType = m_Parser.Get<DWORD>();
					m_MessageDatabaseObject.ServerShortname = m_Parser.GetStringA();

					// Test if we have a magic value. Some PF EIDs also have a mailbox DN and we need to accomodate them
					if (m_MessageDatabaseObject.WrappedType & OPENSTORE_PUBLIC)
					{
						cbRead = m_Parser.GetCurrentOffset();
						m_MessageDatabaseObject.MagicVersion = m_Parser.Get<DWORD>();
						m_Parser.SetCurrentOffset(cbRead);
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
						m_MessageDatabaseObject.MailboxDN = m_Parser.GetStringA();
					}

					// Check again for a magic value
					cbRead = m_Parser.GetCurrentOffset();
					m_MessageDatabaseObject.MagicVersion = m_Parser.Get<DWORD>();
					m_Parser.SetCurrentOffset(cbRead);

					switch (m_MessageDatabaseObject.MagicVersion)
					{
					case MDB_STORE_EID_V2_MAGIC:
						if (m_Parser.RemainingBytes() >= MDB_STORE_EID_V2::size + sizeof(WCHAR))
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
						if (m_Parser.RemainingBytes() >= MDB_STORE_EID_V3::size + sizeof(WCHAR))
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
			auto tempBlock = block();
			tempBlock.setOffset(m_abFlags0.getOffset());
			tempBlock.setSize(4);
			addBlock(tempBlock, L"abFlags = 0x00000000\r\n");
		}
		else if (0 == (m_abFlags1 | m_abFlags23[0] | m_abFlags23[1]))
		{
			addBlock(
				m_abFlags0,
				L"abFlags[0] = 0x%1!02X!= %2!ws!\r\n",
				m_abFlags0.getData(),
				interpretprop::InterpretFlags(flagEntryId0, m_abFlags0).c_str());
			auto tempBlock = block();
			tempBlock.setOffset(m_abFlags1.getOffset());
			tempBlock.setSize(3);
			addBlock(tempBlock, L"abFlags[1..3] = 0x000000\r\n");
		}
		else
		{
			addBlock(
				m_abFlags0,
				L"abFlags[0] = 0x%1!02X!= %2!ws!\r\n",
				m_abFlags0.getData(),
				interpretprop::InterpretFlags(flagEntryId0, m_abFlags0).c_str());
			addBlock(
				m_abFlags1,
				L"abFlags[1] = 0x%1!02X!= %2!ws!\r\n",
				m_abFlags1.getData(),
				interpretprop::InterpretFlags(flagEntryId1, m_abFlags1).c_str());
			addBlock(m_abFlags23, L"abFlags[2..3] = 0x%1!02X!%2!02X!\r\n", m_abFlags23[0], m_abFlags23[1]);
		}

		addBlock(m_ProviderUID, L"Provider GUID = %1!ws!", guid::GUIDToStringAndName(m_ProviderUID).c_str());

		if (eidtEphemeral == m_ObjectType)
		{
			terminateBlock();

			auto szVersion = interpretprop::InterpretFlags(flagExchangeABVersion, m_EphemeralObject.Version);
			addBlock(
				m_EphemeralObject.Version,
				L"Version = 0x%1!08X! = %2!ws!\r\n",
				m_EphemeralObject.Version.getData(),
				szVersion.c_str());

			auto szType = InterpretNumberAsStringProp(m_EphemeralObject.Type, PR_DISPLAY_TYPE);
			addBlock(
				m_EphemeralObject.Type, L"Type = 0x%1!08X! = %2!ws!", m_EphemeralObject.Type.getData(), szType.c_str());
		}
		else if (eidtOneOff == m_ObjectType)
		{
			auto szFlags = interpretprop::InterpretFlags(flagOneOffEntryId, m_OneOffRecipientObject.Bitmask);
			if (MAPI_UNICODE & m_OneOffRecipientObject.Bitmask)
			{
				terminateBlock();
				addBlock(
					m_OneOffRecipientObject.Bitmask,
					L"dwBitmask: 0x%1!08X! = %2!ws!\r\n",
					m_OneOffRecipientObject.Bitmask.getData(),
					szFlags.c_str());
				addBlock(
					m_OneOffRecipientObject.Unicode.DisplayName,
					L"szDisplayName = %1!ws!\r\n",
					m_OneOffRecipientObject.Unicode.DisplayName.c_str());
				addBlock(
					m_OneOffRecipientObject.Unicode.AddressType,
					L"szAddressType = %1!ws!\r\n",
					m_OneOffRecipientObject.Unicode.AddressType.c_str());
				addBlock(
					m_OneOffRecipientObject.Unicode.EmailAddress,
					L"szEmailAddress = %1!ws!",
					m_OneOffRecipientObject.Unicode.EmailAddress.c_str());
			}
			else
			{
				terminateBlock();
				addBlock(
					m_OneOffRecipientObject.Bitmask,
					L"dwBitmask: 0x%1!08X! = %2!ws!\r\n",
					m_OneOffRecipientObject.Bitmask.getData(),
					szFlags.c_str());
				addBlock(
					m_OneOffRecipientObject.ANSI.DisplayName,
					L"szDisplayName = %1!hs!\r\n",
					m_OneOffRecipientObject.ANSI.DisplayName.c_str());
				addBlock(
					m_OneOffRecipientObject.ANSI.AddressType,
					L"szAddressType = %1!hs!\r\n",
					m_OneOffRecipientObject.ANSI.AddressType.c_str());
				addBlock(
					m_OneOffRecipientObject.ANSI.EmailAddress,
					L"szEmailAddress = %1!hs!",
					m_OneOffRecipientObject.ANSI.EmailAddress.c_str());
			}
		}
		else if (eidtAddressBook == m_ObjectType)
		{
			terminateBlock();
			auto szVersion = interpretprop::InterpretFlags(flagExchangeABVersion, m_AddressBookObject.Version);
			addBlock(
				m_AddressBookObject.Version,
				L"Version = 0x%1!08X! = %2!ws!\r\n",
				m_AddressBookObject.Version.getData(),
				szVersion.c_str());
			auto szType = InterpretNumberAsStringProp(m_AddressBookObject.Type, PR_DISPLAY_TYPE);
			addBlock(
				m_AddressBookObject.Type,
				L"Type = 0x%1!08X! = %2!ws!\r\n",
				m_AddressBookObject.Type.getData(),
				szType.c_str());
			addBlock(m_AddressBookObject.X500DN, L"X500DN = %1!hs!", m_AddressBookObject.X500DN.c_str());
		}
		// Contact Address Book / Personal Distribution List (PDL)
		else if (eidtContact == m_ObjectType)
		{
			terminateBlock();
			auto szVersion = interpretprop::InterpretFlags(flagContabVersion, m_ContactAddressBookObject.Version);
			addBlock(
				m_ContactAddressBookObject.Version,
				L"Version = 0x%1!08X! = %2!ws!\r\n",
				m_ContactAddressBookObject.Version.getData(),
				szVersion.c_str());

			auto szType = interpretprop::InterpretFlags(flagContabType, m_ContactAddressBookObject.Type);
			addBlock(
				m_ContactAddressBookObject.Type,
				L"Type = 0x%1!08X! = %2!ws!\r\n",
				m_ContactAddressBookObject.Type.getData(),
				szType.c_str());

			switch (m_ContactAddressBookObject.Type)
			{
			case CONTAB_USER:
			case CONTAB_DISTLIST:
				addBlock(
					m_ContactAddressBookObject.Index,
					L"Index = 0x%1!08X! = %2!ws!\r\n",
					m_ContactAddressBookObject.Index.getData(),
					interpretprop::InterpretFlags(flagContabIndex, m_ContactAddressBookObject.Index).c_str());
				addBlock(
					m_ContactAddressBookObject.EntryIDCount,
					L"EntryIDCount = 0x%1!08X!\r\n",
					m_ContactAddressBookObject.EntryIDCount.getData());
				break;
			case CONTAB_ROOT:
			case CONTAB_CONTAINER:
			case CONTAB_SUBROOT:
				addBlock(
					m_ContactAddressBookObject.muidID,
					L"muidID = %1!ws!\r\n",
					guid::GUIDToStringAndName(m_ContactAddressBookObject.muidID).c_str());
				break;
			default:
				break;
			}

			for (const auto& entry : m_ContactAddressBookObject.lpEntryID)
			{
				addBlock(entry.getBlock());
			}
		}
		else if (eidtWAB == m_ObjectType)
		{
			terminateBlock();
			addBlock(
				m_WAB.Type,
				L"Wrapped Entry Type = 0x%1!02X! = %2!ws!\r\n",
				m_WAB.Type.getData(),
				interpretprop::InterpretFlags(flagWABEntryIDType, m_WAB.Type).c_str());

			for (auto& entry : m_WAB.lpEntryID)
			{
				addBlock(entry.getBlock());
			}
		}
		else if (eidtMessageDatabase == m_ObjectType)
		{
			terminateBlock();
			auto szVersion = interpretprop::InterpretFlags(flagMDBVersion, m_MessageDatabaseObject.Version);
			addBlock(
				m_MessageDatabaseObject.Version,
				L"Version = 0x%1!02X! = %2!ws!\r\n",
				m_MessageDatabaseObject.Version.getData(),
				szVersion.c_str());

			auto szFlag = interpretprop::InterpretFlags(flagMDBFlag, m_MessageDatabaseObject.Flag);
			addBlock(
				m_MessageDatabaseObject.Flag,
				L"Flag = 0x%1!02X! = %2!ws!\r\n",
				m_MessageDatabaseObject.Flag.getData(),
				szFlag.c_str());

			addBlock(
				m_MessageDatabaseObject.DLLFileName,
				L"DLLFileName = %1!hs!",
				m_MessageDatabaseObject.DLLFileName.c_str());
			if (m_MessageDatabaseObject.bIsExchange)
			{
				terminateBlock();

				auto szWrappedType =
					InterpretNumberAsStringProp(m_MessageDatabaseObject.WrappedType, PR_PROFILE_OPEN_FLAGS);
				addBlock(
					m_MessageDatabaseObject.WrappedFlags,
					L"Wrapped Flags = 0x%1!08X!\r\n",
					m_MessageDatabaseObject.WrappedFlags.getData());
				addBlock(
					m_MessageDatabaseObject.WrappedProviderUID,
					L"WrappedProviderUID = %1!ws!\r\n",
					guid::GUIDToStringAndName(m_MessageDatabaseObject.WrappedProviderUID).c_str());
				addBlock(
					m_MessageDatabaseObject.WrappedType,
					L"WrappedType = 0x%1!08X! = %2!ws!\r\n",
					m_MessageDatabaseObject.WrappedType.getData(),
					szWrappedType.c_str());
				addBlock(
					m_MessageDatabaseObject.ServerShortname,
					L"ServerShortname = %1!hs!\r\n",
					m_MessageDatabaseObject.ServerShortname.c_str());
				addBlock(
					m_MessageDatabaseObject.MailboxDN,
					L"MailboxDN = %1!hs!",
					m_MessageDatabaseObject.MailboxDN.c_str());
			}

			switch (m_MessageDatabaseObject.MagicVersion)
			{
			case MDB_STORE_EID_V2_MAGIC:
			{
				terminateBlock();
				auto szV2Magic = interpretprop::InterpretFlags(flagEidMagic, m_MessageDatabaseObject.v2.ulMagic);
				addBlock(
					m_MessageDatabaseObject.v2.ulMagic,
					L"Magic = 0x%1!08X! = %2!ws!\r",
					m_MessageDatabaseObject.v2.ulMagic.getData(),
					szV2Magic.c_str());
				addBlock(
					m_MessageDatabaseObject.v2.ulSize,
					L"Size = 0x%1!08X! = %1!d!\r\n",
					m_MessageDatabaseObject.v2.ulSize.getData());
				auto szV2Version = interpretprop::InterpretFlags(flagEidVersion, m_MessageDatabaseObject.v2.ulVersion);
				addBlock(
					m_MessageDatabaseObject.v2.ulVersion,
					L"Version = 0x%1!08X! = %2!ws!\r\n",
					m_MessageDatabaseObject.v2.ulVersion.getData(),
					szV2Version.c_str());
				addBlock(
					m_MessageDatabaseObject.v2.ulOffsetDN,
					L"OffsetDN = 0x%1!08X!\r\n",
					m_MessageDatabaseObject.v2.ulOffsetDN.getData());
				addBlock(
					m_MessageDatabaseObject.v2.ulOffsetFQDN,
					L"OffsetFQDN = 0x%1!08X!\r\n",
					m_MessageDatabaseObject.v2.ulOffsetFQDN.getData());
				addBlock(m_MessageDatabaseObject.v2DN, L"DN = %1!hs!\r\n", m_MessageDatabaseObject.v2DN.c_str());
				addBlock(m_MessageDatabaseObject.v2FQDN, L"FQDN = %1!ws!\r\n", m_MessageDatabaseObject.v2FQDN.c_str());

				addHeader(L"Reserved Bytes = ");
				addBlock(m_MessageDatabaseObject.v2Reserved);
			}
			break;
			case MDB_STORE_EID_V3_MAGIC:
			{
				terminateBlock();
				auto szV3Magic = interpretprop::InterpretFlags(flagEidMagic, m_MessageDatabaseObject.v3.ulMagic);
				addBlock(
					m_MessageDatabaseObject.v3.ulMagic,
					L"Magic = 0x%1!08X! = %2!ws!\r\n",
					m_MessageDatabaseObject.v3.ulMagic.getData(),
					szV3Magic.c_str());
				addBlock(
					m_MessageDatabaseObject.v3.ulSize,
					L"Size = 0x%1!08X! = %1!d!\r\n",
					m_MessageDatabaseObject.v3.ulSize.getData());
				auto szV3Version = interpretprop::InterpretFlags(flagEidVersion, m_MessageDatabaseObject.v3.ulVersion);
				addBlock(
					m_MessageDatabaseObject.v3.ulVersion,
					L"Version = 0x%1!08X! = %2!ws!\r\n",
					m_MessageDatabaseObject.v3.ulVersion.getData(),
					szV3Version.c_str());
				addBlock(
					m_MessageDatabaseObject.v3.ulOffsetSmtpAddress,
					L"OffsetSmtpAddress = 0x%1!08X!\r\n",
					m_MessageDatabaseObject.v3.ulOffsetSmtpAddress.getData());
				addBlock(
					m_MessageDatabaseObject.v3SmtpAddress,
					L"SmtpAddress = %1!ws!\r\n",
					m_MessageDatabaseObject.v3SmtpAddress.c_str());
				addHeader(L"Reserved Bytes = ");

				addBlock(m_MessageDatabaseObject.v2Reserved);
			}
			break;
			}
		}
		else if (eidtFolder == m_ObjectType)
		{
			terminateBlock();
			auto szType = interpretprop::InterpretFlags(flagMessageDatabaseObjectType, m_FolderOrMessage.Type);
			addBlock(
				m_FolderOrMessage.Type,
				L"Folder Type = 0x%1!04X! = %2!ws!\r\n",
				m_FolderOrMessage.Type.getData(),
				szType.c_str());
			addBlock(
				m_FolderOrMessage.FolderObject.DatabaseGUID,
				L"Database GUID = %1!ws!\r\n",
				guid::GUIDToStringAndName(m_FolderOrMessage.FolderObject.DatabaseGUID).c_str());
			addHeader(L"GlobalCounter = ");
			addBlock(m_FolderOrMessage.FolderObject.GlobalCounter);
			terminateBlock();

			addHeader(L"Pad = ");
			addBlock(m_FolderOrMessage.FolderObject.Pad);
		}
		else if (eidtMessage == m_ObjectType)
		{
			terminateBlock();
			auto szType = interpretprop::InterpretFlags(flagMessageDatabaseObjectType, m_FolderOrMessage.Type);
			addBlock(
				m_FolderOrMessage.Type,
				L"Message Type = 0x%1!04X! = %2!ws!\r\n",
				m_FolderOrMessage.Type.getData(),
				szType.c_str());
			addBlock(
				m_FolderOrMessage.MessageObject.FolderDatabaseGUID,
				L"Folder Database GUID = %1!ws!\r\n",
				guid::GUIDToStringAndName(m_FolderOrMessage.MessageObject.FolderDatabaseGUID).c_str());
			addHeader(L"Folder GlobalCounter = ");
			addBlock(m_FolderOrMessage.MessageObject.FolderGlobalCounter);
			terminateBlock();

			addHeader(L"Pad1 = ");
			addBlock(m_FolderOrMessage.MessageObject.Pad1);
			terminateBlock();

			addBlock(
				m_FolderOrMessage.MessageObject.MessageDatabaseGUID,
				L"Message Database GUID = %1!ws!\r\n",
				guid::GUIDToStringAndName(m_FolderOrMessage.MessageObject.MessageDatabaseGUID).c_str());
			addHeader(L"Message GlobalCounter = ");
			addBlock(m_FolderOrMessage.MessageObject.MessageGlobalCounter);
			terminateBlock();

			addHeader(L"Pad2 = ");
			addBlock(m_FolderOrMessage.MessageObject.Pad2);
		}
	}
} // namespace smartview