#include <StdAfx.h>
#include <Interpret/SmartView/EntryIdStruct.h>
#include <Interpret/SmartView/SmartView.h>
#include <Interpret/String.h>
#include <Interpret/Guids.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>

namespace smartview
{
	EntryIdStruct::EntryIdStruct() {}

	void EntryIdStruct::Parse()
	{
		m_ObjectType = eidtUnknown;
		m_abFlags = m_Parser.GetBlockBYTES(4);
		if (!m_abFlags.getSize()) return;
		m_ProviderUID = m_Parser.GetBlock<GUID>();

		// Ephemeral entry ID:
		if (m_abFlags.getData()[0] == EPHEMERAL)
		{
			m_ObjectType = eidtEphemeral;
		}
		// One Off Recipient
		else if (m_ProviderUID.getData() == guid::muidOOP)
		{
			m_ObjectType = eidtOneOff;
		}
		// Address Book Recipient
		else if (m_ProviderUID.getData() == guid::muidEMSAB)
		{
			m_ObjectType = eidtAddressBook;
		}
		// Contact Address Book / Personal Distribution List (PDL)
		else if (m_ProviderUID.getData() == guid::muidContabDLL)
		{
			m_ObjectType = eidtContact;
		}
		// message store objects
		else if (m_ProviderUID.getData() == guid::muidStoreWrap)
		{
			m_ObjectType = eidtMessageDatabase;
		}
		else if (m_ProviderUID.getData() == guid::WAB_GUID)
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
				m_EphemeralObject.Version = m_Parser.GetBlock<DWORD>();
				m_EphemeralObject.Type = m_Parser.GetBlock<DWORD>();
				break;
				// One Off Recipient
			case eidtOneOff:
				m_OneOffRecipientObject.Bitmask = m_Parser.GetBlock<DWORD>();
				if (MAPI_UNICODE & m_OneOffRecipientObject.Bitmask.getData())
				{
					m_OneOffRecipientObject.Unicode.DisplayName = m_Parser.GetBlockStringW();
					m_OneOffRecipientObject.Unicode.AddressType = m_Parser.GetBlockStringW();
					m_OneOffRecipientObject.Unicode.EmailAddress = m_Parser.GetBlockStringW();
				}
				else
				{
					m_OneOffRecipientObject.ANSI.DisplayName = m_Parser.GetBlockStringA();
					m_OneOffRecipientObject.ANSI.AddressType = m_Parser.GetBlockStringA();
					m_OneOffRecipientObject.ANSI.EmailAddress = m_Parser.GetBlockStringA();
				}
				break;
				// Address Book Recipient
			case eidtAddressBook:
				m_AddressBookObject.Version = m_Parser.GetBlock<DWORD>();
				m_AddressBookObject.Type = m_Parser.GetBlock<DWORD>();
				m_AddressBookObject.X500DN = m_Parser.GetBlockStringA();
				break;
				// Contact Address Book / Personal Distribution List (PDL)
			case eidtContact:
			{
				m_ContactAddressBookObject.Version = m_Parser.GetBlock<DWORD>();
				m_ContactAddressBookObject.Type = m_Parser.GetBlock<DWORD>();

				if (CONTAB_CONTAINER == m_ContactAddressBookObject.Type.getData())
				{
					m_ContactAddressBookObject.muidID = m_Parser.GetBlock<GUID>();
				}
				else // Assume we've got some variation on the user/distlist format
				{
					m_ContactAddressBookObject.Index = m_Parser.GetBlock<DWORD>();
					m_ContactAddressBookObject.EntryIDCount = m_Parser.GetBlock<DWORD>();
				}

				// Read the wrapped entry ID from the remaining data
				auto cbRemainingBytes = m_Parser.RemainingBytes();

				// If we already got a size, use it, else we just read the rest of the structure
				if (m_ContactAddressBookObject.EntryIDCount.getData() != 0 &&
					m_ContactAddressBookObject.EntryIDCount.getData() < cbRemainingBytes)
				{
					cbRemainingBytes = m_ContactAddressBookObject.EntryIDCount.getData();
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

				m_WAB.Type = m_Parser.GetBlock<BYTE>();

				EntryIdStruct entryIdStruct;
				entryIdStruct.Init(m_Parser.RemainingBytes(), m_Parser.GetCurrentAddress());
				entryIdStruct.DisableJunkParsing();
				entryIdStruct.EnsureParsed();
				m_Parser.Advance(entryIdStruct.GetCurrentOffset());
				m_WAB.lpEntryID.push_back(entryIdStruct);
			}
			break;
			// message store objects
			case eidtMessageDatabase:
				m_MessageDatabaseObject.Version = m_Parser.GetBlock<BYTE>();
				m_MessageDatabaseObject.Flag = m_Parser.GetBlock<BYTE>();
				m_MessageDatabaseObject.DLLFileName = m_Parser.GetBlockStringA();
				m_MessageDatabaseObject.bIsExchange = false;

				// We only know how to parse emsmdb.dll's wrapped entry IDs
				if (!m_MessageDatabaseObject.DLLFileName.getData().empty() &&
					CSTR_EQUAL == CompareStringA(
									  LOCALE_INVARIANT,
									  NORM_IGNORECASE,
									  m_MessageDatabaseObject.DLLFileName.getData().c_str(),
									  -1,
									  "emsmdb.dll", // STRING_OK
									  -1))
				{
					m_MessageDatabaseObject.bIsExchange = true;
					auto cbRead = m_Parser.GetCurrentOffset();
					// Advance to the next multiple of 4
					m_Parser.Advance(3 - (cbRead + 3) % 4);
					m_MessageDatabaseObject.WrappedFlags = m_Parser.GetBlock<DWORD>();
					m_MessageDatabaseObject.WrappedProviderUID = m_Parser.GetBlock<GUID>();
					m_MessageDatabaseObject.WrappedType = m_Parser.GetBlock<DWORD>();
					m_MessageDatabaseObject.ServerShortname = m_Parser.GetBlockStringA();

					// Test if we have a magic value. Some PF EIDs also have a mailbox DN and we need to accomodate them
					if (m_MessageDatabaseObject.WrappedType.getData() & OPENSTORE_PUBLIC)
					{
						cbRead = m_Parser.GetCurrentOffset();
						m_MessageDatabaseObject.MagicVersion = m_Parser.GetBlock<DWORD>();
						m_Parser.SetCurrentOffset(cbRead);
					}
					else
					{
						m_MessageDatabaseObject.MagicVersion.setOffset(0);
						m_MessageDatabaseObject.MagicVersion.setSize(0);
						m_MessageDatabaseObject.MagicVersion.setData(MDB_STORE_EID_V1_VERSION);
					}

					// Either we're not a PF eid, or this PF EID wasn't followed directly by a magic value
					if (!(m_MessageDatabaseObject.WrappedType.getData() & OPENSTORE_PUBLIC) ||
						m_MessageDatabaseObject.MagicVersion.getData() != MDB_STORE_EID_V2_MAGIC &&
							m_MessageDatabaseObject.MagicVersion.getData() != MDB_STORE_EID_V3_MAGIC)
					{
						m_MessageDatabaseObject.MailboxDN = m_Parser.GetBlockStringA();
					}

					// Check again for a magic value
					cbRead = m_Parser.GetCurrentOffset();
					m_MessageDatabaseObject.MagicVersion = m_Parser.GetBlock<DWORD>();
					m_Parser.SetCurrentOffset(cbRead);

					switch (m_MessageDatabaseObject.MagicVersion.getData())
					{
					case MDB_STORE_EID_V2_MAGIC:
						if (m_Parser.RemainingBytes() >= m_MessageDatabaseObject.v2.size + sizeof(WCHAR))
						{
							m_MessageDatabaseObject.v2.ulMagic = m_Parser.GetBlock<DWORD>();
							m_MessageDatabaseObject.v2.ulSize = m_Parser.GetBlock<DWORD>();
							m_MessageDatabaseObject.v2.ulVersion = m_Parser.GetBlock<DWORD>();
							m_MessageDatabaseObject.v2.ulOffsetDN = m_Parser.GetBlock<DWORD>();
							m_MessageDatabaseObject.v2.ulOffsetFQDN = m_Parser.GetBlock<DWORD>();
							if (m_MessageDatabaseObject.v2.ulOffsetDN.getData())
							{
								m_MessageDatabaseObject.v2DN = m_Parser.GetBlockStringA();
							}

							if (m_MessageDatabaseObject.v2.ulOffsetFQDN.getData())
							{
								m_MessageDatabaseObject.v2FQDN = m_Parser.GetBlockStringW();
							}

							m_MessageDatabaseObject.v2Reserved = m_Parser.GetBlockBYTES(2);
						}
						break;
					case MDB_STORE_EID_V3_MAGIC:
						if (m_Parser.RemainingBytes() >= m_MessageDatabaseObject.v3.size + sizeof(WCHAR))
						{
							m_MessageDatabaseObject.v3.ulMagic = m_Parser.GetBlock<DWORD>();
							m_MessageDatabaseObject.v3.ulSize = m_Parser.GetBlock<DWORD>();
							m_MessageDatabaseObject.v3.ulVersion = m_Parser.GetBlock<DWORD>();
							m_MessageDatabaseObject.v3.ulOffsetSmtpAddress = m_Parser.GetBlock<DWORD>();
							if (m_MessageDatabaseObject.v3.ulOffsetSmtpAddress.getData())
							{
								m_MessageDatabaseObject.v3SmtpAddress = m_Parser.GetBlockStringW();
							}

							m_MessageDatabaseObject.v2Reserved = m_Parser.GetBlockBYTES(2);
						}
						break;
					}
				}
				break;
				// Exchange message store folder
			case eidtFolder:
				m_FolderOrMessage.Type = m_Parser.GetBlock<WORD>();
				m_FolderOrMessage.FolderObject.DatabaseGUID = m_Parser.GetBlock<GUID>();
				m_FolderOrMessage.FolderObject.GlobalCounter = m_Parser.GetBlockBYTES(6);
				m_FolderOrMessage.FolderObject.Pad = m_Parser.GetBlockBYTES(2);
				break;
				// Exchange message store message
			case eidtMessage:
				m_FolderOrMessage.Type = m_Parser.GetBlock<WORD>();
				m_FolderOrMessage.MessageObject.FolderDatabaseGUID = m_Parser.GetBlock<GUID>();
				m_FolderOrMessage.MessageObject.FolderGlobalCounter = m_Parser.GetBlockBYTES(6);
				m_FolderOrMessage.MessageObject.Pad1 = m_Parser.GetBlockBYTES(2);

				m_FolderOrMessage.MessageObject.MessageDatabaseGUID = m_Parser.GetBlock<GUID>();
				m_FolderOrMessage.MessageObject.MessageGlobalCounter = m_Parser.GetBlockBYTES(6);
				m_FolderOrMessage.MessageObject.Pad2 = m_Parser.GetBlockBYTES(2);
				break;
			}
		}

		// Check if we have an unidentified short term entry ID:
		if (eidtUnknown == m_ObjectType && m_abFlags.getData()[0] & MAPI_SHORTTERM) m_ObjectType = eidtShortTerm;
	}

	void EntryIdStruct::ParseBlocks()
	{
		switch (m_ObjectType)
		{
		case eidtUnknown:
			addHeader(L"Entry ID:\r\n");
			break;
		case eidtEphemeral:
			addHeader(L"Ephemeral Entry ID:\r\n");
			break;
		case eidtShortTerm:
			addHeader(L"Short Term Entry ID:\r\n");
			break;
		case eidtFolder:
			addHeader(L"Exchange Folder Entry ID:\r\n");
			break;
		case eidtMessage:
			addHeader(L"Exchange Message Entry ID:\r\n");
			break;
		case eidtMessageDatabase:
			addHeader(L"MAPI Message Store Entry ID:\r\n");
			break;
		case eidtOneOff:
			addHeader(L"MAPI One Off Recipient Entry ID:\r\n");
			break;
		case eidtAddressBook:
			addHeader(L"Exchange Address Entry ID:\r\n");
			break;
		case eidtContact:
			addHeader(L"Contact Address Book / PDL Entry ID:\r\n");
			break;
		case eidtWAB:
			addHeader(L"Wrapped Entry ID:\r\n");
			break;
		}

		if (!m_abFlags.getSize()) return;
		if (0 == (m_abFlags.getData()[0] | m_abFlags.getData()[1] | m_abFlags.getData()[2] | m_abFlags.getData()[3]))
		{
			addHeader(L"abFlags = 0x00000000\r\n");
		}
		else if (0 == (m_abFlags.getData()[1] | m_abFlags.getData()[2] | m_abFlags.getData()[3]))
		{
			addBlock(
				m_abFlags,
				L"abFlags[0] = 0x%1!02X!= %2!ws!\r\n",
				m_abFlags.getData()[0],
				interpretprop::InterpretFlags(flagEntryId0, m_abFlags.getData()[0]).c_str());
			addHeader(L"abFlags[1..3] = 0x000000\r\n");
		}
		else
		{
			addBlock(
				m_abFlags,
				L"abFlags[0] = 0x%1!02X!= %2!ws!\r\n"
				L"abFlags[1] = 0x%3!02X!= %4!ws!\r\n"
				L"abFlags[2..3] = 0x%5!02X!%6!02X!\r\n",
				m_abFlags.getData()[0],
				interpretprop::InterpretFlags(flagEntryId0, m_abFlags.getData()[0]).c_str(),
				m_abFlags.getData()[1],
				interpretprop::InterpretFlags(flagEntryId1, m_abFlags.getData()[1]).c_str(),
				m_abFlags.getData()[2],
				m_abFlags.getData()[3]);
		}

		addBlock(m_ProviderUID, L"Provider GUID = %1!ws!", guid::GUIDToStringAndName(m_ProviderUID.getData()).c_str());

		if (eidtEphemeral == m_ObjectType)
		{
			addHeader(L"\r\n");

			auto szVersion = interpretprop::InterpretFlags(flagExchangeABVersion, m_EphemeralObject.Version.getData());
			addBlock(
				m_EphemeralObject.Version,
				L"Version = 0x%1!08X! = %2!ws!\r\n",
				m_EphemeralObject.Version.getData(),
				szVersion.c_str());

			auto szType = InterpretNumberAsStringProp(m_EphemeralObject.Type.getData(), PR_DISPLAY_TYPE);
			addBlock(
				m_EphemeralObject.Type, L"Type = 0x%1!08X! = %2!ws!", m_EphemeralObject.Type.getData(), szType.c_str());
		}
		else if (eidtOneOff == m_ObjectType)
		{
			auto szFlags = interpretprop::InterpretFlags(flagOneOffEntryId, m_OneOffRecipientObject.Bitmask.getData());
			if (MAPI_UNICODE & m_OneOffRecipientObject.Bitmask.getData())
			{
				addHeader(L"\r\n");
				addBlock(
					m_OneOffRecipientObject.Bitmask,
					L"dwBitmask: 0x%1!08X! = %2!ws!\r\n",
					m_OneOffRecipientObject.Bitmask.getData(),
					szFlags.c_str());
				addBlock(
					m_OneOffRecipientObject.Unicode.DisplayName,
					L"szDisplayName = %1!ws!\r\n",
					m_OneOffRecipientObject.Unicode.DisplayName.getData().c_str());
				addBlock(
					m_OneOffRecipientObject.Unicode.AddressType,
					L"szAddressType = %1!ws!\r\n",
					m_OneOffRecipientObject.Unicode.AddressType.getData().c_str());
				addBlock(
					m_OneOffRecipientObject.Unicode.EmailAddress,
					L"szEmailAddress = %1!ws!",
					m_OneOffRecipientObject.Unicode.EmailAddress.getData().c_str());
			}
			else
			{
				addHeader(L"\r\n");
				addBlock(
					m_OneOffRecipientObject.Bitmask,
					L"dwBitmask: 0x%1!08X! = %2!ws!\r\n",
					m_OneOffRecipientObject.Bitmask.getData(),
					szFlags.c_str());
				addBlock(
					m_OneOffRecipientObject.ANSI.DisplayName,
					L"szDisplayName = %1!ws!\r\n",
					m_OneOffRecipientObject.ANSI.DisplayName.getData().c_str());
				addBlock(
					m_OneOffRecipientObject.ANSI.AddressType,
					L"szAddressType = %1!ws!\r\n",
					m_OneOffRecipientObject.ANSI.AddressType.getData().c_str());
				addBlock(
					m_OneOffRecipientObject.ANSI.EmailAddress,
					L"szEmailAddress = %1!ws!",
					m_OneOffRecipientObject.ANSI.EmailAddress.getData().c_str());
			}
		}
		else if (eidtAddressBook == m_ObjectType)
		{
			addHeader(L"\r\n");
			auto szVersion =
				interpretprop::InterpretFlags(flagExchangeABVersion, m_AddressBookObject.Version.getData());
			addBlock(
				m_AddressBookObject.Version,
				L"Version = 0x%1!08X! = %2!ws!\r\n",
				m_AddressBookObject.Version.getData(),
				szVersion.c_str());
			auto szType = InterpretNumberAsStringProp(m_AddressBookObject.Type.getData(), PR_DISPLAY_TYPE);
			addBlock(
				m_AddressBookObject.Type,
				L"Type = 0x%1!08X! = %2!ws!\r\n",
				m_AddressBookObject.Type.getData(),
				szType.c_str());
			addBlock(m_AddressBookObject.X500DN, L"X500DN = %1!hs!", m_AddressBookObject.X500DN.getData().c_str());
		}
		// Contact Address Book / Personal Distribution List (PDL)
		else if (eidtContact == m_ObjectType)
		{
			auto szVersion =
				interpretprop::InterpretFlags(flagContabVersion, m_ContactAddressBookObject.Version.getData());
			addBlock(
				m_ContactAddressBookObject.Version,
				L"\r\nVersion = 0x%1!08X! = %2!ws!\r\n",
				m_ContactAddressBookObject.Version.getData(),
				szVersion.c_str());

			auto szType = interpretprop::InterpretFlags(flagContabType, m_ContactAddressBookObject.Type.getData());
			addBlock(
				m_ContactAddressBookObject.Type,
				L"Type = 0x%1!08X! = %2!ws!\r\n",
				m_ContactAddressBookObject.Type.getData(),
				szType.c_str());

			switch (m_ContactAddressBookObject.Type.getData())
			{
			case CONTAB_USER:
			case CONTAB_DISTLIST:
				addBlock(
					m_ContactAddressBookObject.Index,
					L"Index = 0x%1!08X! = %2!ws!\r\n",
					m_ContactAddressBookObject.Index.getData(),
					interpretprop::InterpretFlags(flagContabIndex, m_ContactAddressBookObject.Index.getData()).c_str());
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
					guid::GUIDToStringAndName(m_ContactAddressBookObject.muidID.getData()).c_str());
				break;
			default:
				break;
			}

			// TODO: Finish this
			//for (auto entry : m_ContactAddressBookObject.lpEntryID)
			//{
			//	szEntryId += entry.ToString();
			//}
		}
		else if (eidtWAB == m_ObjectType)
		{
			addHeader(L"\r\n");
			addBlock(
				m_WAB.Type,
				L"Wrapped Entry Type = 0x%1!02X! = %2!ws!\r\n",
				m_WAB.Type.getData(),
				interpretprop::InterpretFlags(flagWABEntryIDType, m_WAB.Type.getData()).c_str());

			// TODO: Finish this
			//for (auto& entry : m_WAB.lpEntryID)
			//{
			//	szEntryId += entry.ToString();
			//}
		}
		else if (eidtMessageDatabase == m_ObjectType)
		{
			auto szVersion = interpretprop::InterpretFlags(flagMDBVersion, m_MessageDatabaseObject.Version.getData());
			addBlock(
				m_MessageDatabaseObject.Version,
				L"\r\nVersion = 0x%1!02X! = %2!ws!\r\n",
				m_MessageDatabaseObject.Version.getData(),
				szVersion.c_str());

			auto szFlag = interpretprop::InterpretFlags(flagMDBFlag, m_MessageDatabaseObject.Flag.getData());
			addBlock(
				m_MessageDatabaseObject.Flag,
				L"Flag = 0x%1!02X! = %2!ws!\r\n",
				m_MessageDatabaseObject.Flag.getData(),
				szFlag.c_str());

			addBlock(
				m_MessageDatabaseObject.DLLFileName,
				L"DLLFileName = %1!hs!",
				m_MessageDatabaseObject.DLLFileName.getData().c_str());
			if (m_MessageDatabaseObject.bIsExchange)
			{
				addHeader(L"\r\n");

				auto szWrappedType =
					InterpretNumberAsStringProp(m_MessageDatabaseObject.WrappedType.getData(), PR_PROFILE_OPEN_FLAGS);
				addBlock(
					m_MessageDatabaseObject.WrappedFlags,
					L"Wrapped Flags = 0x%1!08X!\r\n",
					m_MessageDatabaseObject.WrappedFlags.getData());
				addBlock(
					m_MessageDatabaseObject.WrappedProviderUID,
					L"WrappedProviderUID = %1!ws!\r\n",
					guid::GUIDToStringAndName(m_MessageDatabaseObject.WrappedProviderUID.getData()).c_str());
				addBlock(
					m_MessageDatabaseObject.WrappedType,
					L"WrappedType = 0x%1!08X! = %2!ws!\r\n",
					m_MessageDatabaseObject.WrappedType.getData(),
					szWrappedType.c_str());
				addBlock(
					m_MessageDatabaseObject.ServerShortname,
					L"ServerShortname = %1!hs!\r\n",
					m_MessageDatabaseObject.ServerShortname.getData().c_str());
				addBlock(
					m_MessageDatabaseObject.MailboxDN,
					L"MailboxDN = %1!hs!",
					m_MessageDatabaseObject.MailboxDN.getData().c_str());
			}

			switch (m_MessageDatabaseObject.MagicVersion.getData())
			{
			case MDB_STORE_EID_V2_MAGIC:
			{
				addHeader(L"\r\n");
				auto szV2Magic =
					interpretprop::InterpretFlags(flagEidMagic, m_MessageDatabaseObject.v2.ulMagic.getData());
				addBlock(
					m_MessageDatabaseObject.v2.ulMagic,
					L"Magic = 0x%1!08X! = %2!ws!\r",
					m_MessageDatabaseObject.v2.ulMagic.getData(),
					szV2Magic.c_str());
				addBlock(
					m_MessageDatabaseObject.v2.ulSize,
					L"Size = 0x%1!08X! = %1!d!\r\n",
					m_MessageDatabaseObject.v2.ulSize.getData());
				auto szV2Version =
					interpretprop::InterpretFlags(flagEidVersion, m_MessageDatabaseObject.v2.ulVersion.getData());
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
				addBlock(
					m_MessageDatabaseObject.v2DN, L"DN = %1!hs!\r\n", m_MessageDatabaseObject.v2DN.getData().c_str());
				addBlock(
					m_MessageDatabaseObject.v2FQDN,
					L"FQDN = %1!ws!\r\n",
					m_MessageDatabaseObject.v2FQDN.getData().c_str());

				addHeader(L"Reserved Bytes = ");
				addBlockBytes(m_MessageDatabaseObject.v2Reserved);
			}
			break;
			case MDB_STORE_EID_V3_MAGIC:
			{
				addHeader(L"\r\n");
				auto szV3Magic =
					interpretprop::InterpretFlags(flagEidMagic, m_MessageDatabaseObject.v3.ulMagic.getData());
				addBlock(
					m_MessageDatabaseObject.v3.ulMagic,
					L"\r\nMagic = 0x%1!08X! = %2!ws!\r\n",
					m_MessageDatabaseObject.v3.ulMagic.getData(),
					szV3Magic.c_str());
				addBlock(
					m_MessageDatabaseObject.v3.ulSize,
					L"Size = 0x%1!08X! = %1!d!\r\n",
					m_MessageDatabaseObject.v3.ulSize.getData());
				auto szV3Version =
					interpretprop::InterpretFlags(flagEidVersion, m_MessageDatabaseObject.v3.ulVersion.getData());
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
					m_MessageDatabaseObject.v3SmtpAddress.getData().c_str());
				addHeader(L"Reserved Bytes = ");

				addBlockBytes(m_MessageDatabaseObject.v2Reserved);
			}
			break;
			}
		}
		else if (eidtFolder == m_ObjectType)
		{
			addHeader(L"\r\n");
			auto szType =
				interpretprop::InterpretFlags(flagMessageDatabaseObjectType, m_FolderOrMessage.Type.getData());
			addBlock(
				m_FolderOrMessage.Type,
				L"Folder Type = 0x%1!04X! = %2!ws!\r\n",
				m_FolderOrMessage.Type.getData(),
				szType.c_str());
			addBlock(
				m_FolderOrMessage.FolderObject.DatabaseGUID,
				L"Database GUID = %1!ws!\r\n",
				guid::GUIDToStringAndName(m_FolderOrMessage.FolderObject.DatabaseGUID.getData()).c_str());
			addHeader(L"GlobalCounter = ");
			addBlockBytes(m_FolderOrMessage.FolderObject.GlobalCounter);

			addHeader(L"\r\n");
			addHeader(L"Pad = ");
			addBlockBytes(m_FolderOrMessage.FolderObject.Pad);
		}
		else if (eidtMessage == m_ObjectType)
		{
			addHeader(L"\r\n");
			auto szType =
				interpretprop::InterpretFlags(flagMessageDatabaseObjectType, m_FolderOrMessage.Type.getData());
			addBlock(
				m_FolderOrMessage.Type,
				L"Message Type = 0x%1!04X! = %2!ws!\r\n",
				m_FolderOrMessage.Type.getData(),
				szType.c_str());
			addBlock(
				m_FolderOrMessage.FolderObject.DatabaseGUID,
				L"Folder Database GUID = %1!ws!\r\n",
				guid::GUIDToStringAndName(m_FolderOrMessage.FolderObject.DatabaseGUID.getData()).c_str());
			addHeader(L"Folder GlobalCounter = ");
			addBlockBytes(m_FolderOrMessage.FolderObject.GlobalCounter);

			addHeader(L"\r\n");
			addHeader(L"Pad1 = ");
			addBlockBytes(m_FolderOrMessage.MessageObject.Pad1);

			addHeader(L"\r\n");
			addBlock(
				m_FolderOrMessage.MessageObject.MessageDatabaseGUID,
				L"Message Database GUID = %1!ws!\r\n",
				guid::GUIDToStringAndName(m_FolderOrMessage.MessageObject.MessageDatabaseGUID.getData()).c_str());
			addBlockBytes(m_FolderOrMessage.MessageObject.MessageGlobalCounter);

			addHeader(L"\r\n");
			addHeader(L"Pad2 = ");
			addBlockBytes(m_FolderOrMessage.MessageObject.Pad2);
		}
	}
}