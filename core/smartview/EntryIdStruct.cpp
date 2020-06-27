#include <core/stdafx.h>
#include <core/smartview/EntryIdStruct.h>
#include <core/smartview/SmartView.h>
#include <core/interpret/guid.h>
#include <core/interpret/flags.h>
#include <core/mapi/extraPropTags.h>

namespace smartview
{
	void MDB_STORE_EID_V2::parse()
	{
		ulMagic = blockT<DWORD>::parse(parser);
		ulSize = blockT<DWORD>::parse(parser);
		ulVersion = blockT<DWORD>::parse(parser);
		ulOffsetDN = blockT<DWORD>::parse(parser);
		ulOffsetFQDN = blockT<DWORD>::parse(parser);
		if (*ulOffsetDN)
		{
			v2DN = blockStringA::parse(parser);
		}

		if (*ulOffsetFQDN)
		{
			v2FQDN = blockStringW::parse(parser);
		}

		v2Reserved = blockBytes::parse(parser, 2);
	}

	void MDB_STORE_EID_V2::parseBlocks()
	{
		setText(L"MAPI Message Store Entry ID V2");
		auto szV2Magic = flags::InterpretFlags(flagEidMagic, *ulMagic);
		addChild(ulMagic, L"Magic = 0x%1!08X! = %2!ws!", ulMagic->getData(), szV2Magic.c_str());
		addChild(ulSize, L"Size = 0x%1!08X! = %1!d!", ulSize->getData());
		auto szV2Version = flags::InterpretFlags(flagEidVersion, *ulVersion);
		addChild(ulVersion, L"Version = 0x%1!08X! = %2!ws!", ulVersion->getData(), szV2Version.c_str());
		addChild(ulOffsetDN, L"OffsetDN = 0x%1!08X!", ulOffsetDN->getData());
		addChild(ulOffsetFQDN, L"OffsetFQDN = 0x%1!08X!", ulOffsetFQDN->getData());
		addChild(v2DN, L"DN = %1!hs!", v2DN->c_str());
		addChild(v2FQDN, L"FQDN = %1!ws!", v2FQDN->c_str());

		addLabeledChild(L"Reserved Bytes =", v2Reserved);
	}

	void MDB_STORE_EID_V3::parse()
	{
		ulMagic = blockT<DWORD>::parse(parser);
		ulSize = blockT<DWORD>::parse(parser);
		ulVersion = blockT<DWORD>::parse(parser);
		ulOffsetSmtpAddress = blockT<DWORD>::parse(parser);
		if (ulOffsetSmtpAddress)
		{
			v3SmtpAddress = blockStringW::parse(parser);
		}

		v2Reserved = blockBytes::parse(parser, 2);
	}

	void MDB_STORE_EID_V3::parseBlocks()
	{
		setText(L"MAPI Message Store Entry ID V3");
		auto szV3Magic = flags::InterpretFlags(flagEidMagic, *ulMagic);
		addChild(ulMagic, L"Magic = 0x%1!08X! = %2!ws!", ulMagic->getData(), szV3Magic.c_str());
		addChild(ulSize, L"Size = 0x%1!08X! = %1!d!", ulSize->getData());
		auto szV3Version = flags::InterpretFlags(flagEidVersion, *ulVersion);
		addChild(ulVersion, L"Version = 0x%1!08X! = %2!ws!", ulVersion->getData(), szV3Version.c_str());
		addChild(ulOffsetSmtpAddress, L"OffsetSmtpAddress = 0x%1!08X!", ulOffsetSmtpAddress->getData());
		addChild(v3SmtpAddress, L"SmtpAddress = %1!ws!", v3SmtpAddress->c_str());

		addLabeledChild(L"Reserved Bytes =", v2Reserved);
	}

	void FolderObject::parse()
	{
		Type = blockT<WORD>::parse(parser);
		DatabaseGUID = blockT<GUID>::parse(parser);
		GlobalCounter = blockBytes::parse(parser, 6);
		Pad = blockBytes::parse(parser, 2);
	}

	void FolderObject::parseBlocks()
	{
		setText(L"Exchange Folder Entry ID");
		auto szType = flags::InterpretFlags(flagMessageDatabaseObjectType, *Type);
		addChild(Type, L"Folder Type = 0x%1!04X! = %2!ws!", Type->getData(), szType.c_str());
		addChild(DatabaseGUID, L"Database GUID = %1!ws!", guid::GUIDToStringAndName(*DatabaseGUID).c_str());
		addLabeledChild(L"GlobalCounter =", GlobalCounter);

		addLabeledChild(L"Pad =", Pad);
	}

	void MessageObject ::parse()
	{
		Type = blockT<WORD>::parse(parser);
		FolderDatabaseGUID = blockT<GUID>::parse(parser);
		FolderGlobalCounter = blockBytes::parse(parser, 6);
		Pad1 = blockBytes::parse(parser, 2);

		MessageDatabaseGUID = blockT<GUID>::parse(parser);
		MessageGlobalCounter = blockBytes::parse(parser, 6);
		Pad2 = blockBytes::parse(parser, 2);
	}

	void MessageObject ::parseBlocks()
	{
		setText(L"Exchange Message Entry ID");
		auto szType = flags::InterpretFlags(flagMessageDatabaseObjectType, *Type);
		addChild(Type, L"Message Type = 0x%1!04X! = %2!ws!", Type->getData(), szType.c_str());
		addChild(
			FolderDatabaseGUID,
			L"Folder Database GUID = %1!ws!",
			guid::GUIDToStringAndName(*FolderDatabaseGUID).c_str());
		addLabeledChild(L"Folder GlobalCounter =", FolderGlobalCounter);

		addLabeledChild(L"Pad1 =", Pad1);

		addChild(
			MessageDatabaseGUID,
			L"Message Database GUID = %1!ws!",
			guid::GUIDToStringAndName(*MessageDatabaseGUID).c_str());
		addLabeledChild(L"Message GlobalCounter =", MessageGlobalCounter);

		addLabeledChild(L"Pad2 =", Pad2);
	}

	void MessageDatabaseObject ::parse()
	{
		Version = blockT<byte>::parse(parser);
		Flag = blockT<byte>::parse(parser);
		DLLFileName = blockStringA::parse(parser);
		bIsExchange = false;

		// We only know how to parse emsmdb.dll's wrapped entry IDs
		if (!DLLFileName->empty() && CSTR_EQUAL == CompareStringA(
													   LOCALE_INVARIANT,
													   NORM_IGNORECASE,
													   DLLFileName->c_str(),
													   -1,
													   "emsmdb.dll", // STRING_OK
													   -1))
		{
			bIsExchange = true;
			auto cbRead = parser->getOffset();
			// Advance to the next multiple of 4
			parser->advance(3 - (cbRead + 3) % 4);
			WrappedFlags = blockT<DWORD>::parse(parser);
			WrappedProviderUID = blockT<GUID>::parse(parser);
			WrappedType = blockT<DWORD>::parse(parser);
			ServerShortname = blockStringA::parse(parser);

			// Test if we have a magic value. Some PF EIDs also have a mailbox DN and we need to accomodate them
			if (*WrappedType & OPENSTORE_PUBLIC)
			{
				cbRead = parser->getOffset();
				MagicVersion = blockT<DWORD>::parse(parser);
				parser->setOffset(cbRead);
			}
			else
			{
				MagicVersion->setOffset(0);
				MagicVersion->setSize(0);
				MagicVersion->setData(MDB_STORE_EID_V1_VERSION);
			}

			// Either we're not a PF eid, or this PF EID wasn't followed directly by a magic value
			if (!(*WrappedType & OPENSTORE_PUBLIC) ||
				*MagicVersion != MDB_STORE_EID_V2_MAGIC && *MagicVersion != MDB_STORE_EID_V3_MAGIC)
			{
				MailboxDN = blockStringA::parse(parser);
			}

			// Check again for a magic value
			cbRead = parser->getOffset();
			MagicVersion = blockT<DWORD>::parse(parser);
			parser->setOffset(cbRead);

			switch (*MagicVersion)
			{
			case MDB_STORE_EID_V2_MAGIC:
				if (parser->getSize() >= MDB_STORE_EID_V2::size + sizeof(WCHAR))
				{
					v2 = block::parse<MDB_STORE_EID_V2>(parser, 0, false);
				}
				break;
			case MDB_STORE_EID_V3_MAGIC:
				if (parser->getSize() >= MDB_STORE_EID_V3::size + sizeof(WCHAR))
				{
					v3 = block::parse<MDB_STORE_EID_V3>(parser, 0, false);
				}
				break;
			}
		}
	}

	void MessageDatabaseObject ::parseBlocks()
	{
		setText(L"MAPI Message Store Entry ID");
		auto szVersion = flags::InterpretFlags(flagMDBVersion, *Version);
		addChild(Version, L"Version = 0x%1!02X! = %2!ws!", Version->getData(), szVersion.c_str());

		auto szFlag = flags::InterpretFlags(flagMDBFlag, *Flag);
		addChild(Flag, L"Flag = 0x%1!02X! = %2!ws!", Flag->getData(), szFlag.c_str());

		addChild(DLLFileName, L"DLLFileName = %1!hs!", DLLFileName->c_str());
		if (bIsExchange)
		{
			auto szWrappedType = InterpretNumberAsStringProp(*WrappedType, PR_PROFILE_OPEN_FLAGS);
			addChild(WrappedFlags, L"Wrapped Flags = 0x%1!08X!", WrappedFlags->getData());
			addChild(
				WrappedProviderUID,
				L"WrappedProviderUID = %1!ws!",
				guid::GUIDToStringAndName(*WrappedProviderUID).c_str());
			addChild(WrappedType, L"WrappedType = 0x%1!08X! = %2!ws!", WrappedType->getData(), szWrappedType.c_str());
			addChild(ServerShortname, L"ServerShortname = %1!hs!", ServerShortname->c_str());
			addChild(MailboxDN, L"MailboxDN = %1!hs!", MailboxDN->c_str());
		}

		switch (*MagicVersion)
		{
		case MDB_STORE_EID_V2_MAGIC:
			addChild(v2);
			break;
		case MDB_STORE_EID_V3_MAGIC:
		{
			addChild(v3);
		}
		break;
		}
	}

	void EphemeralObject ::parse()
	{
		Version = blockT<DWORD>::parse(parser);
		Type = blockT<DWORD>::parse(parser);
	}

	void EphemeralObject ::parseBlocks()
	{
		setText(L"Ephemeral Entry ID");
		auto szVersion = flags::InterpretFlags(flagExchangeABVersion, *Version);
		addChild(Version, L"Version = 0x%1!08X! = %2!ws!", Version->getData(), szVersion.c_str());

		auto szType = InterpretNumberAsStringProp(*Type, PR_DISPLAY_TYPE);
		addChild(Type, L"Type = 0x%1!08X! = %2!ws!", Type->getData(), szType.c_str());
	}

	void OneOffRecipientObject ::parse()
	{
		Bitmask = blockT<DWORD>::parse(parser);
		if (*Bitmask & MAPI_UNICODE)
		{
			DisplayNameW = blockStringW::parse(parser);
			AddressTypeW = blockStringW::parse(parser);
			EmailAddressW = blockStringW::parse(parser);
		}
		else
		{
			DisplayNameA = blockStringA::parse(parser);
			AddressTypeA = blockStringA::parse(parser);
			EmailAddressA = blockStringA::parse(parser);
		}
	}

	void OneOffRecipientObject ::parseBlocks()
	{
		setText(L"MAPI One Off Recipient Entry ID");
		auto szFlags = flags::InterpretFlags(flagOneOffEntryId, *Bitmask);
		if (*Bitmask & MAPI_UNICODE)
		{
			addChild(Bitmask, L"dwBitmask: 0x%1!08X! = %2!ws!", Bitmask->getData(), szFlags.c_str());
			addChild(DisplayNameW, L"szDisplayName = %1!ws!", DisplayNameW->c_str());
			addChild(AddressTypeW, L"szAddressType = %1!ws!", AddressTypeW->c_str());
			addChild(EmailAddressW, L"szEmailAddress = %1!ws!", EmailAddressW->c_str());
		}
		else
		{
			addChild(Bitmask, L"dwBitmask: 0x%1!08X! = %2!ws!", Bitmask->getData(), szFlags.c_str());
			addChild(DisplayNameA, L"szDisplayName = %1!hs!", DisplayNameA->c_str());
			addChild(AddressTypeA, L"szAddressType = %1!hs!", AddressTypeA->c_str());
			addChild(EmailAddressA, L"szEmailAddress = %1!hs!", EmailAddressA->c_str());
		}
	}

	void AddressBookObject ::parse()
	{
		Version = blockT<DWORD>::parse(parser);
		Type = blockT<DWORD>::parse(parser);
		X500DN = blockStringA::parse(parser);
	}

	void AddressBookObject ::parseBlocks()
	{
		setText(L"Exchange Address Entry ID");
		auto szVersion = flags::InterpretFlags(flagExchangeABVersion, *Version);
		addChild(Version, L"Version = 0x%1!08X! = %2!ws!", Version->getData(), szVersion.c_str());
		auto szType = InterpretNumberAsStringProp(*Type, PR_DISPLAY_TYPE);
		addChild(Type, L"Type = 0x%1!08X! = %2!ws!", Type->getData(), szType.c_str());
		addChild(X500DN, L"X500DN = %1!hs!", X500DN->c_str());
	}

	void ContactAddressBookObject ::parse()
	{
		Version = blockT<DWORD>::parse(parser);
		Type = blockT<DWORD>::parse(parser);

		if (*Type == CONTAB_CONTAINER)
		{
			muidID = blockT<GUID>::parse(parser);
		}
		else // Assume we've got some variation on the user/distlist format
		{
			Index = blockT<DWORD>::parse(parser);
			EntryIDCount = blockT<DWORD>::parse(parser);
		}

		// Read the wrapped entry ID from the remaining data
		auto cbRemainingBytes = parser->getSize();

		// If we already got a size, use it, else we just read the rest of the structure
		if (EntryIDCount != 0 && *EntryIDCount < cbRemainingBytes)
		{
			cbRemainingBytes = *EntryIDCount;
		}

		lpEntryID = block::parse<EntryIdStruct>(parser, cbRemainingBytes, false);
	}

	void ContactAddressBookObject ::parseBlocks()
	{
		setText(L"Contact Address Book / PDL Entry ID");
		auto szVersion = flags::InterpretFlags(flagContabVersion, *Version);
		addChild(Version, L"Version = 0x%1!08X! = %2!ws!", Version->getData(), szVersion.c_str());

		auto szType = flags::InterpretFlags(flagContabType, *Type);
		addChild(Type, L"Type = 0x%1!08X! = %2!ws!", Type->getData(), szType.c_str());

		switch (*Type)
		{
		case CONTAB_USER:
		case CONTAB_DISTLIST:
			addChild(
				Index,
				L"Index = 0x%1!08X! = %2!ws!",
				Index->getData(),
				flags::InterpretFlags(flagContabIndex, *Index).c_str());
			addChild(EntryIDCount, L"EntryIDCount = 0x%1!08X!", EntryIDCount->getData());
			break;
		case CONTAB_ROOT:
		case CONTAB_CONTAINER:
		case CONTAB_SUBROOT:
			addChild(muidID, L"muidID = %1!ws!", guid::GUIDToStringAndName(*muidID).c_str());
			break;
		default:
			break;
		}

		addChild(lpEntryID);
	}

	void WAB::parse()
	{
		Type = blockT<byte>::parse(parser);
		lpEntryID = block::parse<EntryIdStruct>(parser, false);
	}

	void WAB::parseBlocks()
	{
		setText(L"Wrapped Entry ID");
		addChild(
			Type,
			L"Wrapped Entry Type = 0x%1!02X! = %2!ws!",
			Type->getData(),
			flags::InterpretFlags(flagWABEntryIDType, *Type).c_str());

		addChild(lpEntryID);
	}

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
				m_EphemeralObject = block::parse<EphemeralObject>(parser, 0, false);
				break;
			// One Off Recipient
			case EIDStructType::oneOff:
				m_OneOffRecipientObject = block::parse<OneOffRecipientObject>(parser, 0, false);
				break;
			// Address Book Recipient
			case EIDStructType::addressBook:
				m_AddressBookObject = block::parse<AddressBookObject>(parser, 0, false);
				break;
			// Contact Address Book / Personal Distribution List (PDL)
			case EIDStructType::contact:
				m_ContactAddressBookObject = block::parse<ContactAddressBookObject>(parser, 0, false);
				break;
			case EIDStructType::WAB:
				m_WAB = block::parse<WAB>(parser, 0, false);
				break;
			// message store objects
			case EIDStructType::messageDatabase:
				m_MessageDatabaseObject = block::parse<MessageDatabaseObject>(parser, 0, false);
				break;
			// Exchange message store folder
			case EIDStructType::folder:
				m_FolderObject = block::parse<FolderObject>(parser, 0, false);
				break;
			// Exchange message store message
			case EIDStructType::message:
				m_MessageObject = block::parse<MessageObject>(parser, 0, false);
				break;
			}
		}

		// Check if we have an unidentified short term entry ID:
		if (EIDStructType::unknown == m_ObjectType && *m_abFlags0 & MAPI_SHORTTERM)
			m_ObjectType = EIDStructType::shortTerm;
	}

	void EntryIdStruct::parseBlocks()
	{
		if (m_ObjectType == EIDStructType::shortTerm)
		{
			setText(L"Short Term Entry ID");
		}
		else
		{
			setText(L"Entry ID");
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
			addChild(m_EphemeralObject);
		}
		else if (EIDStructType::oneOff == m_ObjectType)
		{
			addChild(m_OneOffRecipientObject);
		}
		else if (EIDStructType::addressBook == m_ObjectType)
		{
			addChild(m_AddressBookObject);
		}
		else if (EIDStructType::contact == m_ObjectType)
		{
			addChild(m_ContactAddressBookObject);
		}
		else if (EIDStructType::WAB == m_ObjectType)
		{
			addChild(m_WAB);
		}
		else if (EIDStructType::messageDatabase == m_ObjectType)
		{
			addChild(m_MessageDatabaseObject);
		}
		else if (EIDStructType::folder == m_ObjectType)
		{
			addChild(m_FolderObject);
		}
		else if (EIDStructType::message == m_ObjectType)
		{
			addChild(m_MessageObject);
		}
	}
} // namespace smartview