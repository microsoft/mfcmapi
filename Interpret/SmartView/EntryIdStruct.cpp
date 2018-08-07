#include <StdAfx.h>
#include <Interpret/SmartView/EntryIdStruct.h>
#include <Interpret/SmartView/SmartView.h>
#include <Interpret/String.h>
#include <Interpret/Guids.h>
#include <Interpret/InterpretProp.h>
#include <Interpret/ExtraPropTags.h>

namespace smartview
{
	EntryIdStruct::EntryIdStruct()
		: m_FolderOrMessage(), m_MessageDatabaseObject(), m_AddressBookObject(), m_ContactAddressBookObject(), m_WAB()
	{
		m_ObjectType = eidtUnknown;
	}

	void EntryIdStruct::Parse()
	{
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

		if (!m_abFlags.getSize()) return szEntryId;
		if (0 == (m_abFlags.getData()[0] | m_abFlags.getData()[1] | m_abFlags.getData()[2] | m_abFlags.getData()[3]))
		{
			szEntryId += strings::formatmessage(IDS_ENTRYIDHEADER1);
		}
		else if (0 == (m_abFlags.getData()[1] | m_abFlags.getData()[2] | m_abFlags.getData()[3]))
		{
			szEntryId += strings::formatmessage(
				IDS_ENTRYIDHEADER2,
				m_abFlags.getData()[0],
				interpretprop::InterpretFlags(flagEntryId0, m_abFlags.getData()[0]).c_str());
		}
		else
		{
			szEntryId += strings::formatmessage(
				IDS_ENTRYIDHEADER3,
				m_abFlags.getData()[0],
				interpretprop::InterpretFlags(flagEntryId0, m_abFlags.getData()[0]).c_str(),
				m_abFlags.getData()[1],
				interpretprop::InterpretFlags(flagEntryId1, m_abFlags.getData()[1]).c_str(),
				m_abFlags.getData()[2],
				m_abFlags.getData()[3]);
		}

		auto szGUID = guid::GUIDToStringAndName(m_ProviderUID.getData());
		szEntryId += strings::formatmessage(IDS_ENTRYIDPROVIDERGUID, szGUID.c_str());

		if (eidtEphemeral == m_ObjectType)
		{
			auto szVersion = interpretprop::InterpretFlags(flagExchangeABVersion, m_EphemeralObject.Version.getData());
			auto szType = InterpretNumberAsStringProp(m_EphemeralObject.Type.getData(), PR_DISPLAY_TYPE);

			szEntryId += strings::formatmessage(
				IDS_ENTRYIDEPHEMERALADDRESSDATA,
				m_EphemeralObject.Version.getData(),
				szVersion.c_str(),
				m_EphemeralObject.Type.getData(),
				szType.c_str());
		}
		else if (eidtOneOff == m_ObjectType)
		{
			auto szFlags = interpretprop::InterpretFlags(flagOneOffEntryId, m_OneOffRecipientObject.Bitmask.getData());
			if (MAPI_UNICODE & m_OneOffRecipientObject.Bitmask.getData())
			{
				szEntryId += strings::formatmessage(
					IDS_ONEOFFENTRYIDFOOTERUNICODE,
					m_OneOffRecipientObject.Bitmask.getData(),
					szFlags.c_str(),
					m_OneOffRecipientObject.Unicode.DisplayName.getData().c_str(),
					m_OneOffRecipientObject.Unicode.AddressType.getData().c_str(),
					m_OneOffRecipientObject.Unicode.EmailAddress.getData().c_str());
			}
			else
			{
				szEntryId += strings::formatmessage(
					IDS_ONEOFFENTRYIDFOOTERANSI,
					m_OneOffRecipientObject.Bitmask.getData(),
					szFlags.c_str(),
					m_OneOffRecipientObject.ANSI.DisplayName.getData().c_str(),
					m_OneOffRecipientObject.ANSI.AddressType.getData().c_str(),
					m_OneOffRecipientObject.ANSI.EmailAddress.getData().c_str());
			}
		}
		else if (eidtAddressBook == m_ObjectType)
		{
			auto szVersion =
				interpretprop::InterpretFlags(flagExchangeABVersion, m_AddressBookObject.Version.getData());
			auto szType = InterpretNumberAsStringProp(m_AddressBookObject.Type.getData(), PR_DISPLAY_TYPE);

			szEntryId += strings::formatmessage(
				IDS_ENTRYIDEXCHANGEADDRESSDATA,
				m_AddressBookObject.Version.getData(),
				szVersion.c_str(),
				m_AddressBookObject.Type.getData(),
				szType.c_str(),
				m_AddressBookObject.X500DN.getData().c_str());
		}
		// Contact Address Book / Personal Distribution List (PDL)
		else if (eidtContact == m_ObjectType)
		{
			auto szVersion =
				interpretprop::InterpretFlags(flagContabVersion, m_ContactAddressBookObject.Version.getData());
			auto szType = interpretprop::InterpretFlags(flagContabType, m_ContactAddressBookObject.Type.getData());

			szEntryId += strings::formatmessage(
				IDS_ENTRYIDCONTACTADDRESSDATAHEAD,
				m_ContactAddressBookObject.Version.getData(),
				szVersion.c_str(),
				m_ContactAddressBookObject.Type.getData(),
				szType.c_str());

			switch (m_ContactAddressBookObject.Type.getData())
			{
			case CONTAB_USER:
			case CONTAB_DISTLIST:
				szEntryId += strings::formatmessage(
					IDS_ENTRYIDCONTACTADDRESSDATAUSER,
					m_ContactAddressBookObject.Index.getData(),
					interpretprop::InterpretFlags(flagContabIndex, m_ContactAddressBookObject.Index.getData()).c_str(),
					m_ContactAddressBookObject.EntryIDCount.getData());
				break;
			case CONTAB_ROOT:
			case CONTAB_CONTAINER:
			case CONTAB_SUBROOT:
				szGUID = guid::GUIDToStringAndName(m_ContactAddressBookObject.muidID.getData());
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
			szEntryId += strings::formatmessage(
				IDS_ENTRYIDWRAPPEDENTRYIDDATA,
				m_WAB.Type.getData(),
				interpretprop::InterpretFlags(flagWABEntryIDType, m_WAB.Type.getData()).c_str());

			for (auto& entry : m_WAB.lpEntryID)
			{
				szEntryId += entry.ToString();
			}
		}
		else if (eidtMessageDatabase == m_ObjectType)
		{
			auto szVersion = interpretprop::InterpretFlags(flagMDBVersion, m_MessageDatabaseObject.Version.getData());
			auto szFlag = interpretprop::InterpretFlags(flagMDBFlag, m_MessageDatabaseObject.Flag.getData());

			szEntryId += strings::formatmessage(
				IDS_ENTRYIDMAPIMESSAGESTOREDATA,
				m_MessageDatabaseObject.Version.getData(),
				szVersion.c_str(),
				m_MessageDatabaseObject.Flag.getData(),
				szFlag.c_str(),
				m_MessageDatabaseObject.DLLFileName.getData().c_str());
			if (m_MessageDatabaseObject.bIsExchange)
			{
				auto szWrappedType =
					InterpretNumberAsStringProp(m_MessageDatabaseObject.WrappedType.getData(), PR_PROFILE_OPEN_FLAGS);

				szGUID = guid::GUIDToStringAndName(m_MessageDatabaseObject.WrappedProviderUID.getData());
				szEntryId += strings::formatmessage(
					IDS_ENTRYIDMAPIMESSAGESTOREEXCHANGEDATA,
					m_MessageDatabaseObject.WrappedFlags.getData(),
					szGUID.c_str(),
					m_MessageDatabaseObject.WrappedType.getData(),
					szWrappedType.c_str(),
					m_MessageDatabaseObject.ServerShortname.getData().c_str(),
					m_MessageDatabaseObject.MailboxDN.getData().c_str());
			}

			switch (m_MessageDatabaseObject.MagicVersion.getData())
			{
			case MDB_STORE_EID_V2_MAGIC:
			{
				auto szV2Magic =
					interpretprop::InterpretFlags(flagEidMagic, m_MessageDatabaseObject.v2.ulMagic.getData());
				auto szV2Version =
					interpretprop::InterpretFlags(flagEidVersion, m_MessageDatabaseObject.v2.ulVersion.getData());

				szEntryId += strings::formatmessage(
					IDS_ENTRYIDMAPIMESSAGESTOREEXCHANGEDATAV2,
					m_MessageDatabaseObject.v2.ulMagic.getData(),
					szV2Magic.c_str(),
					m_MessageDatabaseObject.v2.ulSize.getData(),
					m_MessageDatabaseObject.v2.ulVersion.getData(),
					szV2Version.c_str(),
					m_MessageDatabaseObject.v2.ulOffsetDN.getData(),
					m_MessageDatabaseObject.v2.ulOffsetFQDN.getData(),
					m_MessageDatabaseObject.v2DN.getData().c_str(),
					m_MessageDatabaseObject.v2FQDN.getData().c_str());

				szEntryId += strings::BinToHexString(m_MessageDatabaseObject.v2Reserved.getData(), true);
			}
			break;
			case MDB_STORE_EID_V3_MAGIC:
			{
				auto szV3Magic =
					interpretprop::InterpretFlags(flagEidMagic, m_MessageDatabaseObject.v3.ulMagic.getData());
				auto szV3Version =
					interpretprop::InterpretFlags(flagEidVersion, m_MessageDatabaseObject.v3.ulVersion.getData());

				szEntryId += strings::formatmessage(
					IDS_ENTRYIDMAPIMESSAGESTOREEXCHANGEDATAV3,
					m_MessageDatabaseObject.v3.ulMagic.getData(),
					szV3Magic.c_str(),
					m_MessageDatabaseObject.v3.ulSize.getData(),
					m_MessageDatabaseObject.v3.ulVersion.getData(),
					szV3Version.c_str(),
					m_MessageDatabaseObject.v3.ulOffsetSmtpAddress.getData(),
					m_MessageDatabaseObject.v3SmtpAddress.getData().c_str());

				szEntryId += strings::BinToHexString(m_MessageDatabaseObject.v2Reserved.getData(), true);
			}
			break;
			}
		}
		else if (eidtFolder == m_ObjectType)
		{
			auto szType =
				interpretprop::InterpretFlags(flagMessageDatabaseObjectType, m_FolderOrMessage.Type.getData());

			auto szDatabaseGUID = guid::GUIDToStringAndName(m_FolderOrMessage.FolderObject.DatabaseGUID.getData());

			szEntryId += strings::formatmessage(
				IDS_ENTRYIDEXCHANGEFOLDERDATA,
				m_FolderOrMessage.Type.getData(),
				szType.c_str(),
				szDatabaseGUID.c_str());
			szEntryId += strings::BinToHexString(m_FolderOrMessage.FolderObject.GlobalCounter.getData(), true);

			szEntryId += strings::formatmessage(IDS_ENTRYIDEXCHANGEDATAPAD);
			szEntryId += strings::BinToHexString(m_FolderOrMessage.FolderObject.Pad.getData(), true);
		}
		else if (eidtMessage == m_ObjectType)
		{
			auto szType =
				interpretprop::InterpretFlags(flagMessageDatabaseObjectType, m_FolderOrMessage.Type.getData());
			auto szFolderDatabaseGUID =
				guid::GUIDToStringAndName(m_FolderOrMessage.MessageObject.FolderDatabaseGUID.getData());
			auto szMessageDatabaseGUID =
				guid::GUIDToStringAndName(m_FolderOrMessage.MessageObject.MessageDatabaseGUID.getData());

			szEntryId += strings::formatmessage(
				IDS_ENTRYIDEXCHANGEMESSAGEDATA,
				m_FolderOrMessage.Type.getData(),
				szType.c_str(),
				szFolderDatabaseGUID.c_str());
			szEntryId += strings::BinToHexString(m_FolderOrMessage.MessageObject.FolderGlobalCounter.getData(), true);

			szEntryId += strings::formatmessage(IDS_ENTRYIDEXCHANGEDATAPADNUM, 1);
			szEntryId += strings::BinToHexString(m_FolderOrMessage.MessageObject.Pad1.getData(), true);

			szEntryId += strings::formatmessage(IDS_ENTRYIDEXCHANGEMESSAGEDATAGUID, szMessageDatabaseGUID.c_str());
			szEntryId += strings::BinToHexString(m_FolderOrMessage.MessageObject.MessageGlobalCounter.getData(), true);

			szEntryId += strings::formatmessage(IDS_ENTRYIDEXCHANGEDATAPADNUM, 2);
			szEntryId += strings::BinToHexString(m_FolderOrMessage.MessageObject.Pad2.getData(), true);
		}

		return szEntryId;
	}
}