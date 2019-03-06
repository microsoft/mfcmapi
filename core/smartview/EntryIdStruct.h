#pragma once
#include <core/smartview/SmartViewParser.h>

namespace smartview
{
	enum EIDStructType
	{
		eidtUnknown = 0,
		eidtEphemeral,
		eidtShortTerm,
		eidtFolder,
		eidtMessage,
		eidtMessageDatabase,
		eidtOneOff,
		eidtAddressBook,
		eidtContact,
		eidtNewsGroupFolder,
		eidtWAB,
	};

	struct MDB_STORE_EID_V2
	{
		static const size_t size = sizeof(ULONG) * 5;
		blockT<ULONG> ulMagic; // MDB_STORE_EID_V2_MAGIC
		blockT<ULONG> ulSize; // size of this struct plus the size of szServerDN and wszServerFQDN
		blockT<ULONG> ulVersion; // MDB_STORE_EID_V2_VERSION
		blockT<ULONG> ulOffsetDN; // offset past the beginning of the MDB_STORE_EID_V2 struct where szServerDN starts
		blockT<ULONG>
			ulOffsetFQDN; // offset past the beginning of the MDB_STORE_EID_V2 struct where wszServerFQDN starts
	};

	struct MDB_STORE_EID_V3
	{
		static const size_t size = sizeof(ULONG) * 4;
		blockT<ULONG> ulMagic; // MDB_STORE_EID_V3_MAGIC
		blockT<ULONG> ulSize; // size of this struct plus the size of szServerSmtpAddress
		blockT<ULONG> ulVersion; // MDB_STORE_EID_V3_VERSION
		blockT<ULONG>
			ulOffsetSmtpAddress; // offset past the beginning of the MDB_STORE_EID_V3 struct where szSmtpAddress starts
	};

	struct FolderObject
	{
		blockT<GUID> DatabaseGUID;
		blockBytes GlobalCounter; // 6 bytes
		blockBytes Pad; // 2 bytes
	};

	struct MessageObject
	{
		blockT<GUID> FolderDatabaseGUID;
		blockBytes FolderGlobalCounter; // 6 bytes
		blockBytes Pad1; // 2 bytes
		blockT<GUID> MessageDatabaseGUID;
		blockBytes MessageGlobalCounter; // 6 bytes
		blockBytes Pad2; // 2 bytes
	};

	struct FolderOrMessage
	{
		blockT<WORD> Type;
		FolderObject FolderObject;
		MessageObject MessageObject;
	};

	struct MessageDatabaseObject
	{
		blockT<BYTE> Version;
		blockT<BYTE> Flag;
		blockStringA DLLFileName;
		bool bIsExchange{};
		blockT<ULONG> WrappedFlags;
		blockT<GUID> WrappedProviderUID;
		blockT<ULONG> WrappedType;
		blockStringA ServerShortname;
		blockStringA MailboxDN;
		blockT<ULONG> MagicVersion;
		MDB_STORE_EID_V2 v2;
		MDB_STORE_EID_V3 v3;
		blockStringA v2DN;
		blockStringW v2FQDN;
		blockStringW v3SmtpAddress;
		blockBytes v2Reserved; // 2 bytes
	};

	struct EphemeralObject
	{
		blockT<ULONG> Version;
		blockT<ULONG> Type;
	};

	struct OneOffRecipientObject
	{
		blockT<DWORD> Bitmask;
		struct Unicode
		{
			blockStringW DisplayName;
			blockStringW AddressType;
			blockStringW EmailAddress;
		} Unicode;
		struct ANSI
		{
			blockStringA DisplayName;
			blockStringA AddressType;
			blockStringA EmailAddress;
		} ANSI;
	};

	struct AddressBookObject
	{
		blockT<DWORD> Version;
		blockT<DWORD> Type;
		blockStringA X500DN;
	};

	class EntryIdStruct;

	struct ContactAddressBookObject
	{
		blockT<DWORD> Version;
		blockT<DWORD> Type;
		blockT<DWORD> Index; // CONTAB_USER, CONTAB_DISTLIST only
		blockT<DWORD> EntryIDCount; // CONTAB_USER, CONTAB_DISTLIST only
		blockT<GUID> muidID; // CONTAB_CONTAINER only
		std::shared_ptr<EntryIdStruct> lpEntryID;
	};

	struct WAB
	{
		blockT<BYTE> Type;
		std::shared_ptr<EntryIdStruct> lpEntryID;
	};

	class EntryIdStruct : public SmartViewParser
	{
	public:
		EntryIdStruct() = default;
		EntryIdStruct(std::shared_ptr<binaryParser> binaryParser, bool bEnableJunk)
		{
			parse(binaryParser, 0, bEnableJunk);
		}
		EntryIdStruct(std::shared_ptr<binaryParser> binaryParser, size_t cbBin, bool bEnableJunk)
		{
			parse(binaryParser, cbBin, bEnableJunk);
		}

	private:
		void Parse() override;
		void ParseBlocks() override;

		blockT<byte> m_abFlags0;
		blockT<byte> m_abFlags1;
		blockBytes m_abFlags23; // 2 bytes
		blockT<GUID> m_ProviderUID;
		EIDStructType m_ObjectType = eidtUnknown; // My own addition to simplify parsing
		FolderOrMessage m_FolderOrMessage;
		MessageDatabaseObject m_MessageDatabaseObject;
		EphemeralObject m_EphemeralObject{};
		OneOffRecipientObject m_OneOffRecipientObject{};
		AddressBookObject m_AddressBookObject;
		ContactAddressBookObject m_ContactAddressBookObject;
		WAB m_WAB;
	};
} // namespace smartview