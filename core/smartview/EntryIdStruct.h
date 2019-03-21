#pragma once
#include <core/smartview/SmartViewParser.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

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
		std::shared_ptr<blockT<ULONG>> ulMagic = emptyT<ULONG>(); // MDB_STORE_EID_V2_MAGIC
		std::shared_ptr<blockT<ULONG>> ulSize =
			emptyT<ULONG>(); // size of this struct plus the size of szServerDN and wszServerFQDN
		std::shared_ptr<blockT<ULONG>> ulVersion = emptyT<ULONG>(); // MDB_STORE_EID_V2_VERSION
		std::shared_ptr<blockT<ULONG>> ulOffsetDN =
			emptyT<ULONG>(); // offset past the beginning of the MDB_STORE_EID_V2 struct where szServerDN starts
		std::shared_ptr<blockT<ULONG>> ulOffsetFQDN =
			emptyT<ULONG>(); // offset past the beginning of the MDB_STORE_EID_V2 struct where wszServerFQDN starts
	};

	struct MDB_STORE_EID_V3
	{
		static const size_t size = sizeof(ULONG) * 4;
		std::shared_ptr<blockT<ULONG>> ulMagic = emptyT<ULONG>(); // MDB_STORE_EID_V3_MAGIC
		std::shared_ptr<blockT<ULONG>> ulSize =
			emptyT<ULONG>(); // size of this struct plus the size of szServerSmtpAddress
		std::shared_ptr<blockT<ULONG>> ulVersion = emptyT<ULONG>(); // MDB_STORE_EID_V3_VERSION
		std::shared_ptr<blockT<ULONG>> ulOffsetSmtpAddress =
			emptyT<ULONG>(); // offset past the beginning of the MDB_STORE_EID_V3 struct where szSmtpAddress starts
	};

	struct FolderObject
	{
		std::shared_ptr<blockT<GUID>> DatabaseGUID = emptyT<GUID>();
		std::shared_ptr<blockBytes> GlobalCounter = emptyBB(); // 6 bytes
		std::shared_ptr<blockBytes> Pad = emptyBB(); // 2 bytes
	};

	struct MessageObject
	{
		std::shared_ptr<blockT<GUID>> FolderDatabaseGUID = emptyT<GUID>();
		std::shared_ptr<blockBytes> FolderGlobalCounter = emptyBB(); // 6 bytes
		std::shared_ptr<blockBytes> Pad1 = emptyBB(); // 2 bytes
		std::shared_ptr<blockT<GUID>> MessageDatabaseGUID = emptyT<GUID>();
		std::shared_ptr<blockBytes> MessageGlobalCounter = emptyBB(); // 6 bytes
		std::shared_ptr<blockBytes> Pad2 = emptyBB(); // 2 bytes
	};

	struct FolderOrMessage
	{
		std::shared_ptr<blockT<WORD>> Type = emptyT<WORD>();
		FolderObject FolderObject;
		MessageObject MessageObject;
	};

	struct MessageDatabaseObject
	{
		std::shared_ptr<blockT<BYTE>> Version = emptyT<BYTE>();
		std::shared_ptr<blockT<BYTE>> Flag = emptyT<BYTE>();
		std::shared_ptr<blockStringA> DLLFileName = emptySA();
		bool bIsExchange{};
		std::shared_ptr<blockT<ULONG>> WrappedFlags = emptyT<ULONG>();
		std::shared_ptr<blockT<GUID>> WrappedProviderUID = emptyT<GUID>();
		std::shared_ptr<blockT<ULONG>> WrappedType = emptyT<ULONG>();
		std::shared_ptr<blockStringA> ServerShortname = emptySA();
		std::shared_ptr<blockStringA> MailboxDN = emptySA();
		std::shared_ptr<blockT<ULONG>> MagicVersion = emptyT<ULONG>();
		MDB_STORE_EID_V2 v2;
		MDB_STORE_EID_V3 v3;
		std::shared_ptr<blockStringA> v2DN = emptySA();
		std::shared_ptr<blockStringW> v2FQDN = emptySW();
		std::shared_ptr<blockStringW> v3SmtpAddress = emptySW();
		std::shared_ptr<blockBytes> v2Reserved = emptyBB(); // 2 bytes
	};

	struct EphemeralObject
	{
		std::shared_ptr<blockT<ULONG>> Version = emptyT<ULONG>();
		std::shared_ptr<blockT<ULONG>> Type = emptyT<ULONG>();
	};

	struct OneOffRecipientObject
	{
		std::shared_ptr<blockT<DWORD>> Bitmask = emptyT<DWORD>();
		struct Unicode
		{
			std::shared_ptr<blockStringW> DisplayName = emptySW();
			std::shared_ptr<blockStringW> AddressType = emptySW();
			std::shared_ptr<blockStringW> EmailAddress = emptySW();
		} Unicode;
		struct ANSI
		{
			std::shared_ptr<blockStringA> DisplayName = emptySA();
			std::shared_ptr<blockStringA> AddressType = emptySA();
			std::shared_ptr<blockStringA> EmailAddress = emptySA();
		} ANSI;
	};

	struct AddressBookObject
	{
		std::shared_ptr<blockT<DWORD>> Version = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> Type = emptyT<DWORD>();
		std::shared_ptr<blockStringA> X500DN = emptySA();
	};

	class EntryIdStruct;

	struct ContactAddressBookObject
	{
		std::shared_ptr<blockT<DWORD>> Version = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> Type = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> Index = emptyT<DWORD>(); // CONTAB_USER, CONTAB_DISTLIST only
		std::shared_ptr<blockT<DWORD>> EntryIDCount = emptyT<DWORD>(); // CONTAB_USER, CONTAB_DISTLIST only
		std::shared_ptr<blockT<GUID>> muidID = emptyT<GUID>(); // CONTAB_CONTAINER only
		std::shared_ptr<EntryIdStruct> lpEntryID;
	};

	struct WAB
	{
		std::shared_ptr<blockT<BYTE>> Type = emptyT<BYTE>();
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

		std::shared_ptr<blockT<byte>> m_abFlags0 = emptyT<byte>();
		std::shared_ptr<blockT<byte>> m_abFlags1 = emptyT<byte>();
		std::shared_ptr<blockBytes> m_abFlags23 = emptyBB(); // 2 bytes
		std::shared_ptr<blockT<GUID>> m_ProviderUID = emptyT<GUID>();
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