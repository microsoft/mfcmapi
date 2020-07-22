#pragma once
#include <core/smartview/block/block.h>
#include <core/smartview/block/blockStringA.h>
#include <core/smartview/block/blockStringW.h>
#include <core/smartview/block/blockBytes.h>
#include <core/smartview/block/blockT.h>

namespace smartview
{
	class MDB_STORE_EID_V2 : public block
	{
	public:
		static const size_t size = sizeof(ULONG) * 5;

	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<ULONG>> ulMagic = emptyT<ULONG>(); // MDB_STORE_EID_V2_MAGIC
		std::shared_ptr<blockT<ULONG>> ulSize =
			emptyT<ULONG>(); // size of this struct plus the size of szServerDN and wszServerFQDN
		std::shared_ptr<blockT<ULONG>> ulVersion = emptyT<ULONG>(); // MDB_STORE_EID_V2_VERSION
		std::shared_ptr<blockT<ULONG>> ulOffsetDN =
			emptyT<ULONG>(); // offset past the beginning of the MDB_STORE_EID_V2 struct where szServerDN starts
		std::shared_ptr<blockT<ULONG>> ulOffsetFQDN =
			emptyT<ULONG>(); // offset past the beginning of the MDB_STORE_EID_V2 struct where wszServerFQDN starts
		std::shared_ptr<blockStringA> v2DN = emptySA();
		std::shared_ptr<blockStringW> v2FQDN = emptySW();
		std::shared_ptr<blockBytes> v2Reserved = emptyBB(); // 2 bytes
	};

	class MDB_STORE_EID_V3 : public block
	{
	public:
		static const size_t size = sizeof(ULONG) * 4;

	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<ULONG>> ulMagic = emptyT<ULONG>(); // MDB_STORE_EID_V3_MAGIC
		std::shared_ptr<blockT<ULONG>> ulSize =
			emptyT<ULONG>(); // size of this struct plus the size of szServerSmtpAddress
		std::shared_ptr<blockT<ULONG>> ulVersion = emptyT<ULONG>(); // MDB_STORE_EID_V3_VERSION
		std::shared_ptr<blockT<ULONG>> ulOffsetSmtpAddress =
			emptyT<ULONG>(); // offset past the beginning of the MDB_STORE_EID_V3 struct where szSmtpAddress starts
		std::shared_ptr<blockStringW> v3SmtpAddress = emptySW();
		std::shared_ptr<blockBytes> v2Reserved = emptyBB(); // 2 bytes
	};

	class FolderObject : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<WORD>> Type = emptyT<WORD>();
		std::shared_ptr<blockT<GUID>> DatabaseGUID = emptyT<GUID>();
		std::shared_ptr<blockBytes> GlobalCounter = emptyBB(); // 6 bytes
		std::shared_ptr<blockBytes> Pad = emptyBB(); // 2 bytes
	};

	class MessageObject : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<WORD>> Type = emptyT<WORD>();
		std::shared_ptr<blockT<GUID>> FolderDatabaseGUID = emptyT<GUID>();
		std::shared_ptr<blockBytes> FolderGlobalCounter = emptyBB(); // 6 bytes
		std::shared_ptr<blockBytes> Pad1 = emptyBB(); // 2 bytes
		std::shared_ptr<blockT<GUID>> MessageDatabaseGUID = emptyT<GUID>();
		std::shared_ptr<blockBytes> MessageGlobalCounter = emptyBB(); // 6 bytes
		std::shared_ptr<blockBytes> Pad2 = emptyBB(); // 2 bytes
	};

	class MessageDatabaseObject : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		bool bIsExchange{};
		std::shared_ptr<blockT<BYTE>> Version = emptyT<BYTE>();
		std::shared_ptr<blockT<BYTE>> Flag = emptyT<BYTE>();
		std::shared_ptr<blockStringA> DLLFileName = emptySA();
		std::shared_ptr<blockT<ULONG>> WrappedFlags = emptyT<ULONG>();
		std::shared_ptr<blockT<GUID>> WrappedProviderUID = emptyT<GUID>();
		std::shared_ptr<blockT<ULONG>> WrappedType = emptyT<ULONG>();
		std::shared_ptr<blockStringA> ServerShortname = emptySA();
		std::shared_ptr<blockStringA> MailboxDN = emptySA();
		std::shared_ptr<blockT<ULONG>> MagicVersion = emptyT<ULONG>();
		std::shared_ptr<MDB_STORE_EID_V2> v2;
		std::shared_ptr<MDB_STORE_EID_V3> v3;
	};

	class EphemeralObject : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<ULONG>> Version = emptyT<ULONG>();
		std::shared_ptr<blockT<ULONG>> Type = emptyT<ULONG>();
	};

	class OneOffRecipientObject : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<DWORD>> Bitmask = emptyT<DWORD>();
		std::shared_ptr<blockStringW> DisplayNameW = emptySW();
		std::shared_ptr<blockStringW> AddressTypeW = emptySW();
		std::shared_ptr<blockStringW> EmailAddressW = emptySW();
		std::shared_ptr<blockStringA> DisplayNameA = emptySA();
		std::shared_ptr<blockStringA> AddressTypeA = emptySA();
		std::shared_ptr<blockStringA> EmailAddressA = emptySA();
	};

	class AddressBookObject : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<DWORD>> Version = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> Type = emptyT<DWORD>();
		std::shared_ptr<blockStringA> X500DN = emptySA();
	};

	class EntryIdStruct;

	class ContactAddressBookObject : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<DWORD>> Version = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> Type = emptyT<DWORD>();
		std::shared_ptr<blockT<DWORD>> Index = emptyT<DWORD>(); // CONTAB_USER, CONTAB_DISTLIST only
		std::shared_ptr<blockT<DWORD>> EntryIDCount = emptyT<DWORD>(); // CONTAB_USER, CONTAB_DISTLIST only
		std::shared_ptr<blockT<GUID>> muidID = emptyT<GUID>(); // CONTAB_CONTAINER only
		std::shared_ptr<EntryIdStruct> lpEntryID;
	};

	class WAB : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		std::shared_ptr<blockT<BYTE>> Type = emptyT<BYTE>();
		std::shared_ptr<EntryIdStruct> lpEntryID;
	};

	class EntryIdStruct : public block
	{
	private:
		void parse() override;
		void parseBlocks() override;

		enum class EIDStructType
		{
			unknown = 0,
			ephemeral,
			shortTerm,
			folder,
			message,
			messageDatabase,
			oneOff,
			addressBook,
			contact,
			newsGroupFolder,
			WAB,
		};

		EIDStructType m_ObjectType = EIDStructType::unknown; // My own addition to simplify parsing
		std::shared_ptr<blockT<byte>> m_abFlags0 = emptyT<byte>();
		std::shared_ptr<blockT<byte>> m_abFlags1 = emptyT<byte>();
		std::shared_ptr<blockBytes> m_abFlags23 = emptyBB(); // 2 bytes
		std::shared_ptr<blockT<GUID>> m_ProviderUID = emptyT<GUID>();
		std::shared_ptr<FolderObject> m_FolderObject;
		std::shared_ptr<MessageObject> m_MessageObject;
		std::shared_ptr<MessageDatabaseObject> m_MessageDatabaseObject;
		std::shared_ptr<EphemeralObject> m_EphemeralObject{};
		std::shared_ptr<OneOffRecipientObject> m_OneOffRecipientObject{};
		std::shared_ptr<AddressBookObject> m_AddressBookObject;
		std::shared_ptr<ContactAddressBookObject> m_ContactAddressBookObject;
		std::shared_ptr<WAB> m_WAB;
	};
} // namespace smartview