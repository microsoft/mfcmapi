#pragma once
#include "SmartViewParser.h"

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
	ULONG ulMagic; // MDB_STORE_EID_V2_MAGIC
	ULONG ulSize; // size of this struct plus the size of szServerDN and wszServerFQDN
	ULONG ulVersion; // MDB_STORE_EID_V2_VERSION
	ULONG ulOffsetDN; // offset past the beginning of the MDB_STORE_EID_V2 struct where szServerDN starts
	ULONG ulOffsetFQDN; // offset past the beginning of the MDB_STORE_EID_V2 struct where wszServerFQDN starts
};

struct MDB_STORE_EID_V3
{
	ULONG ulMagic; // MDB_STORE_EID_V3_MAGIC
	ULONG ulSize; // size of this struct plus the size of szServerSmtpAddress
	ULONG ulVersion; // MDB_STORE_EID_V3_VERSION
	ULONG ulOffsetSmtpAddress; // offset past the beginning of the MDB_STORE_EID_V3 struct where szSmtpAddress starts
};

struct FolderObject
{
	GUID DatabaseGUID;
	std::vector<BYTE> GlobalCounter; // 6 bytes
	std::vector<BYTE> Pad; // 2 bytes
};

struct MessageObject
{
	GUID FolderDatabaseGUID;
	std::vector<BYTE> FolderGlobalCounter; // 6 bytes
	std::vector<BYTE> Pad1; // 2 bytes
	GUID MessageDatabaseGUID;
	std::vector<BYTE> MessageGlobalCounter; // 6 bytes
	std::vector<BYTE> Pad2; // 2 bytes
};

struct FolderOrMessage
{
	WORD Type;
	FolderObject FolderObject;
	MessageObject MessageObject;
};

struct MessageDatabaseObject
{
	BYTE Version;
	BYTE Flag;
	std::string DLLFileName;
	bool bIsExchange;
	ULONG WrappedFlags;
	GUID WrappedProviderUID;
	ULONG WrappedType;
	std::string ServerShortname;
	std::string MailboxDN;
	ULONG MagicVersion;
	MDB_STORE_EID_V2 v2;
	MDB_STORE_EID_V3 v3;
	std::string v2DN;
	std::wstring v2FQDN;
	std::wstring v3SmtpAddress;
	std::vector<BYTE> v2Reserved; // 2 bytes
};

struct EphemeralObject
{
	ULONG Version;
	ULONG Type;
};

struct OneOffRecipientObject
{
	DWORD Bitmask;
	struct Unicode
	{
		std::wstring DisplayName;
		std::wstring AddressType;
		std::wstring EmailAddress;
	} Unicode;
	struct ANSI
	{
		std::string DisplayName;
		std::string AddressType;
		std::string EmailAddress;
	} ANSI;
};

struct AddressBookObject
{
	DWORD Version;
	DWORD Type;
	std::string X500DN;
};

class EntryIdStruct;

struct ContactAddressBookObject
{
	DWORD Version;
	DWORD Type;
	DWORD Index; // CONTAB_USER, CONTAB_DISTLIST only
	DWORD EntryIDCount; // CONTAB_USER, CONTAB_DISTLIST only
	GUID muidID; // CONTAB_CONTAINER only
	std::vector<EntryIdStruct> lpEntryID;
};

struct WAB
{
	BYTE Type;
	std::vector<EntryIdStruct> lpEntryID;
};

class EntryIdStruct : public SmartViewParser
{
public:
	EntryIdStruct();

private:
	void Parse() override;
	_Check_return_ std::wstring ToStringInternal() override;

	std::vector<BYTE> m_abFlags; // 4 bytes
	GUID m_ProviderUID;
	EIDStructType m_ObjectType; // My own addition to simplify parsing
	FolderOrMessage m_FolderOrMessage;
	MessageDatabaseObject m_MessageDatabaseObject;
	EphemeralObject m_EphemeralObject;
	OneOffRecipientObject m_OneOffRecipientObject;
	AddressBookObject m_AddressBookObject;
	ContactAddressBookObject m_ContactAddressBookObject;
	WAB m_WAB;
};