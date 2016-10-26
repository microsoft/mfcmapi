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
	vector<BYTE> DatabaseGUID; // 16 bytes
	vector<BYTE> GlobalCounter; // 6 bytes
	vector<BYTE> Pad; // 2 bytes
};

struct MessageObject
{
	vector<BYTE> FolderDatabaseGUID; // 16 bytes
	vector<BYTE> FolderGlobalCounter; // 6 bytes
	vector<BYTE> Pad1; // 2 bytes
	vector<BYTE> MessageDatabaseGUID; // 16 bytes
	vector<BYTE> MessageGlobalCounter; // 6 bytes
	vector<BYTE> Pad2; // 2 bytes
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
	string DLLFileName;
	bool bIsExchange;
	ULONG WrappedFlags;
	vector<BYTE> WrappedProviderUID; // 16 bytes
	ULONG WrappedType;
	string ServerShortname;
	string MailboxDN;
	ULONG MagicVersion;
	MDB_STORE_EID_V2 v2;
	MDB_STORE_EID_V3 v3;
	string v2DN;
	wstring v2FQDN;
	wstring v3SmtpAddress;
	vector<BYTE> v2Reserved; // 2 bytes
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
		wstring DisplayName;
		wstring AddressType;
		wstring EmailAddress;
	} Unicode;
	struct ANSI
	{
		string DisplayName;
		string AddressType;
		string EmailAddress;
	} ANSI;
};

struct AddressBookObject
{
	DWORD Version;
	DWORD Type;
	string X500DN;
};

class EntryIdStruct;

struct ContactAddressBookObject
{
	DWORD Version;
	DWORD Type;
	DWORD Index; // CONTAB_USER, CONTAB_DISTLIST only
	DWORD EntryIDCount; // CONTAB_USER, CONTAB_DISTLIST only
	vector<BYTE> muidID; // 16 bytes. CONTAB_CONTAINER only
	EntryIdStruct* lpEntryID;
};

struct WAB
{
	BYTE Type;
	EntryIdStruct* lpEntryID;
};

class EntryIdStruct : public SmartViewParser
{
public:
	EntryIdStruct();
	~EntryIdStruct();

private:
	void Parse() override;
	_Check_return_ wstring ToStringInternal() override;

	vector<BYTE> m_abFlags; // 4 bytes
	vector<BYTE> m_ProviderUID; // 16 bytes
	EIDStructType m_ObjectType; // My own addition to simplify parsing
	FolderOrMessage m_FolderOrMessage;
	MessageDatabaseObject m_MessageDatabaseObject;
	EphemeralObject m_EphemeralObject;
	OneOffRecipientObject m_OneOffRecipientObject;
	AddressBookObject m_AddressBookObject;
	ContactAddressBookObject m_ContactAddressBookObject;
	WAB m_WAB;
};