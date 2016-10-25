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
	BYTE DatabaseGUID[16];
	BYTE GlobalCounter[6];
	BYTE Pad[2];
};

struct MessageObject
{
	BYTE FolderDatabaseGUID[16];
	BYTE FolderGlobalCounter[6];
	BYTE Pad1[2];
	BYTE MessageDatabaseGUID[16];
	BYTE MessageGlobalCounter[6];
	BYTE Pad2[2];
};

struct FolderOrMessage
{
	WORD Type;
	union
	{
		FolderObject FolderObject;
		MessageObject MessageObject;
	} Data;
};

struct MessageDatabaseObject
{
	BYTE Version;
	BYTE Flag;
	LPSTR DLLFileName;
	bool bIsExchange;
	ULONG WrappedFlags;
	BYTE WrappedProviderUID[16];
	ULONG WrappedType;
	LPSTR ServerShortname;
	LPSTR MailboxDN;
	ULONG MagicVersion;
	MDB_STORE_EID_V2 v2;
	MDB_STORE_EID_V3 v3;
	LPSTR v2DN;
	LPWSTR v2FQDN;
	LPWSTR v3SmtpAddress;
	BYTE v2Reserved[2];
};

struct EphemeralObject
{
	ULONG Version;
	ULONG Type;
};

struct OneOffRecipientObject
{
	DWORD Bitmask;
	union
	{
		struct
		{
			LPWSTR DisplayName;
			LPWSTR AddressType;
			LPWSTR EmailAddress;
		} Unicode;
		struct
		{
			LPSTR DisplayName;
			LPSTR AddressType;
			LPSTR EmailAddress;
		} ANSI;
	} Strings;
} ;

struct AddressBookObject
{
	DWORD Version;
	DWORD Type;
	LPSTR X500DN;
} ;

class EntryIdStruct;

struct ContactAddressBookObject
{
	DWORD Version;
	DWORD Type;
	DWORD Index; // CONTAB_USER, CONTAB_DISTLIST only
	DWORD EntryIDCount; // CONTAB_USER, CONTAB_DISTLIST only
	BYTE muidID[16]; // CONTAB_CONTAINER only
	EntryIdStruct* lpEntryID;
} ;

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

	BYTE m_abFlags[4];
	BYTE m_ProviderUID[16];
	EIDStructType m_ObjectType; // My own addition to simplify union parsing
	union
	{
		FolderOrMessage FolderOrMessage;
		MessageDatabaseObject MessageDatabaseObject;
		EphemeralObject EphemeralObject;
		OneOffRecipientObject OneOffRecipientObject;
		AddressBookObject AddressBookObject;
		ContactAddressBookObject ContactAddressBookObject;
		WAB WAB;
	} m_ProviderData;
};