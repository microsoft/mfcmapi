// MAPIStorefunctions.cpp : Collection of useful MAPI Store functions

#include "stdafx.h"
#include "MAPIStoreFunctions.h"
#include "MAPIFunctions.h"
#include "String.h"
#include "Editor.h"
#include "Guids.h"
#include "InterpretProp2.h"
#include "ExtraPropTags.h"
#include "MAPIABFunctions.h"

_Check_return_ HRESULT CallOpenMsgStore(
	_In_ LPMAPISESSION lpSession,
	_In_ ULONG_PTR ulUIParam,
	_In_ LPSBinary lpEID,
	ULONG ulFlags,
	_Deref_out_ LPMDB* lpMDB)
{
	if (!lpSession || !lpMDB || !lpEID) return MAPI_E_INVALID_PARAMETER;

	auto hRes = S_OK;

	if (RegKeys[regkeyMDB_ONLINE].ulCurDWORD)
	{
		ulFlags |= MDB_ONLINE;
	}
	DebugPrint(DBGOpenItemProp, L"CallOpenMsgStore ulFlags = 0x%X\n", ulFlags);

	WC_MAPI(lpSession->OpenMsgStore(
		ulUIParam,
		lpEID->cb,
		reinterpret_cast<LPENTRYID>(lpEID->lpb),
		NULL,
		ulFlags,
		static_cast<LPMDB*>(lpMDB)));
	if (MAPI_E_UNKNOWN_FLAGS == hRes && ulFlags & MDB_ONLINE)
	{
		hRes = S_OK;
		// perhaps this store doesn't know the MDB_ONLINE flag - remove and retry
		ulFlags = ulFlags & ~MDB_ONLINE;
		DebugPrint(DBGOpenItemProp, L"CallOpenMsgStore 2nd attempt ulFlags = 0x%X\n", ulFlags);

		WC_MAPI(lpSession->OpenMsgStore(
			ulUIParam,
			lpEID->cb,
			reinterpret_cast<LPENTRYID>(lpEID->lpb),
			NULL,
			ulFlags,
			static_cast<LPMDB*>(lpMDB)));
	}
	return hRes;
}

// Build a server DN.
wstring BuildServerDN(
	wstring szServerName,
	wstring szPost)
{
	static wstring szPre = L"/cn=Configuration/cn=Servers/cn="; // STRING_OK
	return szPre + szServerName + szPost;
}

_Check_return_ HRESULT GetMailboxTable1(
	_In_ LPMDB lpMDB,
	wstring szServerDN,
	ULONG ulFlags,
	_Deref_out_opt_ LPMAPITABLE* lpMailboxTable)
{
	if (!lpMDB || !lpMailboxTable || szServerDN.empty()) return MAPI_E_INVALID_PARAMETER;
	*lpMailboxTable = nullptr;

	auto hRes = S_OK;
	LPEXCHANGEMANAGESTORE lpManageStore1 = nullptr;

	WC_MAPI(lpMDB->QueryInterface(
		IID_IExchangeManageStore,
		reinterpret_cast<LPVOID*>(&lpManageStore1)));

	if (lpManageStore1)
	{
		WC_MAPI(lpManageStore1->GetMailboxTable(
			LPSTR(static_cast<LPCTSTR>(wstringToCString(szServerDN))),
			lpMailboxTable,
			ulFlags));

		lpManageStore1->Release();
	}
	return hRes;
}

_Check_return_ HRESULT GetMailboxTable3(
	_In_ LPMDB lpMDB,
	wstring szServerDN,
	ULONG ulOffset,
	ULONG ulFlags,
	_Deref_out_opt_ LPMAPITABLE* lpMailboxTable)
{
	if (!lpMDB || !lpMailboxTable || szServerDN.empty()) return MAPI_E_INVALID_PARAMETER;
	*lpMailboxTable = nullptr;

	auto hRes = S_OK;
	LPEXCHANGEMANAGESTORE3 lpManageStore3 = nullptr;

	WC_MAPI(lpMDB->QueryInterface(
		IID_IExchangeManageStore3,
		reinterpret_cast<LPVOID*>(&lpManageStore3)));

	if (lpManageStore3)
	{
		WC_MAPI(lpManageStore3->GetMailboxTableOffset(
			LPSTR(static_cast<LPCTSTR>(wstringToCString(szServerDN))),
			lpMailboxTable,
			ulFlags,
			ulOffset));

		lpManageStore3->Release();
	}
	return hRes;
}

_Check_return_ HRESULT GetMailboxTable5(
	_In_ LPMDB lpMDB,
	wstring szServerDN,
	ULONG ulOffset,
	ULONG ulFlags,
	_In_opt_ LPGUID lpGuidMDB,
	_Deref_out_opt_ LPMAPITABLE* lpMailboxTable)
{
	if (!lpMDB || !lpMailboxTable || szServerDN.empty()) return MAPI_E_INVALID_PARAMETER;
	*lpMailboxTable = nullptr;

	auto hRes = S_OK;
	LPEXCHANGEMANAGESTORE5 lpManageStore5 = nullptr;

	EC_MAPI(lpMDB->QueryInterface(
		IID_IExchangeManageStore5,
		reinterpret_cast<LPVOID*>(&lpManageStore5)));

	if (lpManageStore5)
	{
		EC_MAPI(lpManageStore5->GetMailboxTableEx(
			LPSTR(static_cast<LPCTSTR>(wstringToCString(szServerDN))),
			lpGuidMDB,
			lpMailboxTable,
			ulFlags,
			ulOffset));

		lpManageStore5->Release();
	}
	return hRes;
}

// lpMDB needs to be an Exchange MDB - OpenMessageStoreGUID(pbExchangeProviderPrimaryUserGuid) can get one if there's one to be had
// Use GetServerName to get the default server
// Will try IID_IExchangeManageStore3 first and fail back to IID_IExchangeManageStore
_Check_return_ HRESULT GetMailboxTable(
	_In_ LPMDB lpMDB,
	wstring szServerName,
	ULONG ulOffset,
	_Deref_out_opt_ LPMAPITABLE* lpMailboxTable)
{
	if (!lpMDB || !StoreSupportsManageStore(lpMDB) || !lpMailboxTable) return MAPI_E_INVALID_PARAMETER;
	*lpMailboxTable = nullptr;

	auto hRes = S_OK;
	LPMAPITABLE lpLocalTable = nullptr;

	auto szServerDN = BuildServerDN(
		szServerName,
		L"");
	if (!szServerDN.empty())
	{
		WC_H(GetMailboxTable3(
			lpMDB,
			szServerDN,
			ulOffset,
			fMapiUnicode,
			&lpLocalTable));

		if (!lpLocalTable && 0 == ulOffset)
		{
			WC_H(GetMailboxTable1(
				lpMDB,
				szServerDN,
				fMapiUnicode,
				&lpLocalTable));
		}
	}

	*lpMailboxTable = lpLocalTable;
	return hRes;
}

_Check_return_ HRESULT GetPublicFolderTable1(
	_In_ LPMDB lpMDB,
	wstring szServerDN,
	ULONG ulFlags,
	_Deref_out_opt_ LPMAPITABLE* lpPFTable)
{
	if (!lpMDB || !lpPFTable || szServerDN.empty()) return MAPI_E_INVALID_PARAMETER;
	*lpPFTable = nullptr;

	auto hRes = S_OK;
	LPEXCHANGEMANAGESTORE lpManageStore1 = nullptr;

	EC_MAPI(lpMDB->QueryInterface(
		IID_IExchangeManageStore,
		reinterpret_cast<LPVOID*>(&lpManageStore1)));

	if (lpManageStore1)
	{
		EC_MAPI(lpManageStore1->GetPublicFolderTable(
			LPSTR(static_cast<LPCTSTR>(wstringToCString(szServerDN))),
			lpPFTable,
			ulFlags));

		lpManageStore1->Release();
	}

	return hRes;
}

_Check_return_ HRESULT GetPublicFolderTable4(
	_In_ LPMDB lpMDB,
	wstring szServerDN,
	ULONG ulOffset,
	ULONG ulFlags,
	_Deref_out_opt_ LPMAPITABLE* lpPFTable)
{
	if (!lpMDB || !lpPFTable || szServerDN.empty()) return MAPI_E_INVALID_PARAMETER;
	*lpPFTable = nullptr;

	auto hRes = S_OK;
	LPEXCHANGEMANAGESTORE4 lpManageStore4 = nullptr;

	EC_MAPI(lpMDB->QueryInterface(
		IID_IExchangeManageStore4,
		reinterpret_cast<LPVOID*>(&lpManageStore4)));

	if (lpManageStore4)
	{
		EC_MAPI(lpManageStore4->GetPublicFolderTableOffset(
			LPSTR(static_cast<LPCTSTR>(wstringToCString(szServerDN))),
			lpPFTable,
			ulFlags,
			ulOffset));
		lpManageStore4->Release();
	}

	return hRes;
}

_Check_return_ HRESULT GetPublicFolderTable5(
	_In_ LPMDB lpMDB,
	wstring szServerDN,
	ULONG ulOffset,
	ULONG ulFlags,
	_In_opt_ LPGUID lpGuidMDB,
	_Deref_out_opt_ LPMAPITABLE* lpPFTable)
{
	if (!lpMDB || !lpPFTable || szServerDN.empty()) return MAPI_E_INVALID_PARAMETER;
	*lpPFTable = nullptr;

	auto hRes = S_OK;
	LPEXCHANGEMANAGESTORE5 lpManageStore5 = nullptr;

	EC_MAPI(lpMDB->QueryInterface(
		IID_IExchangeManageStore5,
		reinterpret_cast<LPVOID*>(&lpManageStore5)));

	if (lpManageStore5)
	{
		EC_MAPI(lpManageStore5->GetPublicFolderTableEx(
			LPSTR(static_cast<LPCTSTR>(wstringToCString(szServerDN))),
			lpGuidMDB,
			lpPFTable,
			ulFlags,
			ulOffset));

		lpManageStore5->Release();
	}

	return hRes;
}

// Get server name from the profile
wstring GetServerName(_In_ LPMAPISESSION lpSession)
{
	auto hRes = S_OK;
	LPSERVICEADMIN pSvcAdmin = nullptr;
	LPPROFSECT pGlobalProfSect = nullptr;
	LPSPropValue lpServerName = nullptr;
	wstring serverName;

	if (!lpSession) return emptystring;

	EC_MAPI(lpSession->AdminServices(
		0,
		&pSvcAdmin));

	EC_MAPI(pSvcAdmin->OpenProfileSection(
		(LPMAPIUID)pbGlobalProfileSectionGuid,
		NULL,
		0,
		&pGlobalProfSect));

	EC_MAPI(HrGetOneProp(pGlobalProfSect,
		PR_PROFILE_HOME_SERVER,
		&lpServerName));

	if (CheckStringProp(lpServerName, PT_STRING8)) // profiles are ASCII only
	{
		serverName = LPCSTRToWstring(lpServerName->Value.lpszA);
	}
#ifndef MRMAPI
	else
	{
		// prompt the user to enter a server name
		CEditor MyData(
			nullptr,
			IDS_SERVERNAME,
			IDS_SERVERNAMEMISSINGPROMPT,
			1,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, CreateSingleLinePane(IDS_SERVERNAME, false));

		WC_H(MyData.DisplayDialog());

		if (S_OK == hRes)
		{
			serverName = MyData.GetStringW(0);
		}
	}
#endif
	MAPIFreeBuffer(lpServerName);
	if (pGlobalProfSect) pGlobalProfSect->Release();
	if (pSvcAdmin) pSvcAdmin->Release();
	return serverName;
}

_Check_return_ HRESULT CreateStoreEntryID(
	_In_ LPMDB lpMDB, // open message store
	wstring lpszMsgStoreDN, // desired message store DN
	wstring lpszMailboxDN, // desired mailbox DN or NULL
	ULONG ulFlags, // desired flags for CreateStoreEntryID
	_Out_opt_ ULONG* lpcbEntryID,
	_Deref_out_opt_ LPENTRYID * lppEntryID)
{
	auto hRes = S_OK;
	LPEXCHANGEMANAGESTORE lpXManageStore = nullptr;

	if (!lpMDB || lpszMsgStoreDN.empty() || !StoreSupportsManageStore(lpMDB))
	{
		if (!lpMDB) DebugPrint(DBGGeneric, L"CreateStoreEntryID: MDB was NULL\n");
		if (lpszMsgStoreDN.empty()) DebugPrint(DBGGeneric, L"CreateStoreEntryID: lpszMsgStoreDN was missing\n");
		return MAPI_E_INVALID_PARAMETER;
	}

	EC_MAPI(lpMDB->QueryInterface(
		IID_IExchangeManageStore,
		reinterpret_cast<LPVOID*>(&lpXManageStore)));

	if (lpXManageStore)
	{
		DebugPrint(DBGGeneric, L"CreateStoreEntryID: Creating EntryID. StoreDN = \"%hs\", MailboxDN = \"%hs\"\n", lpszMsgStoreDN, lpszMailboxDN);

		EC_MAPI(lpXManageStore->CreateStoreEntryID(
			const_cast<LPSTR>(static_cast<LPCSTR>(wstringToCStringA(lpszMsgStoreDN))),
			const_cast<LPSTR>(static_cast<LPCSTR>(wstringToCStringA(lpszMailboxDN))),
			ulFlags,
			lpcbEntryID,
			lppEntryID));

		lpXManageStore->Release();
	}

	return hRes;
}

_Check_return_ HRESULT CreateStoreEntryID2(
	_In_ LPMDB lpMDB, // open message store
	wstring lpszMsgStoreDN, // desired message store DN
	wstring lpszMailboxDN, // desired mailbox DN or NULL
	ULONG ulFlags, // desired flags for CreateStoreEntryID
	_Out_opt_ ULONG* lpcbEntryID,
	_Deref_out_opt_ LPENTRYID * lppEntryID)
{
	auto hRes = S_OK;
	LPEXCHANGEMANAGESTOREEX lpXManageStoreEx = nullptr;

	if (!lpMDB || lpszMsgStoreDN.empty() || !StoreSupportsManageStoreEx(lpMDB))
	{
		if (!lpMDB) DebugPrint(DBGGeneric, L"CreateStoreEntryID2: MDB was NULL\n");
		if (lpszMsgStoreDN.empty()) DebugPrint(DBGGeneric, L"CreateStoreEntryID2: lpszMsgStoreDN was missing\n");
		return MAPI_E_INVALID_PARAMETER;
	}

	EC_MAPI(lpMDB->QueryInterface(
		IID_IExchangeManageStoreEx,
		reinterpret_cast<LPVOID*>(&lpXManageStoreEx)));

	if (lpXManageStoreEx)
	{
		DebugPrint(DBGGeneric, L"CreateStoreEntryID2: Creating EntryID. StoreDN = \"%hs\", MailboxDN = \"%hs\"\n", lpszMsgStoreDN, lpszMailboxDN);
		SPropValue sProps[3] = { 0 };
		sProps[0].ulPropTag = PR_PROFILE_MAILBOX;
		sProps[0].Value.lpszA = const_cast<LPSTR>(static_cast<LPCSTR>(wstringToCStringA(lpszMailboxDN)));

		sProps[1].ulPropTag = PR_PROFILE_MDB_DN;
		sProps[1].Value.lpszA = const_cast<LPSTR>(static_cast<LPCSTR>(wstringToCStringA(lpszMsgStoreDN)));

		sProps[2].ulPropTag = PR_FORCE_USE_ENTRYID_SERVER;
		sProps[2].Value.b = true;

		EC_MAPI(lpXManageStoreEx->CreateStoreEntryID2(
			_countof(sProps),
			reinterpret_cast<LPSPropValue>(&sProps),
			ulFlags,
			lpcbEntryID,
			lppEntryID));

		lpXManageStoreEx->Release();
	}

	return hRes;
}

_Check_return_ HRESULT CreateStoreEntryID(
	_In_ LPMDB lpMDB, // open message store
	wstring lpszMsgStoreDN, // desired message store DN
	wstring lpszMailboxDN, // desired mailbox DN or NULL
	ULONG ulFlags, // desired flags for CreateStoreEntryID
	bool bForceServer, // Use CreateStoreEntryID2
	_Out_opt_ ULONG* lpcbEntryID,
	_Deref_out_opt_ LPENTRYID * lppEntryID)
{
	auto hRes = S_OK;

	// Use an empty MailboxDN to open the public store
	if (lpszMailboxDN.empty())
	{
		ulFlags |= OPENSTORE_PUBLIC;
	}

	if (!bForceServer)
	{
		EC_MAPI(CreateStoreEntryID(
			lpMDB,
			lpszMsgStoreDN,
			lpszMailboxDN,
			ulFlags,
			lpcbEntryID,
			lppEntryID));
	}
	else
	{
		EC_MAPI(CreateStoreEntryID2(
			lpMDB,
			lpszMsgStoreDN,
			lpszMailboxDN,
			ulFlags,
			lpcbEntryID,
			lppEntryID));
	}

	return hRes;
}

// Stolen from MBLogon.c in the EDK to avoid compiling and linking in the entire beast
// Cleaned up to fit in with other functions
// $--HrMailboxLogon------------------------------------------------------
// Logon to a mailbox. Before calling this function do the following:
// 1) Create a profile that has Exchange administrator privileges.
// 2) Logon to Exchange using this profile.
// 3) Open the mailbox using the Message Store DN and Mailbox DN.
//
// This version of the function needs the server and mailbox names to be
// in the form of distinguished names. They would look something like this:
// /CN=Configuration/CN=Servers/CN=%s/CN=Microsoft Private MDB
// /CN=Configuration/CN=Servers/CN=%s/CN=Microsoft Public MDB
// /O=<Organization>/OU=<Site>/CN=<Container>/CN=<MailboxName>
// where items in <brackets> would need to be set to appropriate values
//
// Note1: The message store DN is nearly identical to the server DN, except
// for the addition of a trailing '/CN=' part. This part is required although
// its actual value is ignored.
//
// Note2: A NULL lpszMailboxDN indicates the public store should be opened.
// -----------------------------------------------------------------------------

_Check_return_ HRESULT HrMailboxLogon(
	_In_ LPMAPISESSION lpMAPISession, // MAPI session handle
	_In_ LPMDB lpMDB, // open message store
	wstring lpszMsgStoreDN, // desired message store DN
	wstring lpszMailboxDN, // desired mailbox DN or NULL
	ULONG ulFlags, // desired flags for CreateStoreEntryID
	bool bForceServer, // Use CreateStoreEntryID2
	_Deref_out_opt_ LPMDB* lppMailboxMDB) // ptr to mailbox message store ptr
{
	auto hRes = S_OK;
	SBinary sbEID = { 0 };

	*lppMailboxMDB = nullptr;

	if (!lpMAPISession)
	{
		DebugPrint(DBGGeneric, L"HrMailboxLogon: Session was NULL\n");
		return MAPI_E_INVALID_PARAMETER;
	}

	WC_H(CreateStoreEntryID(lpMDB, lpszMsgStoreDN, lpszMailboxDN, ulFlags, bForceServer, &sbEID.cb, reinterpret_cast<LPENTRYID*>(&sbEID.lpb)));

	if (SUCCEEDED(hRes))
	{
		WC_H(CallOpenMsgStore(
			lpMAPISession,
			NULL,
			&sbEID,
			MDB_NO_DIALOG |
			MDB_NO_MAIL | // spooler not notified of our presence
			MDB_TEMPORARY | // message store not added to MAPI profile
			MAPI_BEST_ACCESS, // normally WRITE, but allow access to RO store
			lppMailboxMDB));
	}

	MAPIFreeBuffer(sbEID.lpb);
	return hRes;
}

_Check_return_ HRESULT OpenDefaultMessageStore(
	_In_ LPMAPISESSION lpMAPISession,
	_Deref_out_ LPMDB* lppDefaultMDB)
{
	auto hRes = S_OK;
	LPMAPITABLE pStoresTbl = nullptr;
	LPSRowSet pRow = nullptr;
	static SRestriction sres;
	SPropValue spv;

	enum
	{
		EID,
		NUM_COLS
	};
	static const SizedSPropTagArray(NUM_COLS, sptEIDCol) =
	{
	NUM_COLS,
	PR_ENTRYID,
	};
	if (!lpMAPISession) return MAPI_E_INVALID_PARAMETER;

	EC_MAPI(lpMAPISession->GetMsgStoresTable(0, &pStoresTbl));

	// set up restriction for the default store
	sres.rt = RES_PROPERTY; // gonna compare a property
	sres.res.resProperty.relop = RELOP_EQ; // gonna test equality
	sres.res.resProperty.ulPropTag = PR_DEFAULT_STORE; // tag to compare
	sres.res.resProperty.lpProp = &spv; // prop tag to compare against

	spv.ulPropTag = PR_DEFAULT_STORE; // tag type
	spv.Value.b = true; // tag value

	EC_MAPI(HrQueryAllRows(
		pStoresTbl, // table to query
		LPSPropTagArray(&sptEIDCol), // columns to get
		&sres, // restriction to use
		NULL, // sort order
		0, // max number of rows - 0 means ALL
		&pRow));

	if (pRow && pRow->cRows)
	{
		WC_H(CallOpenMsgStore(
			lpMAPISession,
			NULL,
			&pRow->aRow[0].lpProps[EID].Value.bin,
			MDB_WRITE,
			lppDefaultMDB));
	}
	else hRes = MAPI_E_NOT_FOUND;

	if (pRow) FreeProws(pRow);
	if (pStoresTbl) pStoresTbl->Release();
	return hRes;
}

// Build DNs for call to HrMailboxLogon
_Check_return_ HRESULT OpenOtherUsersMailbox(
	_In_ LPMAPISESSION lpMAPISession,
	_In_ LPMDB lpMDB,
	wstring szServerName,
	wstring szMailboxDN,
	ULONG ulFlags, // desired flags for CreateStoreEntryID
	bool bForceServer, // Use CreateStoreEntryID2
	_Deref_out_opt_ LPMDB* lppOtherUserMDB)
{
	auto hRes = S_OK;

	*lppOtherUserMDB = nullptr;

	DebugPrint(DBGGeneric, L"OpenOtherUsersMailbox called with lpMAPISession = %p, lpMDB = %p, Server = \"%ws\", Mailbox = \"%ws\"\n", lpMAPISession, lpMDB, szServerName.c_str(), szMailboxDN.c_str());
	if (!lpMAPISession || !lpMDB || szMailboxDN.empty() || !StoreSupportsManageStore(lpMDB)) return MAPI_E_INVALID_PARAMETER;

	if (szServerName.empty())
	{
		// If we weren't given a server name, get one from the profile
		szServerName = GetServerName(lpMAPISession);
	}

	if (!szServerName.empty())
	{
		auto szServerDN = BuildServerDN(
			szServerName,
			L"/cn=Microsoft Private MDB"); // STRING_OK

		if (!szServerDN.empty())
		{
			DebugPrint(DBGGeneric, L"Calling HrMailboxLogon with Server DN = \"%ws\"\n", szServerDN.c_str());
			WC_H(HrMailboxLogon(
				lpMAPISession,
				lpMDB,
				szServerDN,
				szMailboxDN,
				ulFlags,
				bForceServer,
				lppOtherUserMDB));
		}
	}

	return hRes;
}

#ifndef MRMAPI
_Check_return_ HRESULT OpenMailboxWithPrompt(
	_In_ LPMAPISESSION lpMAPISession,
	_In_ LPMDB lpMDB,
	wstring szServerName,
	wstring szMailboxDN,
	ULONG ulFlags, // desired flags for CreateStoreEntryID
	_Deref_out_opt_ LPMDB* lppOtherUserMDB)
{
	auto hRes = S_OK;
	*lppOtherUserMDB = nullptr;

	if (!lpMAPISession) return MAPI_E_INVALID_PARAMETER;

	CEditor MyPrompt(
		nullptr,
		IDS_OPENOTHERUSER,
		IDS_OPENWITHFLAGSPROMPT,
		4,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyPrompt.SetPromptPostFix(AllFlagsToString(PROP_ID(PR_PROFILE_OPEN_FLAGS), true));
	MyPrompt.InitPane(0, CreateSingleLinePane(IDS_SERVERNAME, szServerName, false));
	MyPrompt.InitPane(1, CreateSingleLinePane(IDS_USERDN, szMailboxDN, false));
	MyPrompt.InitPane(2, CreateSingleLinePane(IDS_CREATESTORENTRYIDFLAGS, false));
	MyPrompt.SetHex(2, ulFlags);
	MyPrompt.InitPane(3, CreateCheckPane(IDS_FORCESERVER, false, false));
	WC_H(MyPrompt.DisplayDialog());
	if (S_OK == hRes)
	{
		WC_H(OpenOtherUsersMailbox(
			lpMAPISession,
			lpMDB,
			MyPrompt.GetStringW(0),
			MyPrompt.GetStringW(1),
			MyPrompt.GetHex(2),
			MyPrompt.GetCheck(3),
			lppOtherUserMDB));
	}

	return hRes;
}

// Display a UI to select a mailbox, then call OpenOtherUsersMailbox with the mailboxDN
// May return MAPI_E_CANCEL
_Check_return_ HRESULT OpenOtherUsersMailboxFromGal(
	_In_ LPMAPISESSION lpMAPISession,
	_In_ LPADRBOOK lpAddrBook,
	_Deref_out_opt_ LPMDB* lppOtherUserMDB)
{
	auto hRes = S_OK;

	LPSPropValue lpEmailAddress = nullptr;
	LPMAILUSER lpMailUser = nullptr;
	LPMDB lpPrivateMDB = nullptr;

	*lppOtherUserMDB = nullptr;

	if (!lpMAPISession || !lpAddrBook) return MAPI_E_INVALID_PARAMETER;

	auto szServerName = GetServerName(lpMAPISession);

	WC_H(OpenMessageStoreGUID(lpMAPISession, pbExchangeProviderPrimaryUserGuid, &lpPrivateMDB));
	if (lpPrivateMDB && StoreSupportsManageStore(lpPrivateMDB))
	{
		EC_H(SelectUser(lpAddrBook, ::GetDesktopWindow(), NULL, &lpMailUser));
		if (lpMailUser)
		{
			EC_MAPI(HrGetOneProp(
				lpMailUser,
				PR_EMAIL_ADDRESS,
				&lpEmailAddress));

			if (CheckStringProp(lpEmailAddress, PT_TSTRING))
			{
				WC_H(OpenMailboxWithPrompt(
					lpMAPISession,
					lpPrivateMDB,
					szServerName,
					LPCTSTRToWstring(lpEmailAddress->Value.LPSZ),
					OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_TAKE_OWNERSHIP,
					lppOtherUserMDB));
			}
		}
	}

	MAPIFreeBuffer(lpEmailAddress);
	if (lpMailUser) lpMailUser->Release();
	if (lpPrivateMDB) lpPrivateMDB->Release();
	return hRes;
}
#endif

// Use these guids:
// pbExchangeProviderPrimaryUserGuid
// pbExchangeProviderDelegateGuid
// pbExchangeProviderPublicGuid
// pbExchangeProviderXportGuid
_Check_return_ HRESULT OpenMessageStoreGUID(
	_In_ LPMAPISESSION lpMAPISession,
	_In_z_ LPCSTR lpGUID, // Do not migrate this to wstring/string
	_Deref_out_opt_ LPMDB* lppMDB)
{
	LPMAPITABLE pStoresTbl = nullptr;
	LPSRowSet pRow = nullptr;
	ULONG ulRowNum;
	auto hRes = S_OK;

	enum
	{
		EID,
		STORETYPE,
		NUM_COLS
	};
	static const SizedSPropTagArray(NUM_COLS, sptCols) =
	{
	NUM_COLS,
	PR_ENTRYID,
	PR_MDB_PROVIDER
	};

	*lppMDB = nullptr;
	if (!lpMAPISession) return MAPI_E_INVALID_PARAMETER;

	EC_MAPI(lpMAPISession->GetMsgStoresTable(0, &pStoresTbl));

	if (pStoresTbl)
	{
		EC_MAPI(HrQueryAllRows(
			pStoresTbl, // table to query
			LPSPropTagArray(&sptCols), // columns to get
			NULL, // restriction to use
			NULL, // sort order
			0, // max number of rows
			&pRow));
		if (pRow)
		{
			if (!FAILED(hRes)) for (ulRowNum = 0; ulRowNum < pRow->cRows; ulRowNum++)
			{
				hRes = S_OK;
				// check to see if we have a folder with a matching GUID
				if (pRow->aRow[ulRowNum].lpProps[STORETYPE].ulPropTag == PR_MDB_PROVIDER &&
					pRow->aRow[ulRowNum].lpProps[EID].ulPropTag == PR_ENTRYID &&
					IsEqualMAPIUID(pRow->aRow[ulRowNum].lpProps[STORETYPE].Value.bin.lpb, lpGUID))
				{
					WC_H(CallOpenMsgStore(
						lpMAPISession,
						NULL,
						&pRow->aRow[ulRowNum].lpProps[EID].Value.bin,
						MDB_WRITE,
						lppMDB));
					break;
				}
			}
		}
	}
	if (!*lppMDB) hRes = MAPI_E_NOT_FOUND;

	if (pRow) FreeProws(pRow);
	if (pStoresTbl) pStoresTbl->Release();
	return hRes;
}

_Check_return_ HRESULT OpenPublicMessageStore(
	_In_ LPMAPISESSION lpMAPISession,
	ULONG ulFlags, // Flags for CreateStoreEntryID
	_Deref_out_opt_ LPMDB* lppPublicMDB)
{
	auto hRes = S_OK;

	LPMDB lpPublicMDBNonAdmin = nullptr;
	LPSPropValue lpServerName = nullptr;

	if (!lpMAPISession || !lppPublicMDB) return MAPI_E_INVALID_PARAMETER;

	WC_H(OpenMessageStoreGUID(
		lpMAPISession,
		pbExchangeProviderPublicGuid,
		&lpPublicMDBNonAdmin));

	// If we don't have flags we're done
	if (!ulFlags)
	{
		*lppPublicMDB = lpPublicMDBNonAdmin;
		return hRes;
	}

	if (lpPublicMDBNonAdmin && StoreSupportsManageStore(lpPublicMDBNonAdmin))
	{
		EC_MAPI(HrGetOneProp(
			lpPublicMDBNonAdmin,
			PR_HIERARCHY_SERVER,
			&lpServerName));

		if (CheckStringProp(lpServerName, PT_TSTRING))
		{
			wstring szServerDN;

			szServerDN = BuildServerDN(
				LPCTSTRToWstring(lpServerName->Value.LPSZ),
				L"/cn=Microsoft Public MDB"); // STRING_OK

			if (!szServerDN.empty())
			{
				WC_H(HrMailboxLogon(
					lpMAPISession,
					lpPublicMDBNonAdmin,
					szServerDN,
					emptystring,
					ulFlags,
					false,
					lppPublicMDB));
			}
		}
	}

	MAPIFreeBuffer(lpServerName);
	if (lpPublicMDBNonAdmin) lpPublicMDBNonAdmin->Release();
	return hRes;
}

_Check_return_ HRESULT OpenStoreFromMAPIProp(_In_ LPMAPISESSION lpMAPISession, _In_ LPMAPIPROP lpMAPIProp, _Deref_out_ LPMDB* lpMDB)
{
	auto hRes = S_OK;
	LPSPropValue lpProp = nullptr;

	if (!lpMAPISession || !lpMAPIProp) return MAPI_E_INVALID_PARAMETER;

	EC_MAPI(HrGetOneProp(
		lpMAPIProp,
		PR_STORE_ENTRYID,
		&lpProp));

	if (lpProp && PT_BINARY == PROP_TYPE(lpProp->ulPropTag))
	{
		WC_H(CallOpenMsgStore(
			lpMAPISession,
			NULL,
			&lpProp->Value.bin,
			MAPI_BEST_ACCESS,
			lpMDB));
	}

	MAPIFreeBuffer(lpProp);
	return hRes;
}

_Check_return_ bool StoreSupportsManageStore(_In_ LPMDB lpMDB)
{
	auto hRes = S_OK;
	LPEXCHANGEMANAGESTORE lpIManageStore = nullptr;

	if (!lpMDB) return false;

	EC_H_MSG(lpMDB->QueryInterface(
		IID_IExchangeManageStore,
		reinterpret_cast<LPVOID*>(&lpIManageStore)),
		IDS_MANAGESTORENOTSUPPORTED);

	if (lpIManageStore)
	{
		lpIManageStore->Release();
		return true;
	}
	return false;
}

_Check_return_ bool StoreSupportsManageStoreEx(_In_ LPMDB lpMDB)
{
	auto hRes = S_OK;
	LPEXCHANGEMANAGESTORE lpIManageStore = nullptr;

	if (!lpMDB) return false;

	EC_H_MSG(lpMDB->QueryInterface(
		IID_IExchangeManageStoreEx,
		reinterpret_cast<LPVOID*>(&lpIManageStore)),
		IDS_MANAGESTOREEXNOTSUPPORTED);

	if (lpIManageStore)
	{
		lpIManageStore->Release();
		return true;
	}
	return false;
}

_Check_return_ HRESULT HrUnWrapMDB(_In_ LPMDB lpMDBIn, _Deref_out_ LPMDB* lppMDBOut)
{
	if (!lpMDBIn || !lppMDBOut) return MAPI_E_INVALID_PARAMETER;
	auto hRes = S_OK;
	IProxyStoreObject* lpProxyObj = nullptr;
	LPMDB lpUnwrappedMDB = nullptr;
	EC_MAPI(lpMDBIn->QueryInterface(IID_IProxyStoreObject, reinterpret_cast<LPVOID*>(&lpProxyObj)));
	if (SUCCEEDED(hRes) && lpProxyObj)
	{
		EC_MAPI(lpProxyObj->UnwrapNoRef(reinterpret_cast<LPVOID*>(&lpUnwrappedMDB)));
		if (SUCCEEDED(hRes) && lpUnwrappedMDB)
		{
			// UnwrapNoRef doesn't addref, so we do it here:
			lpUnwrappedMDB->AddRef();
			*lppMDBOut = lpUnwrappedMDB;
		}
	}
	if (lpProxyObj) lpProxyObj->Release();
	return hRes;
}