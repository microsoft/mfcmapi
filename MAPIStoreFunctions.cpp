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
	_In_ LPMAPISESSION	lpSession,
	_In_ ULONG_PTR		ulUIParam,
	_In_ LPSBinary		lpEID,
	ULONG			ulFlags,
	_Deref_out_ LPMDB*			lpMDB)
{
	if (!lpSession || !lpMDB || !lpEID) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;

	if (RegKeys[regkeyMDB_ONLINE].ulCurDWORD)
	{
		ulFlags |= MDB_ONLINE;
	}
	DebugPrint(DBGOpenItemProp, L"CallOpenMsgStore ulFlags = 0x%X\n", ulFlags);

	WC_MAPI(lpSession->OpenMsgStore(
		ulUIParam,
		lpEID->cb,
		(LPENTRYID)lpEID->lpb,
		NULL,
		ulFlags,
		(LPMDB*)lpMDB));
	if (MAPI_E_UNKNOWN_FLAGS == hRes && (ulFlags & MDB_ONLINE))
	{
		hRes = S_OK;
		// perhaps this store doesn't know the MDB_ONLINE flag - remove and retry
		ulFlags = ulFlags & ~MDB_ONLINE;
		DebugPrint(DBGOpenItemProp, L"CallOpenMsgStore 2nd attempt ulFlags = 0x%X\n", ulFlags);

		WC_MAPI(lpSession->OpenMsgStore(
			ulUIParam,
			lpEID->cb,
			(LPENTRYID)lpEID->lpb,
			NULL,
			ulFlags,
			(LPMDB*)lpMDB));
	}
	return hRes;
} // CallOpenMsgStore

// Build a server DN. Allocates memory. Free with MAPIFreeBuffer.
_Check_return_ HRESULT BuildServerDN(
	_In_z_ LPCTSTR szServerName,
	_In_z_ LPCTSTR szPost,
	_Deref_out_z_ LPTSTR* lpszServerDN)
{
	HRESULT hRes = S_OK;
	if (!lpszServerDN) return MAPI_E_INVALID_PARAMETER;

	static LPCTSTR szPre = _T("/cn=Configuration/cn=Servers/cn="); // STRING_OK
	size_t cbPreLen = 0;
	size_t cbServerLen = 0;
	size_t cbPostLen = 0;
	size_t cbServerDN = 0;

	EC_H(StringCbLength(szPre, STRSAFE_MAX_CCH * sizeof(TCHAR), &cbPreLen));
	EC_H(StringCbLength(szServerName, STRSAFE_MAX_CCH * sizeof(TCHAR), &cbServerLen));
	EC_H(StringCbLength(szPost, STRSAFE_MAX_CCH * sizeof(TCHAR), &cbPostLen));

	cbServerDN = cbPreLen + cbServerLen + cbPostLen + sizeof(TCHAR);

	EC_H(MAPIAllocateBuffer(
		(ULONG)cbServerDN,
		(LPVOID*)lpszServerDN));

	EC_H(StringCbPrintf(
		*lpszServerDN,
		cbServerDN,
		_T("%s%s%s"), // STRING_OK
		szPre,
		szServerName,
		szPost));
	return hRes;
} // BuildServerDN

_Check_return_ HRESULT GetMailboxTable1(
	_In_ LPMDB lpMDB,
	_In_z_ LPCTSTR szServerDN,
	ULONG ulFlags,
	_Deref_out_opt_ LPMAPITABLE* lpMailboxTable)
{
	if (!lpMDB || !lpMailboxTable || !szServerDN) return MAPI_E_INVALID_PARAMETER;
	*lpMailboxTable = NULL;

	HRESULT	hRes = S_OK;
	LPEXCHANGEMANAGESTORE lpManageStore1 = NULL;

	WC_MAPI(lpMDB->QueryInterface(
		IID_IExchangeManageStore,
		(void **)&lpManageStore1));

	if (lpManageStore1)
	{
		WC_MAPI(lpManageStore1->GetMailboxTable(
			(LPSTR)szServerDN,
			lpMailboxTable,
			ulFlags));

		lpManageStore1->Release();
	}
	return hRes;
} // GetMailboxTable1

_Check_return_ HRESULT GetMailboxTable3(
	_In_ LPMDB lpMDB,
	_In_z_ LPCTSTR szServerDN,
	ULONG ulOffset,
	ULONG ulFlags,
	_Deref_out_opt_ LPMAPITABLE* lpMailboxTable)
{
	if (!lpMDB || !lpMailboxTable || !szServerDN) return MAPI_E_INVALID_PARAMETER;
	*lpMailboxTable = NULL;

	HRESULT	hRes = S_OK;
	LPEXCHANGEMANAGESTORE3 lpManageStore3 = NULL;

	WC_MAPI(lpMDB->QueryInterface(
		IID_IExchangeManageStore3,
		(void **)&lpManageStore3));

	if (lpManageStore3)
	{
		WC_MAPI(lpManageStore3->GetMailboxTableOffset(
			(LPSTR)szServerDN,
			lpMailboxTable,
			ulFlags,
			ulOffset));

		lpManageStore3->Release();
	}
	return hRes;
} // GetMailboxTable3

_Check_return_ HRESULT GetMailboxTable5(
	_In_ LPMDB lpMDB,
	_In_z_ LPCTSTR szServerDN,
	ULONG ulOffset,
	ULONG ulFlags,
	_In_opt_ LPGUID lpGuidMDB,
	_Deref_out_opt_ LPMAPITABLE* lpMailboxTable)
{
	if (!lpMDB || !lpMailboxTable || !szServerDN) return MAPI_E_INVALID_PARAMETER;
	*lpMailboxTable = NULL;

	HRESULT	hRes = S_OK;
	LPEXCHANGEMANAGESTORE5 lpManageStore5 = NULL;

	EC_MAPI(lpMDB->QueryInterface(
		IID_IExchangeManageStore5,
		(void **)&lpManageStore5));

	if (lpManageStore5)
	{
		EC_MAPI(lpManageStore5->GetMailboxTableEx(
			(LPSTR)szServerDN,
			lpGuidMDB,
			lpMailboxTable,
			ulFlags,
			ulOffset));

		lpManageStore5->Release();
	}
	return hRes;
} // GetMailboxTable5

// lpMDB needs to be an Exchange MDB - OpenMessageStoreGUID(pbExchangeProviderPrimaryUserGuid) can get one if there's one to be had
// Use GetServerName to get the default server
// Will try IID_IExchangeManageStore3 first and fail back to IID_IExchangeManageStore
_Check_return_ HRESULT GetMailboxTable(
	_In_ LPMDB lpMDB,
	_In_z_ LPCTSTR szServerName,
	ULONG ulOffset,
	_Deref_out_opt_ LPMAPITABLE* lpMailboxTable)
{
	if (!lpMDB || !StoreSupportsManageStore(lpMDB) || !lpMailboxTable) return MAPI_E_INVALID_PARAMETER;
	*lpMailboxTable = NULL;

	HRESULT	hRes = S_OK;
	LPTSTR	szServerDN = NULL;
	LPMAPITABLE lpLocalTable = NULL;

	EC_H(BuildServerDN(
		szServerName,
		_T(""),
		&szServerDN));
	if (szServerDN)
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
	MAPIFreeBuffer(szServerDN);
	return hRes;
} // GetMailboxTable

_Check_return_ HRESULT GetPublicFolderTable1(
	_In_ LPMDB lpMDB,
	_In_z_ LPCTSTR szServerDN,
	ULONG ulFlags,
	_Deref_out_opt_ LPMAPITABLE* lpPFTable)
{
	if (!lpMDB || !lpPFTable || !szServerDN) return MAPI_E_INVALID_PARAMETER;
	*lpPFTable = NULL;

	HRESULT	hRes = S_OK;
	LPEXCHANGEMANAGESTORE lpManageStore1 = NULL;

	EC_MAPI(lpMDB->QueryInterface(
		IID_IExchangeManageStore,
		(void **)&lpManageStore1));

	if (lpManageStore1)
	{
		EC_MAPI(lpManageStore1->GetPublicFolderTable(
			(LPSTR)szServerDN,
			lpPFTable,
			ulFlags));

		lpManageStore1->Release();
	}

	return hRes;
} // GetPublicFolderTable1

_Check_return_ HRESULT GetPublicFolderTable4(
	_In_ LPMDB lpMDB,
	_In_z_ LPCTSTR szServerDN,
	ULONG ulOffset,
	ULONG ulFlags,
	_Deref_out_opt_ LPMAPITABLE* lpPFTable)
{
	if (!lpMDB || !lpPFTable || !szServerDN) return MAPI_E_INVALID_PARAMETER;
	*lpPFTable = NULL;

	HRESULT	hRes = S_OK;
	LPEXCHANGEMANAGESTORE4 lpManageStore4 = NULL;

	EC_MAPI(lpMDB->QueryInterface(
		IID_IExchangeManageStore4,
		(void **)&lpManageStore4));

	if (lpManageStore4)
	{
		EC_MAPI(lpManageStore4->GetPublicFolderTableOffset(
			(LPSTR)szServerDN,
			lpPFTable,
			ulFlags,
			ulOffset));
		lpManageStore4->Release();
	}

	return hRes;
} // GetPublicFolderTable4

_Check_return_ HRESULT GetPublicFolderTable5(
	_In_ LPMDB lpMDB,
	_In_z_ LPCTSTR szServerDN,
	ULONG ulOffset,
	ULONG ulFlags,
	_In_opt_ LPGUID lpGuidMDB,
	_Deref_out_opt_ LPMAPITABLE* lpPFTable)
{
	if (!lpMDB || !lpPFTable || !szServerDN) return MAPI_E_INVALID_PARAMETER;
	*lpPFTable = NULL;

	HRESULT	hRes = S_OK;
	LPEXCHANGEMANAGESTORE5 lpManageStore5 = NULL;

	EC_MAPI(lpMDB->QueryInterface(
		IID_IExchangeManageStore5,
		(void **)&lpManageStore5));

	if (lpManageStore5)
	{
		EC_MAPI(lpManageStore5->GetPublicFolderTableEx(
			(LPSTR)szServerDN,
			lpGuidMDB,
			lpPFTable,
			ulFlags,
			ulOffset));

		lpManageStore5->Release();
	}

	return hRes;
} // GetPublicFolderTable5

// Get server name from the profile
_Check_return_ HRESULT GetServerName(_In_ LPMAPISESSION lpSession, _Deref_out_opt_z_ LPTSTR* szServerName)
{
	HRESULT			hRes = S_OK;
	LPSERVICEADMIN	pSvcAdmin = NULL;
	LPPROFSECT		pGlobalProfSect = NULL;
	LPSPropValue	lpServerName = NULL;

	if (!lpSession) return MAPI_E_INVALID_PARAMETER;

	*szServerName = NULL;

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
#ifdef UNICODE
		LPWSTR	szWideServer = NULL;
		EC_H(AnsiToUnicode(
			lpServerName->Value.lpszA,
			&szWideServer));
		EC_H(CopyStringW(szServerName,szWideServer,NULL));
		delete[] szWideServer;
#else
		EC_H(CopyStringA(szServerName, lpServerName->Value.lpszA, NULL));
#endif
	}
#ifndef MRMAPI
	else
	{
		// prompt the user to enter a server name
		CEditor MyData(
			NULL,
			IDS_SERVERNAME,
			IDS_SERVERNAMEMISSINGPROMPT,
			1,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.InitPane(0, CreateSingleLinePane(IDS_SERVERNAME, NULL, false));

		WC_H(MyData.DisplayDialog());

		if (S_OK == hRes)
		{
			EC_H(CopyString(szServerName, MyData.GetString(0), NULL));
		}
	}
#endif
	MAPIFreeBuffer(lpServerName);
	if (pGlobalProfSect) pGlobalProfSect->Release();
	if (pSvcAdmin) pSvcAdmin->Release();
	return hRes;
} // GetServerName

_Check_return_ HRESULT CreateStoreEntryID(
	_In_ LPMDB lpMDB, // open message store
	_In_z_ LPCSTR lpszMsgStoreDN, // desired message store DN
	_In_opt_z_ LPCSTR lpszMailboxDN, // desired mailbox DN or NULL
	ULONG ulFlags, // desired flags for CreateStoreEntryID
	_Out_opt_ ULONG* lpcbEntryID,
	_Deref_out_opt_ LPENTRYID * lppEntryID)
{
	HRESULT hRes = S_OK;
	LPEXCHANGEMANAGESTORE lpXManageStore = NULL;

	if (!lpMDB || !lpszMsgStoreDN || !StoreSupportsManageStore(lpMDB))
	{
		if (!lpMDB) DebugPrint(DBGGeneric, L"CreateStoreEntryID: MDB was NULL\n");
		if (!lpszMsgStoreDN) DebugPrint(DBGGeneric, L"CreateStoreEntryID: lpszMsgStoreDN was NULL\n");
		return MAPI_E_INVALID_PARAMETER;
	}

	EC_MAPI(lpMDB->QueryInterface(
		IID_IExchangeManageStore,
		(LPVOID*)&lpXManageStore));

	if (lpXManageStore)
	{
		DebugPrint(DBGGeneric, L"CreateStoreEntryID: Creating EntryID. StoreDN = \"%hs\", MailboxDN = \"%hs\"\n", lpszMsgStoreDN, lpszMailboxDN);

		EC_MAPI(lpXManageStore->CreateStoreEntryID(
			(LPSTR)lpszMsgStoreDN,
			(LPSTR)lpszMailboxDN,
			ulFlags,
			lpcbEntryID,
			lppEntryID));

		lpXManageStore->Release();
	}

	return hRes;
}

_Check_return_ HRESULT CreateStoreEntryID2(
	_In_ LPMDB lpMDB, // open message store
	_In_z_ LPCSTR lpszMsgStoreDN, // desired message store DN
	_In_opt_z_ LPCSTR lpszMailboxDN, // desired mailbox DN or NULL
	ULONG ulFlags, // desired flags for CreateStoreEntryID
	_Out_opt_ ULONG* lpcbEntryID,
	_Deref_out_opt_ LPENTRYID * lppEntryID)
{
	HRESULT hRes = S_OK;
	LPEXCHANGEMANAGESTOREEX lpXManageStoreEx = NULL;

	if (!lpMDB || !lpszMsgStoreDN || !StoreSupportsManageStoreEx(lpMDB))
	{
		if (!lpMDB) DebugPrint(DBGGeneric, L"CreateStoreEntryID2: MDB was NULL\n");
		if (!lpszMsgStoreDN) DebugPrint(DBGGeneric, L"CreateStoreEntryID2: lpszMsgStoreDN was NULL\n");
		return MAPI_E_INVALID_PARAMETER;
	}

	EC_MAPI(lpMDB->QueryInterface(
		IID_IExchangeManageStoreEx,
		(LPVOID*)&lpXManageStoreEx));

	if (lpXManageStoreEx)
	{
		DebugPrint(DBGGeneric, L"CreateStoreEntryID2: Creating EntryID. StoreDN = \"%hs\", MailboxDN = \"%hs\"\n", lpszMsgStoreDN, lpszMailboxDN);
		SPropValue sProps[3] = { 0 };
		sProps[0].ulPropTag = PR_PROFILE_MAILBOX;
		sProps[0].Value.lpszA = (LPSTR)lpszMailboxDN;

		sProps[1].ulPropTag = PR_PROFILE_MDB_DN;
		sProps[1].Value.lpszA = (LPSTR)lpszMsgStoreDN;

		sProps[2].ulPropTag = PR_FORCE_USE_ENTRYID_SERVER;
		sProps[2].Value.b = true;

		EC_MAPI(lpXManageStoreEx->CreateStoreEntryID2(
			_countof(sProps),
			(LPSPropValue)&sProps,
			ulFlags,
			lpcbEntryID,
			lppEntryID));

		lpXManageStoreEx->Release();
	}

	return hRes;
}

_Check_return_ HRESULT CreateStoreEntryID(
	_In_ LPMDB lpMDB, // open message store
	_In_z_ LPCTSTR lpszMsgStoreDN, // desired message store DN
	_In_opt_z_ LPCTSTR lpszMailboxDN, // desired mailbox DN or NULL
	ULONG ulFlags, // desired flags for CreateStoreEntryID
	bool bForceServer, // Use CreateStoreEntryID2
	_Out_opt_ ULONG* lpcbEntryID,
	_Deref_out_opt_ LPENTRYID * lppEntryID)
{
	HRESULT hRes = S_OK;

#ifdef UNICODE
	LPSTR szAnsiMsgStoreDN = NULL;
	LPSTR szAnsiMailboxDN = NULL;
	if (lpszMsgStoreDN) EC_H(UnicodeToAnsi(lpszMsgStoreDN, &szAnsiMsgStoreDN));
	if (lpszMailboxDN) EC_H(UnicodeToAnsi(lpszMailboxDN, &szAnsiMailboxDN));
#else
	LPCSTR szAnsiMsgStoreDN = lpszMsgStoreDN;
	LPCSTR szAnsiMailboxDN = lpszMailboxDN;
#endif

	// Use a NULL MailboxDN to open the public store
	if (szAnsiMailboxDN == NULL || !*szAnsiMailboxDN)
	{
		ulFlags |= OPENSTORE_PUBLIC;
	}

	if (!bForceServer)
	{
		EC_MAPI(CreateStoreEntryID(
			lpMDB,
			(LPSTR)szAnsiMsgStoreDN,
			(LPSTR)szAnsiMailboxDN,
			ulFlags,
			lpcbEntryID,
			lppEntryID));
	}
	else
	{
		EC_MAPI(CreateStoreEntryID2(
			lpMDB,
			(LPSTR)szAnsiMsgStoreDN,
			(LPSTR)szAnsiMailboxDN,
			ulFlags,
			lpcbEntryID,
			lppEntryID));
	}

#ifdef UNICODE
	delete[] szAnsiMsgStoreDN;
	delete[] szAnsiMailboxDN;
#endif

	return hRes;
}

// Stolen from MBLogon.c in the EDK to avoid compiling and linking in the entire beast
// Cleaned up to fit in with other functions
// $--HrMailboxLogon------------------------------------------------------
// Logon to a mailbox.  Before calling this function do the following:
//  1) Create a profile that has Exchange administrator privileges.
//  2) Logon to Exchange using this profile.
//  3) Open the mailbox using the Message Store DN and Mailbox DN.
//
// This version of the function needs the server and mailbox names to be
// in the form of distinguished names.  They would look something like this:
//		/CN=Configuration/CN=Servers/CN=%s/CN=Microsoft Private MDB
//		/CN=Configuration/CN=Servers/CN=%s/CN=Microsoft Public MDB
//		/O=<Organization>/OU=<Site>/CN=<Container>/CN=<MailboxName>
// where items in <brackets> would need to be set to appropriate values
//
// Note1: The message store DN is nearly identical to the server DN, except
// for the addition of a trailing '/CN=' part.  This part is required although
// its actual value is ignored.
//
// Note2: A NULL lpszMailboxDN indicates the public store should be opened.
// -----------------------------------------------------------------------------

_Check_return_ HRESULT HrMailboxLogon(
	_In_ LPMAPISESSION lpMAPISession, // MAPI session handle
	_In_ LPMDB lpMDB, // open message store
	_In_z_ LPCTSTR lpszMsgStoreDN, // desired message store DN
	_In_opt_z_ LPCTSTR lpszMailboxDN, // desired mailbox DN or NULL
	ULONG ulFlags, // desired flags for CreateStoreEntryID
	bool bForceServer, // Use CreateStoreEntryID2
	_Deref_out_opt_ LPMDB*	lppMailboxMDB) // ptr to mailbox message store ptr
{
	HRESULT hRes = S_OK;
	SBinary sbEID = { 0 };

	*lppMailboxMDB = NULL;

	if (!lpMAPISession)
	{
		DebugPrint(DBGGeneric, L"HrMailboxLogon: Session was NULL\n");
		return MAPI_E_INVALID_PARAMETER;
	}

	WC_H(CreateStoreEntryID(lpMDB, lpszMsgStoreDN, lpszMailboxDN, ulFlags, bForceServer, &sbEID.cb, (LPENTRYID*)&sbEID.lpb));

	if (SUCCEEDED(hRes))
	{
		WC_H(CallOpenMsgStore(
			lpMAPISession,
			NULL,
			&sbEID,
			MDB_NO_DIALOG |
			MDB_NO_MAIL |     // spooler not notified of our presence
			MDB_TEMPORARY |   // message store not added to MAPI profile
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
	HRESULT				hRes = S_OK;
	LPMAPITABLE			pStoresTbl = NULL;
	LPSRowSet			pRow = NULL;
	static SRestriction	sres;
	SPropValue			spv;

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
		pStoresTbl,						// table to query
		(LPSPropTagArray)&sptEIDCol,	// columns to get
		&sres,							// restriction to use
		NULL,							// sort order
		0,								// max number of rows - 0 means ALL
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
} // OpenDefaultMessageStore

// Build DNs for call to HrMailboxLogon
_Check_return_ HRESULT OpenOtherUsersMailbox(
	_In_ LPMAPISESSION lpMAPISession,
	_In_ LPMDB lpMDB,
	_In_opt_z_ LPCTSTR szServerName,
	_In_z_ LPCTSTR szMailboxDN,
	ULONG ulFlags, // desired flags for CreateStoreEntryID
	bool bForceServer, // Use CreateStoreEntryID2
	_Deref_out_opt_ LPMDB* lppOtherUserMDB)
{
	HRESULT		hRes = S_OK;

	LPTSTR		szServerNameFromProfile = NULL;
	LPCTSTR		szServerNamePTR = NULL; // Just a pointer. Do not free.

	*lppOtherUserMDB = NULL;

	DebugPrint(DBGGeneric, L"OpenOtherUsersMailbox called with lpMAPISession = %p, lpMDB = %p, Server = \"%ws\", Mailbox = \"%ws\"\n", lpMAPISession, lpMDB, LPCTSTRToWstring(szServerName).c_str(), LPCTSTRToWstring(szMailboxDN).c_str());
	if (!lpMAPISession || !lpMDB || !szMailboxDN || !StoreSupportsManageStore(lpMDB)) return MAPI_E_INVALID_PARAMETER;

	szServerNamePTR = szServerName;
	if (!szServerName)
	{
		// If we weren't given a server name, get one from the profile
		EC_H(GetServerName(lpMAPISession, &szServerNameFromProfile));
		szServerNamePTR = szServerNameFromProfile;
	}

	if (szServerNamePTR)
	{
		LPTSTR	szServerDN = NULL;

		EC_H(BuildServerDN(
			szServerNamePTR,
			_T("/cn=Microsoft Private MDB"), // STRING_OK
			&szServerDN));

		if (szServerDN)
		{
			DebugPrint(DBGGeneric, L"Calling HrMailboxLogon with Server DN = \"%ws\"\n", LPCTSTRToWstring(szServerDN).c_str());
			WC_H(HrMailboxLogon(
				lpMAPISession,
				lpMDB,
				szServerDN,
				szMailboxDN,
				ulFlags,
				bForceServer,
				lppOtherUserMDB));
			MAPIFreeBuffer(szServerDN);
		}
	}

	MAPIFreeBuffer(szServerNameFromProfile);
	return hRes;
} // OpenOtherUsersMailbox

#ifndef MRMAPI
_Check_return_ HRESULT OpenMailboxWithPrompt(
	_In_ LPMAPISESSION lpMAPISession,
	_In_ LPMDB lpMDB,
	_In_opt_z_ LPCTSTR szServerName,
	_In_opt_z_ LPCTSTR szMailboxDN,
	ULONG ulFlags, // desired flags for CreateStoreEntryID
	_Deref_out_opt_ LPMDB* lppOtherUserMDB)
{
	HRESULT hRes = S_OK;
	*lppOtherUserMDB = NULL;

	if (!lpMAPISession) return MAPI_E_INVALID_PARAMETER;

	CEditor MyPrompt(
		NULL,
		IDS_OPENOTHERUSER,
		IDS_OPENWITHFLAGSPROMPT,
		4,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyPrompt.SetPromptPostFix(AllFlagsToString(PROP_ID(PR_PROFILE_OPEN_FLAGS), true));
	MyPrompt.InitPane(0, CreateSingleLinePane(IDS_SERVERNAME, szServerName, false));
	MyPrompt.InitPane(1, CreateSingleLinePane(IDS_USERDN, szMailboxDN, false));
	MyPrompt.InitPane(2, CreateSingleLinePane(IDS_CREATESTORENTRYIDFLAGS, NULL, false));
	MyPrompt.SetHex(2, ulFlags);
	MyPrompt.InitPane(3, CreateCheckPane(IDS_FORCESERVER, false, false));
	WC_H(MyPrompt.DisplayDialog());
	if (S_OK == hRes)
	{
		WC_H(OpenOtherUsersMailbox(
			lpMAPISession,
			lpMDB,
			MyPrompt.GetString(0),
			MyPrompt.GetString(1),
			MyPrompt.GetHex(2),
			MyPrompt.GetCheck(3),
			lppOtherUserMDB));
	}

	return hRes;
} // OpenMailboxWithPrompt

// Display a UI to select a mailbox, then call OpenOtherUsersMailbox with the mailboxDN
// May return MAPI_E_CANCEL
_Check_return_ HRESULT OpenOtherUsersMailboxFromGal(
	_In_ LPMAPISESSION lpMAPISession,
	_In_ LPADRBOOK lpAddrBook,
	_Deref_out_opt_ LPMDB* lppOtherUserMDB)
{
	HRESULT hRes = S_OK;

	LPSPropValue lpEmailAddress = NULL;
	LPMAILUSER lpMailUser = NULL;
	LPMDB lpPrivateMDB = NULL;
	LPTSTR szServerName = NULL;

	*lppOtherUserMDB = NULL;

	if (!lpMAPISession || !lpAddrBook) return MAPI_E_INVALID_PARAMETER;

	EC_H(GetServerName(lpMAPISession, &szServerName));

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
					lpEmailAddress->Value.LPSZ,
					OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_TAKE_OWNERSHIP,
					lppOtherUserMDB));
			}
		}
	}

	MAPIFreeBuffer(szServerName);
	MAPIFreeBuffer(lpEmailAddress);
	if (lpMailUser) lpMailUser->Release();
	if (lpPrivateMDB) lpPrivateMDB->Release();
	return hRes;
} // OpenOtherUsersMailboxFromGal
#endif

// Use these guids:
// pbExchangeProviderPrimaryUserGuid
// pbExchangeProviderDelegateGuid
// pbExchangeProviderPublicGuid
// pbExchangeProviderXportGuid

_Check_return_ HRESULT OpenMessageStoreGUID(_In_ LPMAPISESSION lpMAPISession,
	_In_z_ LPCSTR lpGUID,
	_Deref_out_opt_ LPMDB* lppMDB)
{
	LPMAPITABLE	pStoresTbl = NULL;
	LPSRowSet	pRow = NULL;
	ULONG		ulRowNum;
	HRESULT		hRes = S_OK;

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

	*lppMDB = NULL;
	if (!lpMAPISession) return MAPI_E_INVALID_PARAMETER;

	EC_MAPI(lpMAPISession->GetMsgStoresTable(0, &pStoresTbl));

	if (pStoresTbl)
	{
		EC_MAPI(HrQueryAllRows(
			pStoresTbl,					// table to query
			(LPSPropTagArray)&sptCols,	// columns to get
			NULL,						// restriction to use
			NULL,						// sort order
			0,							// max number of rows
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
} // OpenMessageStoreGUID

_Check_return_ HRESULT OpenPublicMessageStore(
	_In_ LPMAPISESSION lpMAPISession,
	ULONG ulFlags, // Flags for CreateStoreEntryID
	_Deref_out_opt_ LPMDB* lppPublicMDB)
{
	HRESULT			hRes = S_OK;

	LPMDB			lpPublicMDBNonAdmin = NULL;
	LPSPropValue	lpServerName = NULL;

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
			LPTSTR	szServerDN = NULL;

			EC_H(BuildServerDN(
				lpServerName->Value.LPSZ,
				_T("/cn=Microsoft Public MDB"), // STRING_OK
				&szServerDN));

			if (szServerDN)
			{
				WC_H(HrMailboxLogon(
					lpMAPISession,
					lpPublicMDBNonAdmin,
					szServerDN,
					NULL,
					ulFlags,
					false,
					lppPublicMDB));
				MAPIFreeBuffer(szServerDN);
			}
		}
	}

	MAPIFreeBuffer(lpServerName);
	if (lpPublicMDBNonAdmin) lpPublicMDBNonAdmin->Release();
	return hRes;
} // OpenPublicMessageStore

_Check_return_ HRESULT OpenStoreFromMAPIProp(_In_ LPMAPISESSION lpMAPISession, _In_ LPMAPIPROP lpMAPIProp, _Deref_out_ LPMDB* lpMDB)
{
	HRESULT hRes = S_OK;
	LPSPropValue lpProp = NULL;

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
} // OpenStoreFromMAPIProp

_Check_return_ bool StoreSupportsManageStore(_In_ LPMDB lpMDB)
{
	HRESULT hRes = S_OK;
	LPEXCHANGEMANAGESTORE lpIManageStore = NULL;

	if (!lpMDB) return false;

	EC_H_MSG(lpMDB->QueryInterface(
		IID_IExchangeManageStore,
		(void **)&lpIManageStore),
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
	HRESULT hRes = S_OK;
	LPEXCHANGEMANAGESTORE lpIManageStore = NULL;

	if (!lpMDB) return false;

	EC_H_MSG(lpMDB->QueryInterface(
		IID_IExchangeManageStoreEx,
		(void **)&lpIManageStore),
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
	HRESULT hRes = S_OK;
	IProxyStoreObject* lpProxyObj = NULL;
	LPMDB lpUnwrappedMDB = NULL;
	EC_MAPI(lpMDBIn->QueryInterface(IID_IProxyStoreObject, (void**)&lpProxyObj));
	if (SUCCEEDED(hRes) && lpProxyObj)
	{
		EC_MAPI(lpProxyObj->UnwrapNoRef((LPVOID*)&lpUnwrappedMDB));
		if (SUCCEEDED(hRes) && lpUnwrappedMDB)
		{
			// UnwrapNoRef doesn't addref, so we do it here:
			lpUnwrappedMDB->AddRef();
			(*lppMDBOut) = lpUnwrappedMDB;
		}
	}
	if (lpProxyObj) lpProxyObj->Release();
	return hRes;
} // HrUnWrapMDB