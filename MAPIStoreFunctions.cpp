// MAPIStorefunctions.cpp : Collection of useful MAPI Store functions

#include "stdafx.h"
#include "MAPIStoreFunctions.h"
#include "MAPIFunctions.h"
#include "Editor.h"
#include "Guids.h"
#include "InterpretProp2.h"

HRESULT CallOpenMsgStore(
						 LPMAPISESSION	lpSession,
						 ULONG_PTR		ulUIParam,
						 LPSBinary		lpEID,
						 ULONG			ulFlags,
						 LPMDB*			lpMDB)
{
	DebugPrint(DBGOpenItemProp,_T("CallOpenMsgStore ulFlags = 0x%X\n"),ulFlags);
	if (!lpSession || !lpMDB || !lpEID) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;

	if (RegKeys[regkeyMDB_ONLINE].ulCurDWORD)
	{
		ulFlags |= MDB_ONLINE;
	}

	WC_H(lpSession->OpenMsgStore(
		ulUIParam,
		lpEID->cb,
		(LPENTRYID) lpEID->lpb,
		NULL,
		ulFlags,
		(LPMDB*) lpMDB));
	if (MAPI_E_UNKNOWN_FLAGS == hRes && (ulFlags & MDB_ONLINE))
	{
		hRes = S_OK;
		// perhaps this store doesn't know the MDB_ONLINE flag - remove and retry
		ulFlags = ulFlags & ~MDB_ONLINE;

		EC_H(lpSession->OpenMsgStore(
			ulUIParam,
			lpEID->cb,
			(LPENTRYID) lpEID->lpb,
			NULL,
			ulFlags,
			(LPMDB*) lpMDB));
	}
	else CHECKHRES(hRes);
	return hRes;
}
// Build a server DN. Allocates memory. Free with MAPIFreeBuffer.
HRESULT BuildServerDN(
					  LPCTSTR szServerName,
					  LPCTSTR szPost,
					  LPTSTR* lpszServerDN)
{
	HRESULT hRes = S_OK;
	if (!lpszServerDN) return MAPI_E_INVALID_PARAMETER;

	static LPTSTR szPre = _T("/cn=Configuration/cn=Servers/cn="); // STRING_OK
	size_t cbPreLen = 0;
	size_t cbServerLen = 0;
	size_t cbPostLen = 0;
	size_t cbServerDN = 0;

	EC_H(StringCbLength(szPre,STRSAFE_MAX_CCH * sizeof(TCHAR),&cbPreLen));
	EC_H(StringCbLength(szServerName,STRSAFE_MAX_CCH * sizeof(TCHAR),&cbServerLen));
	EC_H(StringCbLength(szPost,STRSAFE_MAX_CCH * sizeof(TCHAR),&cbPostLen));

	cbServerDN = cbPreLen + cbServerLen + cbPostLen + sizeof(TCHAR);

	EC_H(MAPIAllocateBuffer(
		(ULONG) cbServerDN,
		(LPVOID*)lpszServerDN));

	EC_H(StringCbPrintf(
		*lpszServerDN,
		cbServerDN,
		_T("%s%s%s"), // STRING_OK
		szPre,
		szServerName,
		szPost));
	return hRes;
}

HRESULT GetMailboxTable1(
						LPMDB lpMDB,
						LPCTSTR szServerDN,
						ULONG ulFlags,
						LPMAPITABLE* lpMailboxTable)
{
	if (!lpMDB || !lpMailboxTable || !szServerDN) return MAPI_E_INVALID_PARAMETER;
	*lpMailboxTable = NULL;

	HRESULT	hRes = S_OK;
	LPEXCHANGEMANAGESTORE lpManageStore1 = NULL;

	WC_H(lpMDB->QueryInterface(
		IID_IExchangeManageStore,
		(void **) &lpManageStore1));

	if (lpManageStore1)
	{
		WC_H(lpManageStore1->GetMailboxTable(
			(LPSTR) szServerDN,
			lpMailboxTable,
			ulFlags));

		lpManageStore1->Release();
	}
	return hRes;
} // GetMailboxTable1

HRESULT GetMailboxTable3(
						LPMDB lpMDB,
						LPCTSTR szServerDN,
						ULONG ulOffset,
						ULONG ulFlags,
						LPMAPITABLE* lpMailboxTable)
{
	if (!lpMDB || !lpMailboxTable || !szServerDN) return MAPI_E_INVALID_PARAMETER;
	*lpMailboxTable = NULL;

	HRESULT	hRes = S_OK;
	LPEXCHANGEMANAGESTORE3 lpManageStore3 = NULL;

	WC_H(lpMDB->QueryInterface(
		IID_IExchangeManageStore3,
		(void **) &lpManageStore3));

	if (lpManageStore3)
	{
		WC_H(lpManageStore3->GetMailboxTableOffset(
			(LPSTR) szServerDN,
			lpMailboxTable,
			ulFlags,
			ulOffset));

		lpManageStore3->Release();
	}
	return hRes;
} // GetMailboxTable3


HRESULT GetMailboxTable5(
						LPMDB lpMDB,
						LPCTSTR szServerDN,
						ULONG ulOffset,
						ULONG ulFlags,
						LPGUID lpGuidMDB,
						LPMAPITABLE* lpMailboxTable)
{
	if (!lpMDB || !lpMailboxTable || !szServerDN) return MAPI_E_INVALID_PARAMETER;
	*lpMailboxTable = NULL;

	HRESULT	hRes = S_OK;
	LPEXCHANGEMANAGESTORE5 lpManageStore5 = NULL;

	EC_H(lpMDB->QueryInterface(
		IID_IExchangeManageStore5,
		(void **) &lpManageStore5));

	if (lpManageStore5)
	{
		EC_H(lpManageStore5->GetMailboxTableEx(
			(LPSTR) szServerDN,
			lpGuidMDB,
			lpMailboxTable,
			ulFlags,
			ulOffset));

		lpManageStore5->Release();
	}
	return hRes;
} // GetMailboxTable5

// lpMDB needs to be an Exchange MDB - OpenPrivateMessageStore can get one if there's one to be had
// Use GetServerName to get the default server
// Will try IID_IExchangeManageStore3 first and fail back to IID_IExchangeManageStore
HRESULT GetMailboxTable(
						LPMDB lpMDB,
						LPCTSTR szServerName,
						ULONG ulOffset,
						LPMAPITABLE* lpMailboxTable)
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

HRESULT GetPublicFolderTable1(
						LPMDB lpMDB,
						LPCTSTR szServerDN,
						ULONG ulFlags,
						LPMAPITABLE* lpPFTable)
{
	if (!lpMDB || !lpPFTable || !szServerDN) return MAPI_E_INVALID_PARAMETER;
	*lpPFTable = NULL;

	HRESULT	hRes = S_OK;
	LPEXCHANGEMANAGESTORE lpManageStore1 = NULL;

	EC_H(lpMDB->QueryInterface(
		IID_IExchangeManageStore,
		(void **) &lpManageStore1));

	if (lpManageStore1)
	{
		EC_H(lpManageStore1->GetPublicFolderTable(
			(LPSTR) szServerDN,
			lpPFTable,
			ulFlags));

		lpManageStore1->Release();
	}

	return hRes;
} // GetPublicFolderTable1

HRESULT GetPublicFolderTable4(
						LPMDB lpMDB,
						LPCTSTR szServerDN,
						ULONG ulOffset,
						ULONG ulFlags,
						LPMAPITABLE* lpPFTable)
{
	if (!lpMDB || !lpPFTable || !szServerDN) return MAPI_E_INVALID_PARAMETER;
	*lpPFTable = NULL;

	HRESULT	hRes = S_OK;
	LPEXCHANGEMANAGESTORE4 lpManageStore4 = NULL;

	EC_H(lpMDB->QueryInterface(
		IID_IExchangeManageStore4,
		(void **) &lpManageStore4));

	if (lpManageStore4)
	{
		EC_H(lpManageStore4->GetPublicFolderTableOffset(
			(LPSTR) szServerDN,
			lpPFTable,
			ulFlags,
			ulOffset));
		lpManageStore4->Release();
	}

	return hRes;
} // GetPublicFolderTable4

HRESULT GetPublicFolderTable5(
						LPMDB lpMDB,
						LPCTSTR szServerDN,
						ULONG ulOffset,
						ULONG ulFlags,
						LPGUID lpGuidMDB,
						LPMAPITABLE* lpPFTable)
{
	if (!lpMDB || !lpPFTable || !szServerDN) return MAPI_E_INVALID_PARAMETER;
	*lpPFTable = NULL;

	HRESULT	hRes = S_OK;
	LPEXCHANGEMANAGESTORE5 lpManageStore5 = NULL;

	EC_H(lpMDB->QueryInterface(
		IID_IExchangeManageStore5,
		(void **) &lpManageStore5));

	if (lpManageStore5)
	{
		EC_H(lpManageStore5->GetPublicFolderTableEx(
			(LPSTR) szServerDN,
			lpGuidMDB,
			lpPFTable,
			ulFlags,
			ulOffset));

		lpManageStore5->Release();
	}

	return hRes;
} // GetPublicFolderTable5

// Get server name from the profile
HRESULT GetServerName(LPMAPISESSION lpSession, LPTSTR* szServerName)
{
	HRESULT			hRes = S_OK;
	LPSERVICEADMIN	pSvcAdmin = NULL;
	LPPROFSECT		pGlobalProfSect = NULL;
	LPSPropValue	lpServerName	= NULL;

	if (!lpSession) return MAPI_E_INVALID_PARAMETER;

	*szServerName = NULL;

	EC_H(lpSession->AdminServices(
		0,
		&pSvcAdmin));

	EC_H(pSvcAdmin->OpenProfileSection(
		(LPMAPIUID)pbGlobalProfileSectionGuid,
		NULL,
		0,
		&pGlobalProfSect));

	EC_H(HrGetOneProp(pGlobalProfSect,
		PR_PROFILE_HOME_SERVER,
		&lpServerName));

	if (CheckStringProp(lpServerName,PT_STRING8)) // profiles are ASCII only
	{
#ifdef UNICODE
		LPWSTR	szWideServer = NULL;
		EC_H(AnsiToUnicode(
			lpServerName->Value.lpszA,
			&szWideServer));
		EC_H(CopyStringW(szServerName,szWideServer,NULL));
		delete[] szWideServer;
#else
		EC_H(CopyStringA(szServerName,lpServerName->Value.lpszA,NULL));
#endif
	}
	else
	{
		// prompt the user to enter a server name
		CEditor MyData(
			NULL,
			IDS_SERVERNAME,
			IDS_SERVERNAMEMISSINGPROMPT,
			1,
			CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
		MyData.InitSingleLine(0,IDS_SERVERNAME,NULL,false);

		WC_H(MyData.DisplayDialog());

		if (S_OK == hRes)
		{
			EC_H(CopyString(szServerName,MyData.GetString(0),NULL));
		}
	}
	MAPIFreeBuffer(lpServerName);
	if (pGlobalProfSect) pGlobalProfSect->Release();
	if (pSvcAdmin) pSvcAdmin->Release();
	return hRes;
} // GetServerName

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

HRESULT HrMailboxLogon(
	LPMAPISESSION	lpMAPISession,	// MAPI session handle
	LPMDB			lpMDB,			// open message store
	LPCTSTR			lpszMsgStoreDN,	// desired message store DN
	LPCTSTR			lpszMailboxDN,	// desired mailbox DN or NULL
	ULONG			ulFlags,		// desired flags for CreateStoreEntryID
	LPMDB			*lppMailboxMDB)	// ptr to mailbox message store ptr
{
	HRESULT					hRes			= S_OK;
	LPEXCHANGEMANAGESTORE	lpXManageStore  = NULL;
	LPMDB					lpMailboxMDB	= NULL;
	SBinary					sbEID			= {0};

	*lppMailboxMDB = NULL;

	if (!lpMAPISession || !lpMDB || !lpszMsgStoreDN || !lppMailboxMDB || !StoreSupportsManageStore(lpMDB))
	{
		if (!lpMAPISession) DebugPrint(DBGGeneric,_T("HrMailboxLogon: Session was NULL\n"));
		if (!lpMDB) DebugPrint(DBGGeneric,_T("HrMailboxLogon: MDB was NULL\n"));
		if (!lpszMsgStoreDN) DebugPrint(DBGGeneric,_T("HrMailboxLogon: lpszMsgStoreDN was NULL\n"));
		if (!lppMailboxMDB) DebugPrint(DBGGeneric,_T("HrMailboxLogon: lppMailboxMDB was NULL\n"));
		return MAPI_E_INVALID_PARAMETER;
	}

	// Use a NULL MailboxDN to open the public store
	if (lpszMailboxDN == NULL || !*lpszMailboxDN)
	{
		ulFlags |= OPENSTORE_PUBLIC;
	}

	EC_H(lpMDB->QueryInterface(
		IID_IExchangeManageStore,
		(LPVOID*) &lpXManageStore));

	if (lpXManageStore)
	{
		DebugPrint(DBGGeneric,_T("HrMailboxLogon: Creating EntryID. StoreDN = \"%s\", MailboxDN = \"%s\"\n"),lpszMsgStoreDN,lpszMailboxDN);

#ifdef UNICODE
		{
			char *szAnsiMsgStoreDN = NULL;
			char *szAnsiMailboxDN = NULL;
			EC_H(UnicodeToAnsi(lpszMsgStoreDN,&szAnsiMsgStoreDN));

			EC_H(UnicodeToAnsi(lpszMailboxDN,&szAnsiMailboxDN));

			EC_H(lpXManageStore->CreateStoreEntryID(
				szAnsiMsgStoreDN,
				szAnsiMailboxDN,
				ulFlags,
				&sbEID.cb,
				(LPENTRYID*) &sbEID.lpb));
			delete[] szAnsiMsgStoreDN;
			delete[] szAnsiMailboxDN;
		}
#else
		EC_H(lpXManageStore->CreateStoreEntryID(
			(LPSTR) lpszMsgStoreDN,
			(LPSTR) lpszMailboxDN,
			ulFlags,
			&sbEID.cb,
			(LPENTRYID*) &sbEID.lpb));
#endif
		DebugPrintBinary(DBGGeneric,&sbEID);

		EC_H(CallOpenMsgStore(
			lpMAPISession,
			NULL,
			&sbEID,
			MDB_NO_DIALOG |
			MDB_NO_MAIL |		// spooler not notified of our presence
			MDB_TEMPORARY |	 // message store not added to MAPI profile
			MAPI_BEST_ACCESS,	// normally WRITE, but allow access to RO store
			&lpMailboxMDB));

		*lppMailboxMDB = lpMailboxMDB;
	}

	MAPIFreeBuffer(sbEID.lpb);
	if (lpXManageStore) lpXManageStore->Release();
	return hRes;
}

HRESULT	OpenDefaultMessageStore(
								LPMAPISESSION lpMAPISession,
								LPMDB* lppDefaultMDB)
{
	HRESULT				hRes = S_OK;
	LPMAPITABLE			pStoresTbl = NULL;
	LPSRowSet			pRow = NULL;
	static SRestriction	sres;
	SPropValue			spv;

	enum {EID,NUM_COLS};
	static SizedSPropTagArray(NUM_COLS,sptEIDCol) = {NUM_COLS,
		PR_ENTRYID,
	};
	if (!lpMAPISession) return MAPI_E_INVALID_PARAMETER;


	EC_H(lpMAPISession->GetMsgStoresTable(0, &pStoresTbl));

	// set up restriction for the default store
	sres.rt = RES_PROPERTY; // gonna compare a property
	sres.res.resProperty.relop = RELOP_EQ; // gonna test equality
	sres.res.resProperty.ulPropTag = PR_DEFAULT_STORE; // tag to compare
	sres.res.resProperty.lpProp = &spv; // prop tag to compare against

	spv.ulPropTag = PR_DEFAULT_STORE; // tag type
	spv.Value.b	= TRUE; // tag value

	EC_H(HrQueryAllRows(
		pStoresTbl,						// table to query
		(LPSPropTagArray) &sptEIDCol,	// columns to get
		&sres,							// restriction to use
		NULL,							// sort order
		0,								// max number of rows - 0 means ALL
		&pRow));

	if (pRow && pRow->cRows)
	{
		EC_H(CallOpenMsgStore(
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

// Build DN's for call to HrMailboxLogon
HRESULT OpenOtherUsersMailbox(
							  LPMAPISESSION	lpMAPISession,
							  LPMDB lpMDB,
							  LPCTSTR szServerName,
							  LPCTSTR szMailboxDN,
							  ULONG ulFlags, // desired flags for CreateStoreEntryID
							  LPMDB* lppOtherUserMDB)
{
	HRESULT		 hRes			= S_OK;

	LPTSTR		szServerNameFromProfile = NULL;
	LPCTSTR		szServerNamePTR = NULL; // Just a pointer. Do not free.

	*lppOtherUserMDB = NULL;

	DebugPrint(DBGGeneric,_T("OpenOtherUsersMailbox called with lpMAPISession = 0x%08X, lpMDB = 0x%08X, Server = \"%s\", Mailbox = \"%s\"\n"),lpMAPISession, lpMDB,szServerName,szMailboxDN);
	if (!lpMAPISession || !lpMDB || !szMailboxDN || !StoreSupportsManageStore(lpMDB)) return MAPI_E_INVALID_PARAMETER;

	szServerNamePTR = szServerName;
	if (!szServerName)
	{
		// If we weren't given a server name, get one from the profile
		EC_H(GetServerName(lpMAPISession,&szServerNameFromProfile));
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
			DebugPrint(DBGGeneric,_T("Calling HrMailboxLogon with Server DN = \"%s\"\n"),szServerDN);
			EC_H(HrMailboxLogon(
				lpMAPISession,
				lpMDB,
				szServerDN,
				szMailboxDN,
				ulFlags,
				lppOtherUserMDB));
			MAPIFreeBuffer(szServerDN);
		}
	}

	MAPIFreeBuffer(szServerNameFromProfile);
	return hRes;
}

// Display a UI to select a mailbox, then call OpenOtherUsersMailbox with the mailboxDN
HRESULT OpenOtherUsersMailboxFromGal(
									 LPMAPISESSION	lpMAPISession,
									 LPADRBOOK lpAddrBook,
									 LPMDB* lppOtherUserMDB)
{
	HRESULT		 hRes			= S_OK;

	ADRPARM			AdrParm = {0};
	LPADRLIST		lpAdrList		= NULL;
	LPSPropValue	lpEmailAddress	= NULL;
	LPSPropValue	lpEntryID		= NULL;
	LPMAILUSER		lpMailUser		= NULL;
	LPMDB			lpPrivateMDB	= NULL;

	*lppOtherUserMDB = NULL;

	if (!lpMAPISession || !lpAddrBook) return MAPI_E_INVALID_PARAMETER;

	// CString doesn't provide a way to extract just ANSI strings, so we do this manually
	CHAR szTitle[256];
	int iRet = NULL;
	EC_D(iRet,LoadStringA(GetModuleHandle(NULL),
		IDS_SELECTMAILBOX,
		szTitle,
		sizeof(szTitle)/sizeof(CHAR)));

	AdrParm.ulFlags			= DIALOG_MODAL | ADDRESS_ONE | AB_SELECTONLY | AB_RESOLVE ;
#pragma warning(push)
#pragma warning(disable:4616)
#pragma warning(disable:6276)
	AdrParm.lpszCaption		= (LPTSTR) szTitle;
#pragma warning(pop)

	ULONG_PTR ulHwnd = (ULONG_PTR)::GetDesktopWindow();

	EC_H(OpenMessageStoreGUID(lpMAPISession,pbExchangeProviderPrimaryUserGuid,&lpPrivateMDB));
	if (lpPrivateMDB && StoreSupportsManageStore(lpPrivateMDB))
	{
		EC_H_CANCEL(lpAddrBook->Address(
			&ulHwnd,
			&AdrParm,
			&lpAdrList));

		if (lpAdrList)
		{
			lpEntryID = PpropFindProp(
				lpAdrList[0].aEntries->rgPropVals,
				lpAdrList[0].aEntries->cValues,
				PR_ENTRYID);

			if (lpEntryID)
			{
				EC_H(CallOpenEntry(
					NULL,
					lpAddrBook,
					NULL,
					NULL,
					lpEntryID->Value.bin.cb,
					(LPENTRYID) lpEntryID->Value.bin.lpb,
					NULL,
					MAPI_BEST_ACCESS,
					NULL,
					(LPUNKNOWN*)&lpMailUser));
				if (lpMailUser)
				{
					EC_H(HrGetOneProp(
						lpMailUser,
						PR_EMAIL_ADDRESS,
						&lpEmailAddress));

					if (CheckStringProp(lpEmailAddress,PT_TSTRING))
					{
						CEditor MyPrompt(
							NULL,
							IDS_OPENOTHERUSER,
							IDS_OPENWITHFLAGSPROMPT,
							1,
							CEDITOR_BUTTON_OK|CEDITOR_BUTTON_CANCEL);
						MyPrompt.SetPromptPostFix(AllFlagsToString(PROP_ID(PR_PROFILE_OPEN_FLAGS),true));
						MyPrompt.InitSingleLine(0,IDS_CREATESTORENTRYIDFLAGS,NULL,false);
						MyPrompt.SetHex(0,OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_TAKE_OWNERSHIP);
						WC_H(MyPrompt.DisplayDialog());
						if (S_OK == hRes)
						{
							EC_H(OpenOtherUsersMailbox(
								lpMAPISession,
								lpPrivateMDB,
								NULL,
								lpEmailAddress->Value.LPSZ,
								MyPrompt.GetHex(0),
								lppOtherUserMDB));
						}
					}
				}
			}
		}
	}

	MAPIFreeBuffer(lpEmailAddress);
	if (lpAdrList) FreePadrlist(lpAdrList);
	if (lpMailUser) lpMailUser->Release();
	if (lpPrivateMDB) lpPrivateMDB->Release();
	return hRes;
} // OpenOtherUsersMailbox

// Use these guids:
// pbExchangeProviderPrimaryUserGuid
// pbExchangeProviderDelegateGuid
// pbExchangeProviderPublicGuid
// pbExchangeProviderXportGuid

HRESULT OpenMessageStoreGUID(
							  LPMAPISESSION	lpMAPISession,
							  LPCSTR lpGUID,
							  LPMDB* lppMDB)
{
	LPMAPITABLE	pStoresTbl = NULL;
	LPSRowSet	pRow		= NULL;
	ULONG		ulRowNum;
	HRESULT		hRes = S_OK;

	enum {EID,STORETYPE,NUM_COLS};
	static SizedSPropTagArray(NUM_COLS,sptCols) = {NUM_COLS,
		PR_ENTRYID,
		PR_MDB_PROVIDER
	};

	*lppMDB = NULL;
	if (!lpMAPISession) return MAPI_E_INVALID_PARAMETER;

	EC_H(lpMAPISession->GetMsgStoresTable(0, &pStoresTbl));

	if (pStoresTbl)
	{
		EC_H(HrQueryAllRows(
			pStoresTbl,					// table to query
			(LPSPropTagArray) &sptCols,	// columns to get
			NULL,						// restriction to use
			NULL,						// sort order
			0,							// max number of rows
			&pRow));
		if (pRow)
		{
			if (!FAILED(hRes)) for (ulRowNum=0; ulRowNum<pRow->cRows; ulRowNum++)
			{
				hRes = S_OK;
				// check to see if we have a folder with a matching GUID
				if (pRow->aRow[ulRowNum].lpProps[STORETYPE].ulPropTag == PR_MDB_PROVIDER &&
					pRow->aRow[ulRowNum].lpProps[EID].ulPropTag == PR_ENTRYID &&
					IsEqualMAPIUID(pRow->aRow[ulRowNum].lpProps[STORETYPE].Value.bin.lpb,lpGUID))
				{
					EC_H(CallOpenMsgStore(
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

HRESULT OpenPublicMessageStore(
							   LPMAPISESSION lpMAPISession,
							   ULONG ulFlags, // Flags for CreateStoreEntryID
							   LPMDB* lppPublicMDB)
{
	HRESULT			hRes = S_OK;

	LPMDB			lpPublicMDBNonAdmin	= NULL;
	LPSPropValue	lpServerName	= NULL;

	if (!lpMAPISession || !lppPublicMDB) return MAPI_E_INVALID_PARAMETER;

	EC_H(OpenMessageStoreGUID(
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
		EC_H(HrGetOneProp(
			lpPublicMDBNonAdmin,
			PR_HIERARCHY_SERVER,
			&lpServerName));

		if (CheckStringProp(lpServerName,PT_TSTRING))
		{
			LPTSTR	szServerDN = NULL;

			EC_H(BuildServerDN(
				lpServerName->Value.LPSZ,
				_T("/cn=Microsoft Public MDB"), // STRING_OK
				&szServerDN));

			if (szServerDN)
			{
				EC_H(HrMailboxLogon(
					lpMAPISession,
					lpPublicMDBNonAdmin,
					szServerDN,
					NULL,
					ulFlags,
					lppPublicMDB));
				MAPIFreeBuffer(szServerDN);
			}
		}
	}

	MAPIFreeBuffer(lpServerName);
	if (lpPublicMDBNonAdmin) lpPublicMDBNonAdmin->Release();
	return hRes;
} // OpenPublicMessageStore

HRESULT OpenStoreFromMAPIProp(LPMAPISESSION lpMAPISession, LPMAPIPROP lpMAPIProp, LPMDB* lpMDB)
{
	HRESULT hRes = S_OK;
	LPSPropValue lpProp = NULL;

	if (!lpMAPISession || !lpMAPIProp) return MAPI_E_INVALID_PARAMETER;

	EC_H(HrGetOneProp(
		lpMAPIProp,
		PR_STORE_ENTRYID,
		&lpProp));

	if (lpProp && PT_BINARY == PROP_TYPE(lpProp->ulPropTag))
	{
		EC_H(CallOpenMsgStore(
			lpMAPISession,
			NULL,
			&lpProp->Value.bin,
			MAPI_BEST_ACCESS,
			lpMDB));
	}

	MAPIFreeBuffer(lpProp);
	return hRes;
}

BOOL StoreSupportsManageStore(LPMDB lpMDB)
{
	HRESULT					hRes = S_OK;
	LPEXCHANGEMANAGESTORE	lpIManageStore = NULL;

	if (!lpMDB) return false;

	EC_H_MSG(lpMDB->QueryInterface(
		IID_IExchangeManageStore,
		(void **) &lpIManageStore),
		IDS_MANAGESTORENOTSUPPORTED);

	if (lpIManageStore)
	{
		lpIManageStore->Release();
		return true;
	}
	return false;
} // StoreSupportsManageStore

HRESULT HrUnWrapMDB(LPMDB lpMDBIn, LPMDB* lppMDBOut)
{
	if (!lpMDBIn || !lppMDBOut) return MAPI_E_INVALID_PARAMETER;
	HRESULT hRes = S_OK;
	IProxyStoreObject* lpProxyObj = NULL;
	LPMDB lpUnwrappedMDB = NULL;
	EC_H(lpMDBIn->QueryInterface(IID_IProxyStoreObject,(void**)&lpProxyObj));
	if (SUCCEEDED(hRes) && lpProxyObj)
	{
		EC_H(lpProxyObj->UnwrapNoRef((LPVOID*)&lpUnwrappedMDB));
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