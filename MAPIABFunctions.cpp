// MAPIABfunctions.cpp : Collection of useful MAPI Address Book functions

#include "stdafx.h"
#include "MAPIABFunctions.h"
#include "MAPIFunctions.h"

_Check_return_ HRESULT HrAllocAdrList(ULONG ulNumProps, _Deref_out_opt_ LPADRLIST* lpAdrList)
{
	if (!lpAdrList || ulNumProps > ULONG_MAX / sizeof(SPropValue)) return MAPI_E_INVALID_PARAMETER;
	HRESULT hRes = S_OK;
	LPADRLIST lpLocalAdrList = NULL;

	*lpAdrList = NULL;

	// Allocate memory for new SRowSet structure.
	EC_H(MAPIAllocateBuffer(CbNewSRowSet(1), (LPVOID*)&lpLocalAdrList));

	if (lpLocalAdrList)
	{
		// Zero out allocated memory.
		ZeroMemory(lpLocalAdrList, CbNewSRowSet(1));

		// Allocate memory for SPropValue structure that indicates what
		// recipient properties will be set.
		EC_H(MAPIAllocateBuffer(
			ulNumProps * sizeof(SPropValue),
			(LPVOID*)&lpLocalAdrList->aEntries[0].rgPropVals));

		// Zero out allocated memory.
		if (lpLocalAdrList->aEntries[0].rgPropVals)
			ZeroMemory(lpLocalAdrList->aEntries[0].rgPropVals, ulNumProps * sizeof(SPropValue));
		if (SUCCEEDED(hRes))
		{
			*lpAdrList = lpLocalAdrList;
		}
		else
		{
			FreePadrlist(lpLocalAdrList);
		}
	}

	return hRes;
} // HrAllocAdrList

_Check_return_ HRESULT AddOneOffAddress(
	_In_ LPMAPISESSION lpMAPISession,
	_In_ LPMESSAGE lpMessage,
	_In_z_ LPCTSTR szDisplayName,
	_In_z_ LPCTSTR szAddrType,
	_In_z_ LPCTSTR szEmailAddress,
	ULONG ulRecipientType)
{
	HRESULT hRes = S_OK;
	LPADRLIST lpAdrList = NULL; // ModifyRecips takes LPADRLIST
	LPADRBOOK lpAddrBook = NULL;
	LPENTRYID lpEID = NULL;

	enum
	{
		NAME,
		ADDR,
		EMAIL,
		RECIP,
		EID,
		NUM_RECIP_PROPS
	};

	if (!lpMessage || !lpMAPISession) return MAPI_E_INVALID_PARAMETER;

	EC_MAPI(lpMAPISession->OpenAddressBook(
		NULL,
		NULL,
		NULL,
		&lpAddrBook));

	EC_MAPI(HrAllocAdrList(NUM_RECIP_PROPS, &lpAdrList));

	// Setup the One Time recipient by indicating how many recipients
	// and how many properties will be set on each recipient.

	if (SUCCEEDED(hRes) && lpAdrList)
	{
		lpAdrList->cEntries = 1;	// How many recipients.
		lpAdrList->aEntries[0].cValues = NUM_RECIP_PROPS; // How many properties per recipient

		// Set the SPropValue members == the desired values.
		lpAdrList->aEntries[0].rgPropVals[NAME].ulPropTag = PR_DISPLAY_NAME;
		lpAdrList->aEntries[0].rgPropVals[NAME].Value.LPSZ = (LPTSTR)szDisplayName;

		lpAdrList->aEntries[0].rgPropVals[ADDR].ulPropTag = PR_ADDRTYPE;
		lpAdrList->aEntries[0].rgPropVals[ADDR].Value.LPSZ = (LPTSTR)szAddrType;

		lpAdrList->aEntries[0].rgPropVals[EMAIL].ulPropTag = PR_EMAIL_ADDRESS;
		lpAdrList->aEntries[0].rgPropVals[EMAIL].Value.LPSZ = (LPTSTR)szEmailAddress;

		lpAdrList->aEntries[0].rgPropVals[RECIP].ulPropTag = PR_RECIPIENT_TYPE;
		lpAdrList->aEntries[0].rgPropVals[RECIP].Value.l = ulRecipientType;

		lpAdrList->aEntries[0].rgPropVals[EID].ulPropTag = PR_ENTRYID;

		// Create the One-off address and get an EID for it.
		EC_MAPI(lpAddrBook->CreateOneOff(
			lpAdrList->aEntries[0].rgPropVals[NAME].Value.LPSZ,
			lpAdrList->aEntries[0].rgPropVals[ADDR].Value.LPSZ,
			lpAdrList->aEntries[0].rgPropVals[EMAIL].Value.LPSZ,
			fMapiUnicode,
			&lpAdrList->aEntries[0].rgPropVals[EID].Value.bin.cb,
			&lpEID));
		lpAdrList->aEntries[0].rgPropVals[EID].Value.bin.lpb = (LPBYTE)lpEID;

		EC_MAPI(lpAddrBook->ResolveName(
			0L,
			fMapiUnicode,
			NULL,
			lpAdrList));

		// If everything goes right, add the new recipient to the message
		// object passed into us.
		EC_MAPI(lpMessage->ModifyRecipients(MODRECIP_ADD, lpAdrList));

		EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
	}

	MAPIFreeBuffer(lpEID);
	if (lpAdrList) FreePadrlist(lpAdrList);
	if (lpAddrBook) lpAddrBook->Release();
	return hRes;
} // AddOneOffAddress

_Check_return_ HRESULT AddRecipient(
	_In_ LPMAPISESSION lpMAPISession,
	_In_ LPMESSAGE lpMessage,
	_In_z_ LPCTSTR szName,
	ULONG ulRecipientType)
{
	HRESULT			hRes = S_OK;
	LPADRLIST		lpAdrList = NULL; // ModifyRecips takes LPADRLIST
	LPADRBOOK		lpAddrBook = NULL;

	enum
	{
		NAME,
		RECIP,
		NUM_RECIP_PROPS
	};

	if (!lpMessage || !lpMAPISession) return MAPI_E_INVALID_PARAMETER;

	EC_MAPI(lpMAPISession->OpenAddressBook(
		NULL,
		NULL,
		NULL,
		&lpAddrBook));

	EC_MAPI(HrAllocAdrList(NUM_RECIP_PROPS, &lpAdrList));

	if (lpAdrList)
	{
		// Setup the One Time recipient by indicating how many recipients
		// and how many properties will be set on each recipient.
		lpAdrList->cEntries = 1;	// How many recipients.
		lpAdrList->aEntries[0].cValues = NUM_RECIP_PROPS; // How many properties per recipient

		// Set the SPropValue members == the desired values.
		lpAdrList->aEntries[0].rgPropVals[NAME].ulPropTag = PR_DISPLAY_NAME;
		lpAdrList->aEntries[0].rgPropVals[NAME].Value.LPSZ = (LPTSTR)szName;

		lpAdrList->aEntries[0].rgPropVals[RECIP].ulPropTag = PR_RECIPIENT_TYPE;
		lpAdrList->aEntries[0].rgPropVals[RECIP].Value.l = ulRecipientType;

		EC_MAPI(lpAddrBook->ResolveName(
			0L,
			fMapiUnicode,
			NULL,
			lpAdrList));

		// If everything goes right, add the new recipient to the message
		// object passed into us.
		EC_MAPI(lpMessage->ModifyRecipients(MODRECIP_ADD, lpAdrList));

		EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));
	}

	if (lpAdrList) FreePadrlist(lpAdrList);
	if (lpAddrBook) lpAddrBook->Release();
	return hRes;
} // AddRecipient

// Same as CreatePropertyStringRestriction, but skips the existence part.
_Check_return_ HRESULT CreateANRRestriction(ULONG ulPropTag,
	_In_z_ LPCTSTR szString,
	_In_opt_ LPVOID lpParent,
	_Deref_out_opt_ LPSRestriction* lppRes)
{
	HRESULT hRes = S_OK;
	LPSRestriction	lpRes = NULL;
	LPSPropValue	lpspvSubject = NULL;
	LPVOID			lpAllocationParent = NULL;

	*lppRes = NULL;

	// Allocate and create our SRestriction
	// Allocate base memory:
	if (lpParent)
	{
		EC_H(MAPIAllocateMore(
			sizeof(SRestriction),
			lpParent,
			(LPVOID*)&lpRes));

		lpAllocationParent = lpParent;
	}
	else
	{
		EC_H(MAPIAllocateBuffer(
			sizeof(SRestriction),
			(LPVOID*)&lpRes));

		lpAllocationParent = lpRes;
	}


	EC_H(MAPIAllocateMore(
		sizeof(SPropValue),
		lpAllocationParent,
		(LPVOID*)&lpspvSubject));

	if (lpRes && lpspvSubject)
	{
		// Zero out allocated memory.
		ZeroMemory(lpRes, sizeof(SRestriction));
		ZeroMemory(lpspvSubject, sizeof(SPropValue));

		// Root Node
		lpRes->rt = RES_PROPERTY;
		lpRes->res.resProperty.relop = RELOP_EQ;
		lpRes->res.resProperty.ulPropTag = ulPropTag;
		lpRes->res.resProperty.lpProp = lpspvSubject;

		// Allocate and fill out properties:
		lpspvSubject->ulPropTag = ulPropTag;
		lpspvSubject->Value.LPSZ = NULL;

		if (szString)
		{
			EC_H(CopyString(
				&lpspvSubject->Value.LPSZ,
				szString,
				lpAllocationParent));
		}

		DebugPrint(DBGGeneric, _T("CreateANRRestriction built restriction:\n"));
		DebugPrintRestriction(DBGGeneric, lpRes, NULL);

		*lppRes = lpRes;
	}

	if (hRes != S_OK)
	{
		DebugPrint(DBGGeneric, _T("Failed to create restriction\n"));
		MAPIFreeBuffer(lpRes);
		*lppRes = NULL;
	}
	return hRes;
} // CreateANRRestriction

_Check_return_ HRESULT GetABContainerTable(_In_ LPADRBOOK lpAdrBook, _Deref_out_opt_ LPMAPITABLE* lpABContainerTable)
{
	HRESULT		hRes = S_OK;
	LPABCONT	lpABRootContainer = NULL;
	LPMAPITABLE	lpTable = NULL;

	*lpABContainerTable = NULL;
	if (!lpAdrBook) return MAPI_E_INVALID_PARAMETER;

	// Open root address book (container).
	EC_H(CallOpenEntry(
		NULL,
		lpAdrBook,
		NULL,
		NULL,
		NULL,
		NULL,
		NULL,
		NULL,
		(LPUNKNOWN*)&lpABRootContainer));

	if (lpABRootContainer)
	{
		// Get a table of all of the Address Books.
		EC_MAPI(lpABRootContainer->GetHierarchyTable(CONVENIENT_DEPTH | fMapiUnicode, &lpTable));
		*lpABContainerTable = lpTable;
		lpABRootContainer->Release();
	}

	return hRes;
} // GetABContainerTable

// Manually resolve a name in the address book and add it to the message
_Check_return_ HRESULT ManualResolve(
	_In_ LPMAPISESSION lpMAPISession,
	_In_ LPMESSAGE lpMessage,
	_In_z_ LPCTSTR szName,
	ULONG PropTagToCompare)
{
	HRESULT			hRes = S_OK;
	ULONG			ulObjType = 0;
	LPADRBOOK		lpAdrBook = NULL;
	LPSRowSet		lpABRow = NULL;
	LPMAPITABLE		lpABContainerTable = NULL;
	LPADRLIST		lpAdrList = NULL;
	LPABCONT		lpABContainer = NULL;
	LPMAPITABLE		pTable = NULL;
	LPSPropValue	lpFoundRow = NULL;

	enum
	{
		abcPR_ENTRYID,
		abcPR_DISPLAY_NAME,
		abcNUM_COLS
	};

	static const SizedSPropTagArray(abcNUM_COLS, abcCols) =
	{
		abcNUM_COLS,
		PR_ENTRYID,
		PR_DISPLAY_NAME,
	};

	enum
	{
		abPR_ENTRYID,
		abPR_DISPLAY_NAME,
		abPR_RECIPIENT_TYPE,
		abPR_ADDRTYPE,
		abPR_DISPLAY_TYPE,
		abPropTagToCompare,
		abNUM_COLS
	};

	if (!lpMAPISession) return MAPI_E_INVALID_PARAMETER;

	DebugPrint(DBGGeneric, _T("ManualResolve: Asked to resolve \"%s\"\n"), szName);

	EC_MAPI(lpMAPISession->OpenAddressBook(
		NULL,
		NULL,
		NULL,
		&lpAdrBook));

	EC_H(GetABContainerTable(lpAdrBook, &lpABContainerTable));

	if (lpABContainerTable)
	{
		// Restrict the table to the properties that we are interested in.
		EC_MAPI(lpABContainerTable->SetColumns((LPSPropTagArray)&abcCols, TBL_BATCH));

		if (!FAILED(hRes)) for (;;)
		{
			hRes = S_OK;

			FreeProws(lpABRow);
			lpABRow = NULL;
			EC_MAPI(lpABContainerTable->QueryRows(
				1,
				NULL,
				&lpABRow));
			if (FAILED(hRes) || !lpABRow || (lpABRow && !lpABRow->cRows)) break;

			// From this point forward, consider any error an error with the current address book container, so just continue and try the next one.
			if (PR_ENTRYID == lpABRow->aRow->lpProps[abcPR_ENTRYID].ulPropTag)
			{
				DebugPrint(DBGGeneric, _T("ManualResolve: Searching this container\n"));
				DebugPrintBinary(DBGGeneric, &lpABRow->aRow->lpProps[abcPR_ENTRYID].Value.bin);

				if (lpABContainer) lpABContainer->Release();
				lpABContainer = NULL;
				EC_H(CallOpenEntry(
					NULL,
					lpAdrBook,
					NULL,
					NULL,
					lpABRow->aRow->lpProps[abcPR_ENTRYID].Value.bin.cb,
					(ENTRYID*)lpABRow->aRow->lpProps[abcPR_ENTRYID].Value.bin.lpb,
					NULL,
					NULL,
					&ulObjType,
					(LPUNKNOWN*)&lpABContainer));
				if (!lpABContainer) continue;

				DebugPrint(DBGGeneric, _T("ManualResolve: Object opened as 0x%X\n"), ulObjType);

				if (lpABContainer && ulObjType == MAPI_ABCONT)
				{
					if (pTable) pTable->Release();
					pTable = NULL;
					WC_MAPI(lpABContainer->GetContentsTable(fMapiUnicode, &pTable));
					if (!pTable)
					{
						DebugPrint(DBGGeneric, _T("ManualResolve: Container did not support contents table\n"));
						if (MAPI_E_NO_SUPPORT == hRes) hRes = S_OK;
						continue;
					}

					MAPIFreeBuffer(lpFoundRow);
					lpFoundRow = NULL;
					EC_H(SearchContentsTableForName(
						pTable,
						szName,
						PropTagToCompare,
						&lpFoundRow));
					if (!lpFoundRow) continue;

					if (lpAdrList) FreePadrlist(lpAdrList);
					lpAdrList = NULL;
					// Allocate memory for new Address List structure.
					EC_H(MAPIAllocateBuffer(CbNewADRLIST(1), (LPVOID*)&lpAdrList));
					if (!lpAdrList) continue;

					ZeroMemory(lpAdrList, CbNewADRLIST(1));
					lpAdrList->cEntries = 1;
					// Allocate memory for SPropValue structure that indicates what
					// recipient properties will be set. To resolve a name that
					// already exists in the Address book, this will always be 1.

					EC_H(MAPIAllocateBuffer(
						(ULONG)(abNUM_COLS * sizeof(SPropValue)),
						(LPVOID*)&lpAdrList->aEntries->rgPropVals));
					if (!lpAdrList->aEntries->rgPropVals) continue;

					// TODO: We are setting 5 properties below. If this changes, modify these two lines.
					ZeroMemory(lpAdrList->aEntries->rgPropVals, 5 * sizeof(SPropValue));
					lpAdrList->aEntries->cValues = 5;

					// Fill out addresslist with required property values.
					LPSPropValue pProps = lpAdrList->aEntries->rgPropVals;
					LPSPropValue pProp; // Just a pointer, do not free.

					pProp = &pProps[abPR_ENTRYID];
					pProp->ulPropTag = PR_ENTRYID;
					EC_H(CopySBinary(&pProp->Value.bin, &lpFoundRow[abPR_ENTRYID].Value.bin, lpAdrList));

					pProp = &pProps[abPR_RECIPIENT_TYPE];
					pProp->ulPropTag = PR_RECIPIENT_TYPE;
					pProp->Value.l = MAPI_TO;

					pProp = &pProps[abPR_DISPLAY_NAME];
					pProp->ulPropTag = PR_DISPLAY_NAME;

					if (!CheckStringProp(&lpFoundRow[abPR_DISPLAY_NAME], PT_TSTRING)) continue;

					EC_H(CopyString(
						&pProp->Value.LPSZ,
						lpFoundRow[abPR_DISPLAY_NAME].Value.LPSZ,
						lpAdrList));

					pProp = &pProps[abPR_ADDRTYPE];
					pProp->ulPropTag = PR_ADDRTYPE;

					if (!CheckStringProp(&lpFoundRow[abPR_ADDRTYPE], PT_TSTRING)) continue;

					EC_H(CopyString(
						&pProp->Value.LPSZ,
						lpFoundRow[abPR_ADDRTYPE].Value.LPSZ,
						lpAdrList));

					pProp = &pProps[abPR_DISPLAY_TYPE];
					pProp->ulPropTag = PR_DISPLAY_TYPE;
					pProp->Value.l = lpFoundRow[abPR_DISPLAY_TYPE].Value.l;

					EC_MAPI(lpMessage->ModifyRecipients(
						MODRECIP_ADD,
						lpAdrList));

					if (lpAdrList) FreePadrlist(lpAdrList);
					lpAdrList = NULL;

					EC_MAPI(lpMessage->SaveChanges(KEEP_OPEN_READWRITE));

					// since we're done with our work, let's get out of here.
					break;
				}
			}
		}
		lpABContainerTable->Release();
	}
	FreeProws(lpABRow);
	MAPIFreeBuffer(lpFoundRow);
	if (lpAdrList) FreePadrlist(lpAdrList);
	if (pTable) pTable->Release();
	if (lpABContainer) lpABContainer->Release();
	if (lpAdrBook) lpAdrBook->Release();
	return hRes;
} // ManualResolve

_Check_return_ HRESULT SearchContentsTableForName(
	_In_ LPMAPITABLE pTable,
	_In_z_ LPCTSTR szName,
	ULONG PropTagToCompare,
	_Deref_out_opt_ LPSPropValue *lppPropsFound)
{
	HRESULT			hRes = S_OK;

	LPSRowSet		pRows = NULL;

	enum
	{
		abPR_ENTRYID,
		abPR_DISPLAY_NAME,
		abPR_RECIPIENT_TYPE,
		abPR_ADDRTYPE,
		abPR_DISPLAY_TYPE,
		abPropTagToCompare,
		abNUM_COLS
	};

	const SizedSPropTagArray(abNUM_COLS, abCols) =
	{
		abNUM_COLS,
		PR_ENTRYID,
		PR_DISPLAY_NAME,
		PR_RECIPIENT_TYPE,
		PR_ADDRTYPE,
		PR_DISPLAY_TYPE,
		PropTagToCompare
	};

	*lppPropsFound = NULL;
	if (!pTable || !szName) return MAPI_E_INVALID_PARAMETER;

	DebugPrint(DBGGeneric, _T("SearchContentsTableForName: Looking for \"%s\"\n"), szName);

	// Set a restriction so we only find close matches
	LPSRestriction	lpSRes = NULL;

	EC_H(CreateANRRestriction(
		PR_ANR,
		szName,
		NULL,
		&lpSRes));

	EC_MAPI(pTable->SetColumns((LPSPropTagArray)&abCols, TBL_BATCH));

	// Jump to the top of the table...
	EC_MAPI(pTable->SeekRow(
		BOOKMARK_BEGINNING,
		0,
		NULL));

	// ..and jump to the first matching entry in the table
	EC_MAPI(pTable->Restrict(
		lpSRes,
		NULL
		));

	// Now we iterate through each of the matching entries
	if (!FAILED(hRes)) for (;;)
	{
		hRes = S_OK;
		if (pRows) FreeProws(pRows);
		pRows = NULL;
		EC_MAPI(pTable->QueryRows(
			1,
			NULL,
			&pRows));
		if (FAILED(hRes) || !pRows || (pRows && !pRows->cRows)) break;

		// An error at this point is an error with the current entry, so we can continue this for statement
		// Unless it's an allocation error. Those are bad.
		if (PropTagToCompare == pRows->aRow->lpProps[abPropTagToCompare].ulPropTag &&
			CheckStringProp(&pRows->aRow->lpProps[abPropTagToCompare], PT_TSTRING))
		{
			DebugPrint(DBGGeneric, _T("SearchContentsTableForName: found \"%s\"\n"), pRows->aRow->lpProps[abPropTagToCompare].Value.LPSZ);
			if (lstrcmpi(szName, pRows->aRow->lpProps[abPropTagToCompare].Value.LPSZ) == 0)
			{
				DebugPrint(DBGGeneric, _T("SearchContentsTableForName: This is an exact match!\n"));
				// We found a match! Return it!
				EC_MAPI(ScDupPropset(
					abNUM_COLS,
					pRows->aRow->lpProps,
					MAPIAllocateBuffer,
					lppPropsFound));
				break;
			}
		}
	}

	if (!*lppPropsFound)
	{
		DebugPrint(DBGGeneric, _T("SearchContentsTableForName: No exact match found.\n"));
	}
	MAPIFreeBuffer(lpSRes);
	if (pRows) FreeProws(pRows);
	return hRes;
} // SearchContentsTableForName

_Check_return_ HRESULT SelectUser(_In_ LPADRBOOK lpAdrBook, HWND hwnd, _Out_opt_ ULONG* lpulObjType, _Deref_out_opt_ LPMAILUSER* lppMailUser)
{
	if (!lpAdrBook || !hwnd || !lppMailUser) return MAPI_E_INVALID_PARAMETER;

	HRESULT hRes = S_OK;

	ADRPARM AdrParm = { 0 };
	LPADRLIST lpAdrList = NULL;
	LPSPropValue lpEntryID = NULL;
	LPMAILUSER lpMailUser = NULL;

	*lppMailUser = NULL;
	if (lpulObjType)
	{
		*lpulObjType = NULL;
	}

	CHAR szTitle[256];
	int iRet = NULL;
	EC_D(iRet, LoadStringA(GetModuleHandle(NULL),
		IDS_SELECTMAILBOX,
		szTitle,
		_countof(szTitle)));

	AdrParm.ulFlags = DIALOG_MODAL | ADDRESS_ONE | AB_SELECTONLY | AB_RESOLVE;
#pragma warning(push)
#pragma warning(disable:4616)
#pragma warning(disable:6276)
	AdrParm.lpszCaption = (LPTSTR)szTitle;
#pragma warning(pop)

	EC_H_CANCEL(lpAdrBook->Address(
		(ULONG_PTR*)&hwnd,
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
			ULONG ulObjType = NULL;

			EC_H(CallOpenEntry(
				NULL,
				lpAdrBook,
				NULL,
				NULL,
				lpEntryID->Value.bin.cb,
				(LPENTRYID)lpEntryID->Value.bin.lpb,
				NULL,
				MAPI_BEST_ACCESS,
				&ulObjType,
				(LPUNKNOWN*)&lpMailUser));
			if (SUCCEEDED(hRes) && lpMailUser)
			{
				*lppMailUser = lpMailUser;
			}
			else
			{
				if (lpMailUser)
				{
					lpMailUser->Release();
				}

				if (lpulObjType)
				{
					*lpulObjType = ulObjType;
				}
			}
		}
	}

	if (lpAdrList) FreePadrlist(lpAdrList);
	return hRes;
}