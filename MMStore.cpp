#include "stdafx.h"
#include "MrMAPI.h"
#include "MMStore.h"
#include "MAPIFunctions.h"
#include "MAPIStoreFunctions.h"
#include "ColumnTags.h"
#include "MMFolder.h"
#include "InterpretProp2.h"

HRESULT OpenStore(_In_ LPMAPISESSION lpMAPISession, ULONG ulIndex, _Out_ LPMDB* lppMDB)
{
	HRESULT hRes = S_OK;
	LPMAPITABLE lpStoreTable = NULL;
	LPMDB lpMDB = NULL;
	if (!lpMAPISession || !lppMDB) return MAPI_E_INVALID_PARAMETER;

	*lppMDB = NULL;

	WC_MAPI(lpMAPISession->GetMsgStoresTable(0, &lpStoreTable));

	if (lpStoreTable)
	{
		static const SizedSPropTagArray(1,sptStore) =
		{
			1,
			PR_ENTRYID,
		};
		WC_MAPI(lpStoreTable->SetColumns((LPSPropTagArray) &sptStore, TBL_ASYNC));

		if (SUCCEEDED(hRes))
		{
			LPSRowSet lpRow = NULL;
			WC_MAPI(lpStoreTable->SeekRow(BOOKMARK_BEGINNING, ulIndex, NULL));
			WC_MAPI(lpStoreTable->QueryRows(
				1,
				NULL,
				&lpRow));
			if (SUCCEEDED(hRes) && lpRow && 1 == lpRow->cRows && PR_ENTRYID == lpRow->aRow[0].lpProps[0].ulPropTag)
			{
				WC_H(CallOpenMsgStore(
					lpMAPISession,
					NULL,
					&lpRow->aRow[0].lpProps[0].Value.bin,
					MDB_NO_DIALOG | MDB_WRITE,
					&lpMDB));
				if (SUCCEEDED(hRes) && lpMDB)
				{
					*lppMDB = lpMDB;
				}
				else if (SUCCEEDED(hRes))
				{
					hRes = MAPI_E_CALL_FAILED;
				}
			}

			if (lpRow) FreeProws(lpRow);
		}
	}

	if (lpStoreTable) lpStoreTable->Release();

	return hRes;
} // OpenStore

HRESULT HrMAPIOpenStoreAndFolder(
	_In_ LPMAPISESSION lpMAPISession,
	_In_ ULONG ulFolder,
	_In_z_ LPCWSTR lpszFolderPath,
	_Out_opt_ LPMDB* lppMDB,
	_Deref_out_opt_ LPMAPIFOLDER* lppFolder)
{
	HRESULT hRes = S_OK;
	LPMDB lpMDB = NULL;
	LPMAPIFOLDER lpFolder = NULL;

	if (lpMAPISession)
	{
		// Check if we were told which store to open
		if (lpszFolderPath && lpszFolderPath[0] == L'#')
		{
			// Skip the '#'
			lpszFolderPath++;
			SBinary Bin = {0};
			LPSTR szPath = NULL;
#ifdef UNICODE
			szPath = lpszFolderPath;
#else
			EC_H(UnicodeToAnsi(lpszFolderPath, &szPath));
#endif
			// Find our slash if we have one and null terminate at it
			LPSTR szSlash = strchr(szPath + 1, '\\');
			if (szSlash)
			{
				szSlash[0] = L'\0';
			}

			// MyBinFromHex will balk at odd string length or forbidden characters
			// In order for cb to get bigger than 1, the string has to have at least 4 characters
			// Which is larger than any reasonable store number. So we use that to distinguish.
			if (MyBinFromHex((LPCTSTR) szPath, NULL, &Bin.cb) && Bin.cb > 1)
			{
				Bin.lpb = new BYTE[Bin.cb];
				if (Bin.lpb)
				{
					(void) MyBinFromHex((LPCTSTR) szPath, Bin.lpb, &Bin.cb);
					WC_H(CallOpenMsgStore(
						lpMAPISession,
						NULL,
						&Bin,
						MDB_NO_DIALOG | MDB_WRITE,
						&lpMDB));
					lpszFolderPath = wcschr(lpszFolderPath + 1, L'\\');
					delete[] Bin.lpb;
				}
			}
			else
			{
				hRes = S_OK;
				LPWSTR szEndPtr = NULL;
				ULONG ulStore = wcstoul(lpszFolderPath, &szEndPtr, 10);

				// Only '\' and NULL are acceptable next characters after our store number
				if (szEndPtr && (szEndPtr[0] == L'\\' || szEndPtr[0] == L'\0'))
				{
					// We have a store. Let's open it
					WC_H(OpenStore(lpMAPISession, ulStore, &lpMDB));
					lpszFolderPath = szEndPtr;
				}
				else
				{
					hRes = MAPI_E_INVALID_PARAMETER;
				}
			}

#ifdef UNICODE
			delete[] szPath;
#endif
		}
		else
		{
			WC_H(OpenExchangeOrDefaultMessageStore(lpMAPISession, &lpMDB));
		}
	}

	if (SUCCEEDED(hRes) && lpMDB)
	{
		if (lpszFolderPath)
		{
			WC_H(HrMAPIOpenFolderExW(lpMDB, lpszFolderPath, &lpFolder));
		}
		else
		{
			WC_H(OpenDefaultFolder(ulFolder, lpMDB, &lpFolder));
		}
	}

	if (SUCCEEDED(hRes))
	{
		if (lpFolder)
		{
			if (lppFolder)
			{
				*lppFolder = lpFolder;
			}
			else
			{
				lpFolder->Release();
			}
		}
		if (lpMDB)
		{
			if (lppMDB)
			{
				*lppMDB = lpMDB;
			}
			else
			{
				lpMDB->Release();
			}
		}
	}
	return hRes;
} // HrMAPIOpenStoreAndFolder

void PrintObjectProperties(_In_z_ LPCTSTR szObjType, _In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag)
{
	HRESULT hRes = S_OK;
	if (!lpMAPIProp) return;

	printf(g_szXMLHeader);
	printf(_T("<%s>\n"), szObjType);

	LPSPropValue lpAllProps = NULL;
	ULONG cValues = 0L;

	if (ulPropTag)
	{
		SPropTagArray sTag = {0};
		sTag.cValues = 1;
		sTag.aulPropTag[0] = ulPropTag;

		WC_H_GETPROPS(lpMAPIProp->GetProps(&sTag,
			fMapiUnicode,
			&cValues,
			&lpAllProps)); 
	}
	else
	{
		WC_H_GETPROPS(GetPropsNULL(lpMAPIProp,
			fMapiUnicode,
			&cValues,
			&lpAllProps));
	}
	if (FAILED(hRes))
	{
		printf(_T("<properties error=\"0x%08X\" />\n"),hRes);
	}
	else if (lpAllProps)
	{
		printf(_T("<properties>\n"));

		_OutputProperties(DBGNoDebug, stdout, cValues, lpAllProps, lpMAPIProp, true);

		printf(_T("</properties>\n"));

		MAPIFreeBuffer(lpAllProps);
	}

	printf(_T("</%s>\n"), szObjType);
} // PrintObjectProperties

void PrintStoreProperties(_In_ LPMDB lpMDB, ULONG ulPropTag)
{
	PrintObjectProperties(_T("messagestoreprops"), lpMDB, ulPropTag);
} // PrintStoreProperties

void PrintStoreTable(_In_ LPMAPISESSION lpMAPISession, ULONG ulPropTag)
{
	HRESULT hRes = S_OK;
	LPMAPITABLE lpStoreTable = NULL;
	if (!lpMAPISession) return;

	printf(g_szXMLHeader);
	printf(_T("<storetable>\n"));
	WC_MAPI(lpMAPISession->GetMsgStoresTable(0, &lpStoreTable));

	if (lpStoreTable)
	{
		LPSPropTagArray sTags = (LPSPropTagArray) &sptSTORECols;
		SPropTagArray sTag = {0};
		if (ulPropTag)
		{
			sTag.cValues = 1;
			sTag.aulPropTag[0] = ulPropTag;
			sTags = &sTag;
		}

		WC_MAPI(lpStoreTable->SetColumns(sTags, TBL_ASYNC));

		if (SUCCEEDED(hRes))
		{
			LPSRowSet lpRows = NULL;
			ULONG iCurStore = 0;
			if (!FAILED(hRes)) for (;;)
			{
				hRes = S_OK;
				if (lpRows) FreeProws(lpRows);
				lpRows = NULL;
				WC_MAPI(lpStoreTable->QueryRows(
					10,
					NULL,
					&lpRows));
				if (FAILED(hRes) || !lpRows || !lpRows->cRows) break;

				ULONG i = 0;
				for (i = 0; i < lpRows->cRows; i++)
				{
					printf(_T("<properties index=\"%d\">\n"), iCurStore);
					_OutputProperties(DBGNoDebug, stdout, lpRows->aRow[i].cValues, lpRows->aRow[i].lpProps, NULL, false);
					printf(_T("</properties>\n"));
					iCurStore++;
				}
			}
			if (lpRows) FreeProws(lpRows);
		}
	}

	printf(_T("</storetable>\n"));
	if (lpStoreTable) lpStoreTable->Release();
} // PrintStoreTable

void DoStore(_In_ MYOPTIONS ProgOpts)
{
	HRESULT hRes = S_OK;
	LPMAPISESSION lpMAPISession = NULL;
	LPMDB lpMDB = NULL;
	ULONG ulPropTag = NULL;
	// If we have a prop tag, parse it
	// For now, we don't support dispids
	if (ProgOpts.lpszUnswitchedOption && !(ProgOpts.ulOptions & OPT_DODISPID))
	{
		WC_H(PropNameToPropTagW(ProgOpts.lpszUnswitchedOption, &ulPropTag));
	}
	hRes = S_OK;

	if (ProgOpts.lpMAPISession)
	{
		if (!ProgOpts.lpszStore)
		{
			PrintStoreTable(ProgOpts.lpMAPISession, ulPropTag);
		}
		else 
		{
			LPWSTR szEndPtr = NULL;
			ULONG ulStore = wcstoul(ProgOpts.lpszStore, &szEndPtr, 10);
			WC_H(OpenStore(ProgOpts.lpMAPISession, ulStore, &lpMDB));
		}
	}

	if (lpMDB)
	{
		PrintStoreProperties(lpMDB, ulPropTag);
	}

	if (lpMDB) lpMDB->Release();
} // DoStore