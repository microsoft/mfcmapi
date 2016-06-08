#include "stdafx.h"
#include "..\stdafx.h"
#include "MrMAPI.h"
#include "MMStore.h"
#include "MMFolder.h"
#include "..\ColumnTags.h"
#include "..\InterpretProp2.h"
#include "..\MAPIFunctions.h"
#include "..\MAPIStoreFunctions.h"
#include "..\String.h"

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
		static const SizedSPropTagArray(1, sptStore) =
		{
			1,
			PR_ENTRYID,
		};
		WC_MAPI(lpStoreTable->SetColumns((LPSPropTagArray)&sptStore, TBL_ASYNC));

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
}

HRESULT HrMAPIOpenStoreAndFolder(
	_In_ LPMAPISESSION lpMAPISession,
	_In_ ULONG ulFolder,
	_In_ wstring lpszFolderPath,
	_Out_opt_ LPMDB* lppMDB,
	_Deref_out_opt_ LPMAPIFOLDER* lppFolder)
{
	HRESULT hRes = S_OK;
	LPMDB lpMDB = NULL;
	LPMAPIFOLDER lpFolder = NULL;

	if (lpMAPISession)
	{
		// Check if we were told which store to open
		if (!lpszFolderPath.empty() && lpszFolderPath[0] == L'#')
		{
			// Skip the '#'
			lpszFolderPath.erase(0, 1);
			SBinary Bin = { 0 };
			LPSTR szPath = wstringToLPTSTR(lpszFolderPath);

			// Find our slash if we have one and null terminate at it
			LPSTR szSlash = strchr(szPath + 1, '\\');
			if (szSlash)
			{
				szSlash[0] = L'\0';
			}

			// MyBinFromHex will balk at odd string length or forbidden characters
			// In order for cb to get bigger than 1, the string has to have at least 4 characters
			// Which is larger than any reasonable store number. So we use that to distinguish.
			if (MyBinFromHex((LPCTSTR)szPath, NULL, &Bin.cb) && Bin.cb > 1)
			{
				Bin.lpb = new BYTE[Bin.cb];
				if (Bin.lpb)
				{
					(void)MyBinFromHex((LPCTSTR)szPath, Bin.lpb, &Bin.cb);
					WC_H(CallOpenMsgStore(
						lpMAPISession,
						NULL,
						&Bin,
						MDB_NO_DIALOG | MDB_WRITE,
						&lpMDB));
					size_t slashPos = lpszFolderPath.find_first_of(L'\\');
					if (slashPos != string::npos)
					{
						lpszFolderPath = lpszFolderPath.substr(slashPos, string::npos);
					}
					else
					{
						lpszFolderPath = L"";
					}

					delete[] Bin.lpb;
				}
			}
			else
			{
				hRes = S_OK;
				LPWSTR szEndPtr = NULL;
				ULONG ulStore = wcstoul(lpszFolderPath.c_str(), &szEndPtr, 10);

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

			delete[] szPath;
		}
		else
		{
			WC_H(OpenExchangeOrDefaultMessageStore(lpMAPISession, &lpMDB));
		}
	}

	if (SUCCEEDED(hRes) && lpMDB)
	{
		if (!lpszFolderPath.empty())
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
	else
	{
		if (lpFolder) lpFolder->Release();
		if (lpMDB) lpMDB->Release();
	}

	return hRes;
}

void PrintObjectProperty(_In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag)
{
	HRESULT hRes = S_OK;
	if (!lpMAPIProp || !ulPropTag) return;

	LPSPropValue lpAllProps = NULL;
	ULONG cValues = 0L;

	SPropTagArray sTag = { 0 };
	sTag.cValues = 1;
	sTag.aulPropTag[0] = ulPropTag;

	WC_H_GETPROPS(lpMAPIProp->GetProps(&sTag,
		fMapiUnicode,
		&cValues,
		&lpAllProps));

	_OutputProperties(DBGNoDebug, stdout, cValues, lpAllProps, lpMAPIProp, true);

	MAPIFreeBuffer(lpAllProps);
}

void PrintObjectProperties(wstring const& szObjType, _In_ LPMAPIPROP lpMAPIProp, ULONG ulPropTag)
{
	HRESULT hRes = S_OK;
	if (!lpMAPIProp) return;

	wprintf(g_szXMLHeader.c_str());
	wprintf(L"<%ws>\n", szObjType.c_str());

	LPSPropValue lpAllProps = NULL;
	ULONG cValues = 0L;

	if (ulPropTag)
	{
		SPropTagArray sTag = { 0 };
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
		wprintf(L"<properties error=\"0x%08X\" />\n", hRes);
	}
	else if (lpAllProps)
	{
		wprintf(L"<properties>\n");

		_OutputProperties(DBGNoDebug, stdout, cValues, lpAllProps, lpMAPIProp, true);

		wprintf(L"</properties>\n");

		MAPIFreeBuffer(lpAllProps);
	}

	wprintf(L"</%ws>\n", szObjType.c_str());
}

void PrintStoreProperty(_In_ LPMAPISESSION lpMAPISession, ULONG ulIndex, ULONG ulPropTag)
{
	if (!lpMAPISession || !ulPropTag) return;

	HRESULT hRes = S_OK;
	LPMDB lpMDB = NULL;
	WC_H(OpenStore(lpMAPISession, ulIndex, &lpMDB));
	if (lpMDB)
	{
		PrintObjectProperty(lpMDB, ulPropTag);

		lpMDB->Release();
	}
}

void PrintStoreProperties(_In_ LPMDB lpMDB, ULONG ulPropTag)
{
	PrintObjectProperties(L"messagestoreprops", lpMDB, ulPropTag);
}

void PrintStoreTable(_In_ LPMAPISESSION lpMAPISession, ULONG ulPropTag)
{
	HRESULT hRes = S_OK;
	LPMAPITABLE lpStoreTable = NULL;
	if (!lpMAPISession) return;

	wprintf(g_szXMLHeader.c_str());
	wprintf(L"<storetable>\n");
	WC_MAPI(lpMAPISession->GetMsgStoresTable(0, &lpStoreTable));

	if (lpStoreTable)
	{
		LPSPropTagArray sTags = (LPSPropTagArray)&sptSTORECols;
		SPropTagArray sTag = { 0 };
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
					hRes = S_OK;
					wprintf(L"<properties index=\"%u\">\n", iCurStore);
					if (ulPropTag &&
						lpRows->aRow[0].lpProps &&
						PT_ERROR == PROP_TYPE(lpRows->aRow[0].lpProps->ulPropTag) &&
						MAPI_E_NOT_FOUND == lpRows->aRow[0].lpProps->Value.err)
					{
						PrintStoreProperty(lpMAPISession, i, ulPropTag);
					}
					else
					{
						_OutputProperties(DBGNoDebug, stdout, lpRows->aRow[i].cValues, lpRows->aRow[i].lpProps, NULL, false);
					}

					wprintf(L"</properties>\n");
					iCurStore++;
				}
			}

			if (lpRows) FreeProws(lpRows);
		}
	}

	wprintf(L"</storetable>\n");
	if (lpStoreTable) lpStoreTable->Release();
}

void DoStore(_In_ MYOPTIONS ProgOpts)
{
	HRESULT hRes = S_OK;
	ULONG ulPropTag = NULL;

	// If we have a prop tag, parse it
	// For now, we don't support dispids
	if (!ProgOpts.lpszUnswitchedOption.empty() && !(ProgOpts.ulOptions & OPT_DODISPID))
	{
		ulPropTag = PropNameToPropTag(ProgOpts.lpszUnswitchedOption);
	}

	LPMDB lpMDB = NULL;
	if (ProgOpts.lpMAPISession)
	{
		if (0 == ProgOpts.ulStore)
		{
			PrintStoreTable(ProgOpts.lpMAPISession, ulPropTag);
		}
		else
		{
			// ulStore was incremented by 1 before, so drop it back now
			WC_H(OpenStore(ProgOpts.lpMAPISession, ProgOpts.ulStore - 1, &lpMDB));
		}
	}

	if (lpMDB)
	{
		PrintStoreProperties(lpMDB, ulPropTag);
		lpMDB->Release();
	}
}