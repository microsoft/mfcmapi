// QuickStart.cpp : implementation file
//

#include "stdafx.h"
#include "MAPIStoreFunctions.h"
#include "MAPIFunctions.h"
#include "MapiObjects.h"
#include "MFCUtilityFunctions.h"
#include "AbContDlg.h"
#include "ExtraPropTags.h"
#include "InterpretProp2.h"
#include "SmartView.h"
#include "Editor.h"
#include "MainDlg.h"

LPMAPISESSION OpenSessionForQuickStart(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd)
{
	CMapiObjects* lpMapiObjects = lpHostDlg->GetMapiObjects(); // do not release
	if (!lpMapiObjects) return NULL;

	LPMAPISESSION lpMAPISession = lpMapiObjects->LogonGetSession(hwnd); // do not release
	if (lpMAPISession)
	{
		// Since we've opened a session, populate the store table in the UI
		lpHostDlg->OnOpenMessageStoreTable();
		return lpMAPISession;
	}

	return NULL;
} // OpenSessionForQuickStart

HRESULT OpenStoreForQuickStart(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd, _Out_ LPMDB* lppMDB)
{
	HRESULT hRes = S_OK;
	if (!lppMDB) return MAPI_E_INVALID_PARAMETER;
	*lppMDB = NULL;

	CMapiObjects* lpMapiObjects = lpHostDlg->GetMapiObjects(); // do not release
	if (!lpMapiObjects) return MAPI_E_CALL_FAILED;

	LPMAPISESSION lpMAPISession = OpenSessionForQuickStart(lpHostDlg, hwnd); // do not release
	if (lpMAPISession)
	{
		LPMDB lpMDB = NULL;
		WC_H(OpenDefaultMessageStore(lpMAPISession, &lpMDB));
		if (SUCCEEDED(hRes))
		{
			*lppMDB = lpMDB;
			lpMapiObjects->SetMDB(lpMDB);
		}
	}

	return hRes;
} // OpenStoreForQuickStart

HRESULT OpenABForQuickStart(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd, _Out_ LPADRBOOK* lppAdrBook)
{
	HRESULT hRes = S_OK;
	if (!lppAdrBook) return MAPI_E_INVALID_PARAMETER;
	*lppAdrBook = NULL;

	CMapiObjects* lpMapiObjects = lpHostDlg->GetMapiObjects(); // do not release
	if (!lpMapiObjects) return MAPI_E_CALL_FAILED;

	// ensure we have an AB
	(void) OpenSessionForQuickStart(lpHostDlg, hwnd); // do not release
	LPADRBOOK lpAdrBook =  lpMapiObjects->GetAddrBook(true); // do not release

	if (lpAdrBook)
	{
		lpAdrBook->AddRef();
		*lppAdrBook = lpAdrBook;
	}

	return hRes;
} // OpenABForQuickStart

void OnQSDisplayFolder(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd, _In_ ULONG ulFolder)
{
	HRESULT hRes = S_OK;

	LPMDB lpMDB = NULL;
	WC_H(OpenStoreForQuickStart(lpHostDlg, hwnd, &lpMDB));

	if (lpMDB)
	{
		LPMAPIFOLDER lpFolder = NULL;
		WC_H(OpenDefaultFolder(ulFolder, lpMDB, &lpFolder));

		if (lpFolder)
		{
			WC_H(DisplayObject(
				lpFolder,
				NULL,
				otContents,
				lpHostDlg));

			lpFolder->Release();
		}

		lpMDB->Release();
	}
} // OnQSDisplayFolder

void OnQSDisplayTable(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd, _In_ ULONG ulFolder, _In_ ULONG ulProp, _In_ ObjectType tType)
{
	HRESULT hRes = S_OK;

	LPMDB lpMDB = NULL;
	WC_H(OpenStoreForQuickStart(lpHostDlg, hwnd, &lpMDB));

	if (lpMDB)
	{
		LPMAPIFOLDER lpFolder = NULL;

		WC_H(OpenDefaultFolder(ulFolder, lpMDB, &lpFolder));

		if (lpFolder)
		{
			WC_H(DisplayExchangeTable(
				lpFolder,
				ulProp,
				tType,
				lpHostDlg));
			lpFolder->Release();
		}

		lpMDB->Release();
	}
} // OnQSDisplayTable

void OnQSDisplayDefaultDir(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd)
{
	HRESULT hRes = S_OK;

	LPADRBOOK lpAdrBook = NULL;
	WC_H(OpenABForQuickStart(lpHostDlg, hwnd, &lpAdrBook));
	if (SUCCEEDED(hRes) && lpAdrBook)
	{
		ULONG cbEID = NULL;
		LPENTRYID lpEID = NULL;
		ULONG ulObjType = NULL;
		LPABCONT lpDefaultDir = NULL;

		WC_MAPI(lpAdrBook->GetDefaultDir(
			&cbEID,
			&lpEID));

		WC_H(CallOpenEntry(
			NULL,
			lpAdrBook,
			NULL,
			NULL,
			cbEID,
			lpEID,
			NULL,
			MAPI_MODIFY,
			&ulObjType,
			(LPUNKNOWN*)&lpDefaultDir));

		if (lpDefaultDir)
		{
			WC_H(DisplayObject(lpDefaultDir, ulObjType, otDefault, lpHostDlg));

			lpDefaultDir->Release();
		}
		MAPIFreeBuffer(lpEID);
	}
	if (lpAdrBook) lpAdrBook->Release();
} // OnQSDisplayDefaultDir

void OnQSDisplayAB(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd)
{
	HRESULT hRes = S_OK;

	CMapiObjects* lpMapiObjects = lpHostDlg->GetMapiObjects(); // do not release
	if (!lpMapiObjects) return;

	CParentWnd* lpParentWnd = lpHostDlg->GetParentWnd(); // do not release
	if (!lpParentWnd) return;

	LPADRBOOK lpAdrBook = NULL;
	WC_H(OpenABForQuickStart(lpHostDlg, hwnd, &lpAdrBook));
	if (SUCCEEDED(hRes) && lpAdrBook)
	{
		// call the dialog
		new CAbContDlg(
			lpParentWnd,
			lpMapiObjects);
	}
	if (lpAdrBook) lpAdrBook->Release();
} // OnQSDisplayAB

void OnQSDisplayNicknameCache(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd)
{
	HRESULT hRes = S_OK;
	LPWSTR szNicknames = NULL;
	LPSPropValue lpsProp = NULL;

	lpHostDlg->UpdateStatusBarText(STATUSINFOTEXT, IDS_STATUSTEXTLOADINGNICKNAME, NULL, NULL, NULL);
	lpHostDlg->SendMessage(WM_PAINT, NULL, NULL); // force paint so we update the status now

	LPMDB lpMDB = NULL;
	WC_H(OpenStoreForQuickStart(lpHostDlg, hwnd, &lpMDB));

	if (lpMDB)
	{
		LPMAPIFOLDER lpFolder = NULL;
		WC_H(OpenDefaultFolder(DEFAULT_INBOX, lpMDB, &lpFolder));

		if (lpFolder)
		{
			LPMAPITABLE lpTable = NULL;
			WC_MAPI(lpFolder->GetContentsTable(MAPI_ASSOCIATED, &lpTable));

			if (lpTable)
			{
				SRestriction sRes = {0};
				SPropValue sPV = {0};
				sRes.rt = RES_PROPERTY;
				sRes.res.resProperty.ulPropTag = PR_MESSAGE_CLASS_A;
				sRes.res.resProperty.relop = RELOP_EQ;
				sRes.res.resProperty.lpProp = &sPV;
				sPV.ulPropTag = sRes.res.resProperty.ulPropTag;
				sPV.Value.LPSZ = _T("IPM.Configuration.Autocomplete"); // STRING_OK
				WC_MAPI(lpTable->Restrict(&sRes, TBL_BATCH));

				if (SUCCEEDED(hRes))
				{
					enum
					{
						eidPR_ENTRYID,
						eidNUM_COLS
					};
					static const SizedSPropTagArray(eidNUM_COLS, eidCols) =
					{
						eidNUM_COLS,
						PR_ENTRYID,
					};
					WC_MAPI(lpTable->SetColumns((LPSPropTagArray) &eidCols, TBL_BATCH));

					if (SUCCEEDED(hRes))
					{
						LPSRowSet lpRows = NULL;
						WC_MAPI(lpTable->QueryRows(1, NULL, &lpRows));

						if (lpRows && 1 == lpRows->cRows && PR_ENTRYID == lpRows->aRow[0].lpProps[eidPR_ENTRYID].ulPropTag)
						{
							LPMESSAGE lpMSG = NULL;
							WC_H(CallOpenEntry(lpMDB, NULL, NULL, NULL, &lpRows->aRow[0].lpProps[eidPR_ENTRYID].Value.bin, NULL, NULL, NULL, (LPUNKNOWN*) &lpMSG));

							if (SUCCEEDED(hRes) && lpMSG)
							{
								WC_H(GetLargeBinaryProp(lpMSG, PR_ROAMING_BINARYSTREAM, &lpsProp));

								if (lpsProp)
								{
									// Get the string interpretation
									InterpretBinaryAsString(lpsProp->Value.bin, IDS_STNICKNAMECACHE, lpMSG, PR_ROAMING_BINARYSTREAM, &szNicknames);
								}
								lpMSG->Release();
							}
						}

						FreeProws(lpRows);
					}
				}
				lpTable->Release();
			}
			lpFolder->Release();
		}
		lpMDB->Release();

		// Display our dialog
		if (szNicknames || lpsProp)
		{
			CEditor MyResults(
				lpHostDlg,
				IDS_NICKNAME,
				NULL,
				3,
				CEDITOR_BUTTON_OK);
			MyResults.InitMultiLine(0, NULL, NULL, true);
			MyResults.InitSingleLine(1, IDS_CB, NULL, true);
			MyResults.InitMultiLine(2, IDS_HEX, NULL, true);

			if (szNicknames)
			{
				MyResults.SetStringW(0, szNicknames);
			}

			if (lpsProp)
			{
				MyResults.SetSize(1, lpsProp->Value.bin.cb);
				MyResults.SetBinary(2, lpsProp->Value.bin.lpb, lpsProp->Value.bin.cb);
			}

			WC_H(MyResults.DisplayDialog());

			delete[] szNicknames;
		}

		MAPIFreeBuffer(lpsProp);
	}
	lpHostDlg->UpdateStatusBarText(STATUSINFOTEXT, _T(""));
} // OnQSDisplayNicknameCache

enum
{
	qPR_STORE_SUPPORT_MASK,
	qPR_DISPLAY_NAME_W,
	qPR_MESSAGE_SIZE_EXTENDED,
	qPR_STORAGE_QUOTA_LIMIT,
	qPR_PROHIBIT_SEND_QUOTA,
	qPR_PROHIBIT_RECEIVE_QUOTA,
	qPR_MAX_SUBMIT_MESSAGE_SIZE,
	qPR_QUOTA_WARNING,
	qPR_QUOTA_SEND,
	qPR_QUOTA_RECEIVE,
	qPR_MDB_PROVIDER,
	qNUM_COLS
};
static const SizedSPropTagArray(qNUM_COLS,sptaQuota) =
{
	qNUM_COLS,
	PR_STORE_SUPPORT_MASK,
	PR_DISPLAY_NAME_W,
	PR_MESSAGE_SIZE_EXTENDED,
	PR_STORAGE_QUOTA_LIMIT,
	PR_PROHIBIT_SEND_QUOTA,
	PR_PROHIBIT_RECEIVE_QUOTA,
	PR_MAX_SUBMIT_MESSAGE_SIZE,
	PR_QUOTA_WARNING,
	PR_QUOTA_SEND,
	PR_QUOTA_RECEIVE,
	PR_MDB_PROVIDER,
};

CString FormatQuota(LPSPropValue lpProp, ULONG ulPropTag, LPCTSTR szName)
{
	if (lpProp && lpProp->ulPropTag == ulPropTag)
	{
		CString szTmp;
		szTmp.FormatMessage(IDS_QUOTAPROP, szName, lpProp->Value.l);
		return szTmp;
	}
	return _T("");
} // FormatQuota

#define AddFormattedQuota(__TAG)   szQuotaString += FormatQuota(&lpProps[q##__TAG], __TAG, _T(#__TAG));

void OnQSDisplayQuota(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd)
{
	HRESULT hRes = S_OK;
	CString szQuotaString;

	lpHostDlg->UpdateStatusBarText(STATUSINFOTEXT, IDS_STATUSTEXTLOADINGQUOTA, NULL, NULL, NULL);
	lpHostDlg->SendMessage(WM_PAINT, NULL, NULL); // force paint so we update the status now

	LPMDB lpMDB = NULL;
	WC_H(OpenStoreForQuickStart(lpHostDlg, hwnd, &lpMDB));

	if (SUCCEEDED(hRes) && lpMDB)
	{
		ULONG cProps = 0;
		LPSPropValue lpProps = NULL;

		// Get quota properties
		WC_H_GETPROPS(lpMDB->GetProps(
			(LPSPropTagArray) &sptaQuota,
			fMapiUnicode,
			&cProps,
			&lpProps));

		if (lpProps)
		{
			CString szTmp;

			if (lpProps[qPR_DISPLAY_NAME_W].ulPropTag == PR_DISPLAY_NAME_W)
			{
				WCHAR szNotFound[16];
				LPWSTR szDisplayName = lpProps[qPR_DISPLAY_NAME_W].Value.lpszW;
				if (!szDisplayName[0])
				{
					int iRet = NULL;
					WC_D(iRet,LoadStringW(GetModuleHandle(NULL),
						IDS_NOTFOUND,
						szNotFound,
						_countof(szNotFound)));
					szDisplayName = szNotFound;
				}

				szTmp.FormatMessage(IDS_QUOTADISPLAYNAME, szDisplayName);
				szQuotaString += szTmp;
			}

			if (lpProps[qPR_MESSAGE_SIZE_EXTENDED].ulPropTag == PR_MESSAGE_SIZE_EXTENDED)
			{
				szTmp.FormatMessage(IDS_QUOTASIZE,
					lpProps[qPR_MESSAGE_SIZE_EXTENDED].Value.li.QuadPart,
					lpProps[qPR_MESSAGE_SIZE_EXTENDED].Value.li.QuadPart / 1024);
				szQuotaString += szTmp;
			}

			// All of these properties are in kilobytes. Be careful adding a property not in kilobytes.
			AddFormattedQuota(PR_PROHIBIT_SEND_QUOTA);
			AddFormattedQuota(PR_PROHIBIT_RECEIVE_QUOTA);
			AddFormattedQuota(PR_STORAGE_QUOTA_LIMIT);
			AddFormattedQuota(PR_QUOTA_SEND);
			AddFormattedQuota(PR_QUOTA_RECEIVE);
			AddFormattedQuota(PR_QUOTA_WARNING);
			AddFormattedQuota(PR_MAX_SUBMIT_MESSAGE_SIZE);

			if (lpProps[qPR_STORE_SUPPORT_MASK].ulPropTag == PR_STORE_SUPPORT_MASK)
			{
				LPWSTR szFlags = NULL;
				InterpretNumberAsStringProp(lpProps[qPR_STORE_SUPPORT_MASK].Value.l, PR_STORE_SUPPORT_MASK, &szFlags);
				szTmp.FormatMessage(IDS_QUOTAMASK, lpProps[qPR_STORE_SUPPORT_MASK].Value.l, szFlags);
				szQuotaString += szTmp;
				delete[] szFlags;
				szFlags = NULL;
			}

			if (lpProps[qPR_MDB_PROVIDER].ulPropTag == PR_MDB_PROVIDER)
			{
				szTmp.FormatMessage(IDS_QUOTAPROVIDER, (LPCTSTR) BinToHexString(&lpProps[qPR_MDB_PROVIDER].Value.bin, true));
				szQuotaString += szTmp;
			}

			MAPIFreeBuffer(lpProps);
		}
		lpMDB->Release();

		// Display our dialog
		CEditor MyResults(
			lpHostDlg,
			IDS_QUOTA,
			NULL,
			1,
			CEDITOR_BUTTON_OK);
		MyResults.InitMultiLine(0, NULL, NULL, true);
		MyResults.SetString(0, szQuotaString);

		WC_H(MyResults.DisplayDialog());
	}

	lpHostDlg->UpdateStatusBarText(STATUSINFOTEXT, _T(""));
} // OnQSDisplayQuota

bool HandleQuickStart(_In_ WORD wMenuSelect, _In_ CMainDlg* lpHostDlg, _In_ HWND hwnd)
{
	switch (wMenuSelect)
	{
	case ID_QSINBOX: OnQSDisplayFolder(lpHostDlg, hwnd, DEFAULT_INBOX); return true;
	case ID_QSCALENDAR: OnQSDisplayFolder(lpHostDlg, hwnd, DEFAULT_CALENDAR); return true;
	case ID_QSCONTACTS: OnQSDisplayFolder(lpHostDlg, hwnd, DEFAULT_CONTACTS); return true;
	case ID_QSRULES: OnQSDisplayTable(lpHostDlg, hwnd, DEFAULT_INBOX, PR_RULES_TABLE, otRules); return true;
	case ID_QSDEFAULTDIR: OnQSDisplayDefaultDir(lpHostDlg, hwnd); return true;
	case ID_QSAB: OnQSDisplayAB(lpHostDlg, hwnd); return true;
	case ID_QSCALPERM: OnQSDisplayTable(lpHostDlg, hwnd, DEFAULT_CALENDAR, PR_ACL_TABLE, otACL); return true;
	case ID_QSNICKNAME: OnQSDisplayNicknameCache(lpHostDlg, hwnd); return true;
	case ID_QSQUOTA: OnQSDisplayQuota(lpHostDlg, hwnd); return true;
	}
	return false;
}