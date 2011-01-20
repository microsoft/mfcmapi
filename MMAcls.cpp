#include "stdafx.h"
#include "MrMAPI.h"
#include "MMAcls.h"
#include "MAPIFunctions.h"
#include "MAPIStoreFunctions.h"
#include "ExtraPropTags.h"

STDMETHODIMP OpenPropFromMDB(LPMDB lpMDB, ULONG ulPropTag, LPMAPIFOLDER *lpFolder)
{
	HRESULT			hRes = S_OK;
	ULONG			ulObjType = 0;
	LPSPropValue	lpEIDProp = NULL;

	*lpFolder = NULL;

	WC_H(HrGetOneProp(lpMDB, ulPropTag, &lpEIDProp));

	// Open whatever folder we got..
	if (SUCCEEDED(hRes) && lpEIDProp)
	{
		LPMAPIFOLDER	lpTemp = NULL;

		WC_H(CallOpenEntry(
			lpMDB,
			NULL,
			NULL,
			NULL,
			lpEIDProp->Value.bin.cb,
			(LPENTRYID) lpEIDProp->Value.bin.lpb,
			NULL,
			MAPI_BEST_ACCESS,
			NULL,
			(LPUNKNOWN*)&lpTemp));
		if (SUCCEEDED(hRes) && lpTemp)
		{
			*lpFolder = lpTemp;
		}
	}

	MAPIFreeBuffer(lpEIDProp);
	return hRes;
}

STDMETHODIMP OpenDefaultFolder(ULONG ulFolder, LPMDB lpMDB, LPMAPIFOLDER *lpFolder)
{
	HRESULT			hRes = S_OK;

	if (!lpMDB || !lpFolder)
	{
		return MAPI_E_INVALID_PARAMETER;
	}

	*lpFolder = NULL;

	switch(ulFolder)
	{
	case DEFAULT_CALENDAR:
		hRes = GetSpecialFolder(lpMDB,PR_IPM_APPOINTMENT_ENTRYID,lpFolder);
		break;
	case DEFAULT_CONTACTS:
		hRes = GetSpecialFolder(lpMDB,PR_IPM_CONTACT_ENTRYID,lpFolder);
		break;
	case DEFAULT_JOURNAL:
		hRes = GetSpecialFolder(lpMDB,PR_IPM_JOURNAL_ENTRYID,lpFolder);
		break;
	case DEFAULT_NOTES:
		hRes = GetSpecialFolder(lpMDB,PR_IPM_NOTE_ENTRYID,lpFolder);
		break;
	case DEFAULT_TASKS:
		hRes = GetSpecialFolder(lpMDB,PR_IPM_TASK_ENTRYID,lpFolder);
		break;
	case DEFAULT_REMINDERS:
		hRes = GetSpecialFolder(lpMDB,PR_REM_ONLINE_ENTRYID,lpFolder);
		break;
	case DEFAULT_DRAFTS:
		hRes = GetSpecialFolder(lpMDB,PR_IPM_DRAFTS_ENTRYID,lpFolder);
		break;
	case DEFAULT_SENTITEMS:
		hRes = OpenPropFromMDB(lpMDB,PR_IPM_SENTMAIL_ENTRYID,lpFolder);
		break;
	case DEFAULT_OUTBOX:
		hRes = OpenPropFromMDB(lpMDB,PR_IPM_OUTBOX_ENTRYID,lpFolder);
		break;
	case DEFAULT_DELETEDITEMS:
		hRes = OpenPropFromMDB(lpMDB,PR_IPM_WASTEBASKET_ENTRYID,lpFolder);
		break;
	case DEFAULT_FINDER:
		hRes = OpenPropFromMDB(lpMDB,PR_FINDER_ENTRYID,lpFolder);
		break;
	case DEFAULT_IPM_SUBTREE:
		hRes = OpenPropFromMDB(lpMDB,PR_IPM_SUBTREE_ENTRYID,lpFolder);
		break;
	case DEFAULT_INBOX:
		hRes = GetInbox(lpMDB, lpFolder);
		break;
	default:
		hRes = MAPI_E_INVALID_PARAMETER;
	}

	return hRes;
} // OpenDefaultFolder

void DumpExchangeTable(_In_z_ LPWSTR lpszProfile, _In_ ULONG ulPropTag, _In_ ULONG ulFolder)
{
	InitMFC();
	HRESULT hRes = S_OK;
	LPMAPISESSION lpMAPISession = NULL;
	LPMDB lpMDB = NULL;
	LPMAPIFOLDER lpFolder = NULL;
	LPEXCHANGEMODIFYTABLE lpExchTbl = NULL;
	LPMAPITABLE lpTbl = NULL;

	WC_H(MAPIInitialize(NULL));

	ULONG ulFlags = MAPI_EXTENDED | MAPI_NO_MAIL | MAPI_UNICODE | MAPI_NEW_SESSION;
	if (!lpszProfile) ulFlags |= MAPI_USE_DEFAULT;

	WC_H(MAPILogonEx(NULL, (LPTSTR) lpszProfile, NULL,
		ulFlags,
		&lpMAPISession));

	if (lpMAPISession)
	{
		WC_H(OpenMessageStoreGUID(lpMAPISession, pbExchangeProviderPrimaryUserGuid, &lpMDB));
	}
	if (lpMDB)
	{
		WC_H(OpenDefaultFolder(ulFolder,lpMDB,&lpFolder));
	}
	if (lpFolder)
	{
		// Open the table in an IExchangeModifyTable interface
		WC_H(lpFolder->OpenProperty(
			ulPropTag,
			(LPGUID)&IID_IExchangeModifyTable,
			0,
			MAPI_DEFERRED_ERRORS,
			(LPUNKNOWN FAR *)&lpExchTbl));
	}
	if (lpExchTbl)
	{
		WC_H(lpExchTbl->GetTable(NULL,&lpTbl));
	}
	if (lpTbl)
	{
		RegKeys[regkeyDEBUG_TAG].ulCurDWORD |= DBGGeneric;
		_OutputTable(DBGGeneric,NULL,lpTbl);
	}

	if (lpTbl) lpTbl->Release();
	if (lpExchTbl) lpExchTbl->Release();
	if (lpFolder) lpFolder->Release();
	if (lpMDB) lpMDB->Release();
	if (lpMAPISession) lpMAPISession->Release();
	MAPIUninitialize();
} // DumpExchangeTable

void DoAcls(_In_ MYOPTIONS ProgOpts)
{
	DumpExchangeTable(ProgOpts.lpszProfile,PR_ACL_TABLE,ProgOpts.ulFolder);
} // DoAcls