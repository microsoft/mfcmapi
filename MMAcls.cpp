#include "stdafx.h"
#include "MrMAPI.h"
#include "MMAcls.h"
#include "MMStore.h"

void DumpExchangeTable(_In_z_ LPWSTR lpszProfile, _In_ ULONG ulPropTag, _In_ ULONG ulFolder, _In_z_ LPWSTR lpszFolder)
{
	InitMFC();
	HRESULT hRes = S_OK;
	LPMAPISESSION lpMAPISession = NULL;
	LPMAPIFOLDER lpFolder = NULL;
	LPEXCHANGEMODIFYTABLE lpExchTbl = NULL;
	LPMAPITABLE lpTbl = NULL;

	WC_MAPI(MAPIInitialize(NULL));

	WC_H(MrMAPILogonEx(lpszProfile,&lpMAPISession));

	if (lpMAPISession)
	{
		WC_H(HrMAPIOpenStoreAndFolder(lpMAPISession, ulFolder, lpszFolder, NULL, &lpFolder));

		if (lpFolder)
		{
			// Open the table in an IExchangeModifyTable interface
			WC_MAPI(lpFolder->OpenProperty(
				ulPropTag,
				(LPGUID)&IID_IExchangeModifyTable,
				0,
				MAPI_DEFERRED_ERRORS,
				(LPUNKNOWN*)&lpExchTbl));
		}
		if (lpExchTbl)
		{
			WC_MAPI(lpExchTbl->GetTable(NULL,&lpTbl));
		}
		if (lpTbl)
		{
			RegKeys[regkeyDEBUG_TAG].ulCurDWORD |= DBGGeneric;
			_OutputTable(DBGGeneric,NULL,lpTbl);
		}
	}

	if (lpTbl) lpTbl->Release();
	if (lpExchTbl) lpExchTbl->Release();
	if (lpFolder) lpFolder->Release();
	if (lpMAPISession) lpMAPISession->Release();
	MAPIUninitialize();
} // DumpExchangeTable

void DoAcls(_In_ MYOPTIONS ProgOpts)
{
	DumpExchangeTable(ProgOpts.lpszProfile,PR_ACL_TABLE,ProgOpts.ulFolder,ProgOpts.lpszFolderPath);
} // DoAcls