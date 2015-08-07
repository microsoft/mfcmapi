#include "stdafx.h"
#include "..\stdafx.h"
#include "MrMAPI.h"
#include "MMAcls.h"
#include "MMStore.h"

void DumpExchangeTable(_In_ ULONG ulPropTag, _In_ LPMAPIFOLDER lpFolder)
{
	HRESULT hRes = S_OK;
	LPEXCHANGEMODIFYTABLE lpExchTbl = NULL;
	LPMAPITABLE lpTbl = NULL;

	if (lpFolder)
	{
		// Open the table in an IExchangeModifyTable interface
		WC_MAPI(lpFolder->OpenProperty(
			ulPropTag,
			(LPGUID)&IID_IExchangeModifyTable,
			0,
			MAPI_DEFERRED_ERRORS,
			(LPUNKNOWN*)&lpExchTbl));
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
} // DumpExchangeTable

void DoAcls(_In_ MYOPTIONS ProgOpts)
{
	DumpExchangeTable(PR_ACL_TABLE, ProgOpts.lpFolder);
} // DoAcls