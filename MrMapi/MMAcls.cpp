#include <StdAfx.h>
#include <MrMapi/MrMAPI.h>
#include <MrMapi/MMAcls.h>

void DumpExchangeTable(_In_ ULONG ulPropTag, _In_ LPMAPIFOLDER lpFolder)
{
	auto hRes = S_OK;
	LPEXCHANGEMODIFYTABLE lpExchTbl = nullptr;
	LPMAPITABLE lpTbl = nullptr;

	if (lpFolder)
	{
		// Open the table in an IExchangeModifyTable interface
		WC_MAPI(lpFolder->OpenProperty(
			ulPropTag,
			const_cast<LPGUID>(&IID_IExchangeModifyTable),
			0,
			MAPI_DEFERRED_ERRORS,
			reinterpret_cast<LPUNKNOWN*>(&lpExchTbl)));
		if (lpExchTbl)
		{
			WC_MAPI(lpExchTbl->GetTable(NULL, &lpTbl));
		}
		if (lpTbl)
		{
			RegKeys[regkeyDEBUG_TAG].ulCurDWORD |= DBGGeneric;
			_OutputTable(DBGGeneric, nullptr, lpTbl);
		}
	}

	if (lpTbl) lpTbl->Release();
	if (lpExchTbl) lpExchTbl->Release();
}

void DoAcls(_In_ MYOPTIONS ProgOpts)
{
	DumpExchangeTable(PR_ACL_TABLE, ProgOpts.lpFolder);
}