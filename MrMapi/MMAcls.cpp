#include <StdAfx.h>
#include <MrMapi/MMAcls.h>
#include <MrMapi/cli.h>
#include <IO/Registry.h>
#include <IO/MFCOutput.h>
#include <IO/output.h>

void DumpExchangeTable(_In_ ULONG ulPropTag, _In_ LPMAPIFOLDER lpFolder)
{
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
			registry::debugTag |= DBGGeneric;
			output::_OutputTable(DBGGeneric, nullptr, lpTbl);
		}
	}

	if (lpTbl) lpTbl->Release();
	if (lpExchTbl) lpExchTbl->Release();
}

void DoAcls(_In_ cli::MYOPTIONS ProgOpts) { DumpExchangeTable(PR_ACL_TABLE, ProgOpts.lpFolder); }