#include <StdAfx.h>
#include <MrMapi/MMReceiveFolder.h>
#include <MrMapi/mmcli.h>
#include <core/mapi/columnTags.h>
#include <core/mapi/mapiOutput.h>
#include <core/utility/output.h>

void PrintReceiveFolderTable(_In_ LPMDB lpMDB)
{
	LPMAPITABLE lpReceiveFolderTable = nullptr;
	auto hRes = WC_MAPI(lpMDB->GetReceiveFolderTable(0, &lpReceiveFolderTable));
	if (FAILED(hRes))
	{
		wprintf(L"<receivefoldertable error=0x%lx />\n", hRes);
		return;
	}

	wprintf(L"<receivefoldertable>\n");

	if (lpReceiveFolderTable)
	{
		hRes = WC_MAPI(lpReceiveFolderTable->SetColumns(&columns::sptRECEIVECols.tags, TBL_ASYNC));
	}

	if (SUCCEEDED(hRes))
	{
		LPSRowSet lpRows = nullptr;
		ULONG iRow = 0;

		for (;;)
		{
			if (lpRows) FreeProws(lpRows);
			lpRows = nullptr;
			hRes = WC_MAPI(lpReceiveFolderTable->QueryRows(10, NULL, &lpRows));
			if (FAILED(hRes) || !lpRows || !lpRows->cRows) break;

			for (ULONG i = 0; i < lpRows->cRows; i++)
			{
				wprintf(L"<properties index=\"%lu\">\n", iRow);
				output::outputProperties(
					output::dbgLevel::NoDebug,
					stdout,
					lpRows->aRow[i].cValues,
					lpRows->aRow[i].lpProps,
					nullptr,
					false);
				wprintf(L"</properties>\n");
				iRow++;
			}
		}

		if (lpRows) FreeProws(lpRows);
	}

	wprintf(L"</receivefoldertable>\n");
	if (lpReceiveFolderTable)
	{
		lpReceiveFolderTable->Release();
	}
}

void DoReceiveFolder(_In_opt_ LPMDB lpMDB)
{
	if (lpMDB)
	{
		PrintReceiveFolderTable(lpMDB);
	}
}