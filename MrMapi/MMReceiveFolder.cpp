#include <StdAfx.h>
#include <MrMapi/MrMAPI.h>
#include <MAPI/ColumnTags.h>
#include <MrMapi/MMReceiveFolder.h>

void PrintReceiveFolderTable(_In_ LPMDB lpMDB)
{
	auto hRes = S_OK;
	LPMAPITABLE lpReceiveFolderTable = nullptr;
	WC_MAPI(lpMDB->GetReceiveFolderTable(0, &lpReceiveFolderTable));
	if (FAILED(hRes))
	{
		printf("<receivefoldertable error=0x%lx />\n", hRes);
		return;
	}

	printf("<receivefoldertable>\n");

	if (lpReceiveFolderTable)
	{
		WC_MAPI(lpReceiveFolderTable->SetColumns(&columns::sptRECEIVECols.tags, TBL_ASYNC));
	}

	if (SUCCEEDED(hRes))
	{
		LPSRowSet lpRows = nullptr;
		ULONG iRow = 0;

		for (;;)
		{
			hRes = S_OK;
			if (lpRows) FreeProws(lpRows);
			lpRows = nullptr;
			WC_MAPI(lpReceiveFolderTable->QueryRows(10, NULL, &lpRows));
			if (FAILED(hRes) || !lpRows || !lpRows->cRows) break;

			for (ULONG i = 0; i < lpRows->cRows; i++)
			{
				printf("<properties index=\"%lu\">\n", iRow);
				output::_OutputProperties(
					DBGNoDebug, stdout, lpRows->aRow[i].cValues, lpRows->aRow[i].lpProps, nullptr, false);
				printf("</properties>\n");
				iRow++;
			}
		}

		if (lpRows) FreeProws(lpRows);
	}

	printf("</receivefoldertable>\n");
	if (lpReceiveFolderTable)
	{
		lpReceiveFolderTable->Release();
	}
} // PrintReceiveFolderTable

void DoReceiveFolder(_In_ cli::MYOPTIONS ProgOpts)
{
	if (ProgOpts.lpMDB)
	{
		PrintReceiveFolderTable(ProgOpts.lpMDB);
	}
}