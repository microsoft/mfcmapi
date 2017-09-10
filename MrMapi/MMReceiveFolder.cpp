#include "stdafx.h"
#include "MrMAPI.h"
#include <MAPI/ColumnTags.h>
#include "MMReceiveFolder.h"

void PrintReceiveFolderTable(_In_ LPMDB lpMDB)
{
	auto hRes = S_OK;
	LPMAPITABLE lpReceiveFolderTable = nullptr;
	WC_MAPI(lpMDB->GetReceiveFolderTable(0, &lpReceiveFolderTable));
	if (FAILED(hRes))
	{
		printf("<receivefoldertable error=0x%x />\n", hRes);
		return;
	}

	printf("<receivefoldertable>\n");

	if (lpReceiveFolderTable)
	{
		auto sTags = LPSPropTagArray(&sptRECEIVECols);

		WC_MAPI(lpReceiveFolderTable->SetColumns(sTags, TBL_ASYNC));
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
			WC_MAPI(lpReceiveFolderTable->QueryRows(
				10,
				NULL,
				&lpRows));
			if (FAILED(hRes) || !lpRows || !lpRows->cRows) break;

			for (ULONG i = 0; i < lpRows->cRows; i++)
			{
				printf("<properties index=\"%u\">\n", iRow);
				_OutputProperties(DBGNoDebug, stdout, lpRows->aRow[i].cValues, lpRows->aRow[i].lpProps, nullptr, false);
				printf("</properties>\n");
				iRow++;
			}
		}

		if (lpRows) FreeProws(lpRows);
	}

	printf("</receivefoldertable>\n");
	if (lpReceiveFolderTable) { lpReceiveFolderTable->Release(); }
}  // PrintReceiveFolderTable

void DoReceiveFolder(_In_ MYOPTIONS ProgOpts)
{
	if (ProgOpts.lpMDB)
	{
		PrintReceiveFolderTable(ProgOpts.lpMDB);
	}
}