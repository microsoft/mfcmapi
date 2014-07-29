#include "stdafx.h"
#include "MrMAPI.h"
#include "ColumnTags.h"
#include "MMReceiveFolder.h"

void PrintReceiveFolderTable(_In_ LPMDB lpMDB)
{
	HRESULT hRes = S_OK;
	LPMAPITABLE lpReceiveFolderTable = NULL;
	WC_MAPI(lpMDB->GetReceiveFolderTable(0, &lpReceiveFolderTable));
	if (FAILED(hRes))
	{
		printf(_T("<receivefoldertable error=0x%x />\n"), hRes);
		return;
	}

	printf(_T("<receivefoldertable>\n"));

	if (lpReceiveFolderTable)
	{
		LPSPropTagArray sTags = (LPSPropTagArray)&sptRECEIVECols;

		WC_MAPI(lpReceiveFolderTable->SetColumns(sTags, TBL_ASYNC));
	}

	if (SUCCEEDED(hRes))
	{
		LPSRowSet lpRows = NULL;
		ULONG iRow = 0;

		for (;;)
		{
			hRes = S_OK;
			if (lpRows) FreeProws(lpRows);
			lpRows = NULL;
			WC_MAPI(lpReceiveFolderTable->QueryRows(
				10,
				NULL,
				&lpRows));
			if (FAILED(hRes) || !lpRows || !lpRows->cRows) break;

			ULONG i = 0;
			for (i = 0; i < lpRows->cRows; i++)
			{
				printf(_T("<properties index=\"%u\">\n"), iRow);
				_OutputProperties(DBGNoDebug, stdout, lpRows->aRow[i].cValues, lpRows->aRow[i].lpProps, NULL, false);
				printf(_T("</properties>\n"));
				iRow++;
			}
		}

		if (lpRows) FreeProws(lpRows);
	}

	printf(_T("</receivefoldertable>\n"));
	if (lpReceiveFolderTable) { lpReceiveFolderTable->Release(); }
}  // PrintReceiveFolderTable

void DoReceiveFolder(_In_ MYOPTIONS ProgOpts)
{
	if (ProgOpts.lpMDB)
	{
		PrintReceiveFolderTable(ProgOpts.lpMDB);
	}
}