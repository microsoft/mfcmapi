#include "stdafx.h"
#include "..\stdafx.h"
#include "MrMAPI.h"
#include "MMProfile.h"
#include "..\ExportProfile.h"
#include "..\MAPIFunctions.h"
#include "..\String.h"
#include "..\MAPIProfileFunctions.h"

void ExportProfileList()
{
	printf("Profile List\n");
	printf(" # Default Name\n");
	HRESULT hRes = S_OK;
	LPMAPITABLE lpProfTable = NULL;
	LPPROFADMIN lpProfAdmin = NULL;

	static const SizedSPropTagArray(2, rgPropTag) =
	{
		2,
		PR_DEFAULT_PROFILE,
		PR_DISPLAY_NAME_A,
	};

	EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));
	if (!lpProfAdmin) return;

	EC_MAPI(lpProfAdmin->GetProfileTable(
		0, // fMapiUnicode is not supported
		&lpProfTable));

	if (lpProfTable)
	{
		LPSRowSet lpRows = NULL;

		EC_MAPI(HrQueryAllRows(
			lpProfTable,
			(LPSPropTagArray)&rgPropTag,
			NULL,
			NULL,
			0,
			&lpRows));

		if (lpRows)
		{
			if (lpRows->cRows == 0)
			{
				printf("No profiles exist\n");
				hRes = S_OK;
			}
			else
			{
				ULONG i = 0;
				if (!FAILED(hRes)) for (i = 0; i < lpRows->cRows; i++)
				{
					printf("%2d ", i);
					if (PR_DEFAULT_PROFILE == lpRows->aRow[i].lpProps[0].ulPropTag && lpRows->aRow[i].lpProps[0].Value.b)
					{
						printf("*       ");
					}
					else
					{
						printf("        ");
					}

					if (CheckStringProp(&lpRows->aRow[i].lpProps[1], PT_STRING8))
					{
						printf("%hs\n", lpRows->aRow[i].lpProps[1].Value.lpszA);
					}
					else
					{
						printf("UNKNOWN\n");
					}
				}
			}

			FreeProws(lpRows);
		}

		lpProfTable->Release();
	}

	lpProfAdmin->Release();
}

void DoProfile(_In_ MYOPTIONS ProgOpts)
{
	if (ProgOpts.lpszProfile && ProgOpts.lpszOutput)
	{
		printf("Profile Export\n");
		printf("Options specified:\n");
		printf("   Profile: %ws\n", ProgOpts.lpszProfile);
		printf("   Output File: %ws\n", ProgOpts.lpszOutput);
		LPSTR szProfileA = NULL;
		(void)UnicodeToAnsi(ProgOpts.lpszProfile, &szProfileA);
		if (szProfileA)
		{
			ExportProfile(szProfileA, ProgOpts.lpszProfileSection, ProgOpts.bByteSwapped, ProgOpts.lpszOutput);
		}

		delete[] szProfileA;
	}
	else
	{
		ExportProfileList();
	}
}