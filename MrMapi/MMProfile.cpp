#include "stdafx.h"
#include "MrMAPI.h"
#include "MMProfile.h"
#include "ExportProfile.h"
#include "MAPIFunctions.h"
#include "String.h"

void ExportProfileList()
{
	printf("Profile List\n");
	printf(" # Default Name\n");
	auto hRes = S_OK;
	LPMAPITABLE lpProfTable = nullptr;
	LPPROFADMIN lpProfAdmin = nullptr;

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
		LPSRowSet lpRows = nullptr;

		EC_MAPI(HrQueryAllRows(
			lpProfTable,
			LPSPropTagArray(&rgPropTag),
			NULL,
			NULL,
			0,
			&lpRows));

		if (lpRows)
		{
			if (lpRows->cRows == 0)
			{
				printf("No profiles exist\n");
			}
			else
			{
				if (!FAILED(hRes)) for (ULONG i = 0; i < lpRows->cRows; i++)
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
	if (!ProgOpts.lpszProfile.empty() && !ProgOpts.lpszOutput.empty())
	{
		printf("Profile Export\n");
		printf("Options specified:\n");
		printf("   Profile: %ws\n", ProgOpts.lpszProfile.c_str());
		printf("   Output File: %ws\n", ProgOpts.lpszOutput.c_str());
		ExportProfile(wstringTostring(ProgOpts.lpszProfile).c_str(), ProgOpts.lpszProfileSection.c_str(), ProgOpts.bByteSwapped, ProgOpts.lpszOutput);
	}
	else
	{
		ExportProfileList();
	}
}