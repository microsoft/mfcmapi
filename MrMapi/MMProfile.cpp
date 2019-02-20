#include <StdAfx.h>
#include <MrMapi/MMProfile.h>
#include <MrMapi/mmcli.h>
#include <core/mapi/exportProfile.h>
#include <core/utility/strings.h>

namespace output
{
	void ExportProfileList()
	{
		printf("Profile List\n");
		printf(" # Default Name\n");
		LPMAPITABLE lpProfTable = nullptr;
		LPPROFADMIN lpProfAdmin = nullptr;

		static const SizedSPropTagArray(2, rgPropTag) = {
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

			EC_MAPI(HrQueryAllRows(lpProfTable, LPSPropTagArray(&rgPropTag), NULL, NULL, 0, &lpRows));

			if (lpRows)
			{
				if (lpRows->cRows == 0)
				{
					printf("No profiles exist\n");
				}
				else
				{
					for (ULONG i = 0; i < lpRows->cRows; i++)
					{
						printf("%2d ", i);
						if (PR_DEFAULT_PROFILE == lpRows->aRow[i].lpProps[0].ulPropTag &&
							lpRows->aRow[i].lpProps[0].Value.b)
						{
							printf("*       ");
						}
						else
						{
							printf("        ");
						}

						if (strings::CheckStringProp(&lpRows->aRow[i].lpProps[1], PT_STRING8))
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

	void DoProfile(_In_ cli::MYOPTIONS ProgOpts)
	{
		if (cli::switchProfile.hasArg(0) && !ProgOpts.lpszOutput.empty())
		{
			const auto szProfile = cli::switchProfile.getArg(0);
			const auto szProfileSection = cli::switchProfileSection.getArg(0);
			printf("Profile Export\n");
			printf("Options specified:\n");
			printf("   Profile: %ws\n", szProfile.c_str());
			printf("   Profile section: %ws\n", szProfileSection.c_str());
			printf("   Output File: %ws\n", ProgOpts.lpszOutput.c_str());
			ExportProfile(
				strings::wstringTostring(szProfile),
				szProfileSection,
				cli::switchByteSwapped.isSet(),
				ProgOpts.lpszOutput);
		}
		else
		{
			ExportProfileList();
		}
	}
} // namespace output