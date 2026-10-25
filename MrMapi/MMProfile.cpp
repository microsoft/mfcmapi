#include <StdAfx.h>
#include <MrMapi/MMProfile.h>
#include <MrMapi/mmcli.h>
#include <core/mapi/exportProfile.h>
#include <core/utility/strings.h>

namespace output
{
	void ExportProfileList()
	{
		wprintf_s(L"Profile List\n");
		wprintf_s(L" # Default Name\n");
		LPMAPITABLE lpProfTable = nullptr;
		LPPROFADMIN lpProfAdmin = nullptr;

		static const SizedSPropTagArray(2, rgPropTag) = {
			2,
			PR_DEFAULT_PROFILE,
			PR_DISPLAY_NAME_A,
		};

		EC_MAPI_S(MAPIAdminProfiles(0, &lpProfAdmin));
		if (!lpProfAdmin) return;

		EC_MAPI_S(lpProfAdmin->GetProfileTable(
			0, // fMapiUnicode is not supported
			&lpProfTable));

		if (lpProfTable)
		{
			LPSRowSet lpRows = nullptr;

			EC_MAPI_S(HrQueryAllRows(lpProfTable, LPSPropTagArray(&rgPropTag), nullptr, nullptr, 0, &lpRows));

			if (lpRows)
			{
				if (lpRows->cRows == 0)
				{
					wprintf_s(L"No profiles exist\n");
				}
				else
				{
					for (ULONG i = 0; i < lpRows->cRows; i++)
					{
						wprintf_s(L"%2ul ", i);
						if (PR_DEFAULT_PROFILE == lpRows->aRow[i].lpProps[0].ulPropTag &&
							lpRows->aRow[i].lpProps[0].Value.b)
						{
							wprintf_s(L"*       ");
						}
						else
						{
							wprintf_s(L"        ");
						}

						if (strings::CheckStringProp(&lpRows->aRow[i].lpProps[1], PT_STRING8))
						{
							wprintf_s(L"%hs\n", lpRows->aRow[i].lpProps[1].Value.lpszA);
						}
						else
						{
							wprintf_s(L"UNKNOWN\n");
						}
					}
				}

				FreeProws(lpRows);
			}

			lpProfTable->Release();
		}

		lpProfAdmin->Release();
	}

	void DoProfile()
	{
		const auto output = cli::switchOutput[0];
		if (cli::switchProfile.has(0) && !output.empty())
		{
			const auto szProfile = cli::switchProfile[0];
			const auto szProfileSection = cli::switchProfileSection[0];
			wprintf_s(L"Profile Export\n");
			wprintf_s(L"Options specified:\n");
			wprintf_s(L"   Profile: %ws\n", szProfile.c_str());
			wprintf_s(L"   Profile section: %ws\n", szProfileSection.c_str());
			wprintf_s(L"   Output File: %ws\n", output.c_str());
			ExportProfile(szProfile, szProfileSection, cli::switchByteSwapped.isSet(), output);
		}
		else
		{
			ExportProfileList();
		}
	}
} // namespace output