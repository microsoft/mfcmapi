// Export profile to file

#include <StdAfx.h>
#include <MAPI/MAPIProfileFunctions.h>
#include <MAPI/MAPIFunctions.h>
#include <Interpret/GUIDArray.h>

void ExportProfileSection(FILE* fProfile, LPPROFSECT lpSect, LPSBinary lpSectBin)
{
	if (!fProfile || !lpSect) return;

	auto hRes = S_OK;
	LPSPropValue lpAllProps = nullptr;
	ULONG cValues = 0L;

	WC_H_GETPROPS(mapi::GetPropsNULL(lpSect,
		fMapiUnicode,
		&cValues,
		&lpAllProps));
	if (FAILED(hRes))
	{
		OutputToFilef(fProfile, L"<properties error=\"0x%08X\" />\n", hRes);
	}
	else if (lpAllProps)
	{
		std::wstring szBin;
		if (lpSectBin)
		{
			szBin = strings::BinToHexString(lpSectBin, false);
		}

		OutputToFilef(fProfile, L"<properties listtype=\"profilesection\" profilesection=\"%ws\">\n", szBin.c_str());

		OutputPropertiesToFile(fProfile, cValues, lpAllProps, nullptr, false);

		OutputToFile(fProfile, L"</properties>\n");

		MAPIFreeBuffer(lpAllProps);
	}
}

void ExportProfileProvider(FILE* fProfile, int iRow, LPPROVIDERADMIN lpProviderAdmin, LPSRow lpRow)
{
	if (!fProfile || !lpRow) return;

	Outputf(DBGNoDebug, fProfile, true, L"<provider index = \"0x%08X\">\n", iRow);

	OutputToFile(fProfile, L"<properties listtype=\"row\">\n");
	OutputSRowToFile(fProfile, lpRow, nullptr);
	OutputToFile(fProfile, L"</properties>\n");

	auto hRes = S_OK;
	auto lpProviderUID = PpropFindProp(
		lpRow->lpProps,
		lpRow->cValues,
		PR_PROVIDER_UID);

	if (lpProviderUID)
	{
		LPPROFSECT lpSect = nullptr;
		EC_H(mapi::profile::OpenProfileSection(
			lpProviderAdmin,
			&lpProviderUID->Value.bin,
			&lpSect));
		if (lpSect)
		{
			ExportProfileSection(fProfile, lpSect, &lpProviderUID->Value.bin);
			lpSect->Release();
		}
	}

	OutputToFile(fProfile, L"</provider>\n");
}

void ExportProfileService(FILE* fProfile, int iRow, LPSERVICEADMIN lpServiceAdmin, LPSRow lpRow)
{
	if (!fProfile || !lpRow) return;

	Outputf(DBGNoDebug, fProfile, true, L"<service index = \"0x%08X\">\n", iRow);

	OutputToFile(fProfile, L"<properties listtype=\"row\">\n");
	OutputSRowToFile(fProfile, lpRow, nullptr);
	OutputToFile(fProfile, L"</properties>\n");

	auto hRes = S_OK;
	auto lpServiceUID = PpropFindProp(
		lpRow->lpProps,
		lpRow->cValues,
		PR_SERVICE_UID);

	if (lpServiceUID)
	{
		LPPROFSECT lpSect = nullptr;
		EC_H(mapi::profile::OpenProfileSection(
			lpServiceAdmin,
			&lpServiceUID->Value.bin,
			&lpSect));
		if (lpSect)
		{
			ExportProfileSection(fProfile, lpSect, &lpServiceUID->Value.bin);
			lpSect->Release();
		}

		LPPROVIDERADMIN lpProviderAdmin = nullptr;

		EC_MAPI(lpServiceAdmin->AdminProviders(
			reinterpret_cast<LPMAPIUID>(lpServiceUID->Value.bin.lpb),
			0, // fMapiUnicode is not supported
			&lpProviderAdmin));

		if (lpProviderAdmin)
		{
			LPMAPITABLE lpProviderTable = nullptr;
			EC_MAPI(lpProviderAdmin->GetProviderTable(
				0, // fMapiUnicode is not supported
				&lpProviderTable));

			if (lpProviderTable)
			{
				LPSRowSet lpRowSet = nullptr;
				EC_MAPI(HrQueryAllRows(lpProviderTable, nullptr, nullptr, nullptr, 0, &lpRowSet));
				if (lpRowSet && lpRowSet->cRows >= 1)
				{
					for (ULONG i = 0; i < lpRowSet->cRows; i++)
					{
						ExportProfileProvider(fProfile, i, lpProviderAdmin, &lpRowSet->aRow[i]);
					}
				}

				FreeProws(lpRowSet);
				lpProviderTable->Release();
				lpProviderTable = nullptr;
			}

			lpProviderAdmin->Release();
			lpProviderAdmin = nullptr;
		}
	}

	OutputToFile(fProfile, L"</service>\n");
}

void ExportProfile(_In_ const std::string& szProfile, _In_ const std::wstring& szProfileSection, bool bByteSwapped, const std::wstring& szFileName)
{
	if (szProfile.empty()) return;

	DebugPrint(DBGGeneric, L"ExportProfile: Saving profile \"%hs\" to \"%ws\"\n", szProfile.c_str(), szFileName.c_str());
	if (!szProfileSection.empty())
	{
		DebugPrint(DBGGeneric, L"ExportProfile: Restricting to \"%ws\"\n", szProfileSection.c_str());
	}

	auto hRes = S_OK;
	LPPROFADMIN lpProfAdmin = nullptr;
	FILE* fProfile = nullptr;

	if (!szFileName.empty())
	{
		fProfile = MyOpenFile(szFileName, true);
	}

	OutputToFile(fProfile, g_szXMLHeader);
	Outputf(DBGNoDebug, fProfile, true, L"<profile profilename= \"%hs\">\n", szProfile.c_str());

	EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));

	if (lpProfAdmin)
	{
		LPSERVICEADMIN lpServiceAdmin = nullptr;
		EC_MAPI(lpProfAdmin->AdminServices(
			LPTSTR(szProfile.c_str()),
			LPTSTR(""),
			NULL,
			MAPI_DIALOG,
			&lpServiceAdmin));
		if (lpServiceAdmin)
		{
			if (!szProfileSection.empty())
			{
				const auto lpGuid = guid::GUIDNameToGUID(szProfileSection, bByteSwapped);

				if (lpGuid)
				{
					LPPROFSECT lpSect = nullptr;
					SBinary sBin = { 0 };
					sBin.cb = sizeof(GUID);
					sBin.lpb = LPBYTE(lpGuid);

					EC_H(mapi::profile::OpenProfileSection(
						lpServiceAdmin,
						&sBin,
						&lpSect));

					ExportProfileSection(fProfile, lpSect, &sBin);
					delete[] lpGuid;
				}
			}
			else
			{
				LPMAPITABLE lpServiceTable = nullptr;

				EC_MAPI(lpServiceAdmin->GetMsgServiceTable(
					0, // fMapiUnicode is not supported
					&lpServiceTable));

				if (lpServiceTable)
				{
					LPSRowSet lpRowSet = nullptr;
					EC_MAPI(HrQueryAllRows(lpServiceTable, nullptr, nullptr, nullptr, 0, &lpRowSet));
					if (lpRowSet && lpRowSet->cRows >= 1)
					{
						for (ULONG i = 0; i < lpRowSet->cRows; i++)
						{
							ExportProfileService(fProfile, i, lpServiceAdmin, &lpRowSet->aRow[i]);
						}
					}

					FreeProws(lpRowSet);
					lpServiceTable->Release();
				}
			}
			lpServiceAdmin->Release();
		}
		lpProfAdmin->Release();

	}

	OutputToFile(fProfile, L"</profile>");

	if (fProfile)
	{
		CloseFile(fProfile);
	}
}