// Export profile to file

#include <core/stdafx.h>
#include <core/mapi/exportProfile.h>
#include <core/mapi/mapiProfileFunctions.h>
#include <core/interpret/guid.h>
#include <core/utility/strings.h>
#include <core/mapi/mapiOutput.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFunctions.h>
#include <core/utility/error.h>

namespace output
{
	void ExportProfileSection(FILE* fProfile, LPPROFSECT lpSect, const SBinary* lpSectBin)
	{
		if (!fProfile || !lpSect) return;

		LPSPropValue lpAllProps = nullptr;
		ULONG cValues = 0L;

		const auto hRes = WC_H_GETPROPS(mapi::GetPropsNULL(lpSect, fMapiUnicode, &cValues, &lpAllProps));
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

			OutputToFilef(
				fProfile, L"<properties listtype=\"profilesection\" profilesection=\"%ws\">\n", szBin.c_str());

			outputProperties(dbgLevel::NoDebug, fProfile, cValues, lpAllProps, nullptr, false);

			OutputToFile(fProfile, L"</properties>\n");

			MAPIFreeBuffer(lpAllProps);
		}
	}

	void ExportProfileProvider(FILE* fProfile, int iRow, LPPROVIDERADMIN lpProviderAdmin, LPSRow lpRow)
	{
		if (!fProfile || !lpRow) return;

		Outputf(dbgLevel::NoDebug, fProfile, true, L"<provider index = \"0x%08X\">\n", iRow);

		OutputToFile(fProfile, L"<properties listtype=\"row\">\n");
		outputSRow(dbgLevel::NoDebug, fProfile, lpRow, nullptr);
		OutputToFile(fProfile, L"</properties>\n");

		auto lpProviderUID = PpropFindProp(lpRow->lpProps, lpRow->cValues, PR_PROVIDER_UID);

		if (lpProviderUID)
		{
			auto lpSect = mapi::profile::OpenProfileSection(lpProviderAdmin, &mapi::getBin(lpProviderUID));
			if (lpSect)
			{
				ExportProfileSection(fProfile, lpSect, &mapi::getBin(lpProviderUID));
				lpSect->Release();
			}
		}

		OutputToFile(fProfile, L"</provider>\n");
	}

	void ExportProfileService(FILE* fProfile, int iRow, LPSERVICEADMIN lpServiceAdmin, LPSRow lpRow)
	{
		if (!fProfile || !lpRow) return;

		Outputf(dbgLevel::NoDebug, fProfile, true, L"<service index = \"0x%08X\">\n", iRow);

		OutputToFile(fProfile, L"<properties listtype=\"row\">\n");
		outputSRow(dbgLevel::NoDebug, fProfile, lpRow, nullptr);
		OutputToFile(fProfile, L"</properties>\n");

		const auto lpServiceUID = PpropFindProp(lpRow->lpProps, lpRow->cValues, PR_SERVICE_UID);

		if (lpServiceUID)
		{
			auto bin = mapi::getBin(lpServiceUID);
			auto lpSect = mapi::profile::OpenProfileSection(lpServiceAdmin, &bin);
			if (lpSect)
			{
				ExportProfileSection(fProfile, lpSect, &bin);
				lpSect->Release();
			}

			LPPROVIDERADMIN lpProviderAdmin = nullptr;

			EC_MAPI_S(lpServiceAdmin->AdminProviders(
				reinterpret_cast<LPMAPIUID>(bin.lpb),
				0, // fMapiUnicode is not supported
				&lpProviderAdmin));

			if (lpProviderAdmin)
			{
				LPMAPITABLE lpProviderTable = nullptr;
				EC_MAPI_S(lpProviderAdmin->GetProviderTable(
					0, // fMapiUnicode is not supported
					&lpProviderTable));

				if (lpProviderTable)
				{
					LPSRowSet lpRowSet = nullptr;
					EC_MAPI_S(HrQueryAllRows(lpProviderTable, nullptr, nullptr, nullptr, 0, &lpRowSet));
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

	void ExportProfile(
		_In_ const std::wstring& szProfile,
		_In_ const std::wstring& szProfileSection,
		bool bByteSwapped,
		const std::wstring& szFileName)
	{
		if (szProfile.empty()) return;

		DebugPrint(
			dbgLevel::Generic,
			L"ExportProfile: Saving profile \"%ws\" to \"%ws\"\n",
			szProfile.c_str(),
			szFileName.c_str());
		if (!szProfileSection.empty())
		{
			DebugPrint(dbgLevel::Generic, L"ExportProfile: Restricting to \"%ws\"\n", szProfileSection.c_str());
		}

		LPPROFADMIN lpProfAdmin = nullptr;
		FILE* fProfile = nullptr;

		if (!szFileName.empty())
		{
			fProfile = MyOpenFile(szFileName, true);
		}

		OutputToFile(fProfile, g_szXMLHeader);
		Outputf(dbgLevel::NoDebug, fProfile, true, L"<profile profilename= \"%ws\">\n", szProfile.c_str());

		EC_MAPI_S(MAPIAdminProfiles(0, &lpProfAdmin));
		if (lpProfAdmin)
		{
			LPSERVICEADMIN lpServiceAdmin = nullptr;
			EC_MAPI_S(lpProfAdmin->AdminServices(
				strings::LPCSTRToLPTSTR(strings::wstringTostring(szProfile).c_str()),
				strings::LPCSTRToLPTSTR(""),
				NULL,
				MAPI_DIALOG,
				&lpServiceAdmin));
			if (lpServiceAdmin)
			{
				if (!szProfileSection.empty())
				{
					const auto guid = guid::GUIDNameToGUID(szProfileSection, bByteSwapped);
					auto sBin = SBinary{sizeof(GUID), reinterpret_cast<LPBYTE>(const_cast<GUID*>(&guid))};

					auto lpSect = mapi::profile::OpenProfileSection(lpServiceAdmin, &sBin);
					if (lpSect)
					{
						ExportProfileSection(fProfile, lpSect, &sBin);
						lpSect->Release();
					}
				}
				else
				{
					LPMAPITABLE lpServiceTable = nullptr;
					EC_MAPI_S(lpServiceAdmin->GetMsgServiceTable(
						0, // fMapiUnicode is not supported
						&lpServiceTable));
					if (lpServiceTable)
					{
						LPSRowSet lpRowSet = nullptr;
						EC_MAPI_S(HrQueryAllRows(lpServiceTable, nullptr, nullptr, nullptr, 0, &lpRowSet));
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
} // namespace output
