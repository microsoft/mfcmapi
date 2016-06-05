// ExportProfile.cpp : implementation file
// Export profile to file

#include "stdafx.h"
#include "MAPIProfileFunctions.h"
#include "MAPIFunctions.h"
#include "InterpretProp2.h"

void ExportProfileSection(FILE* fProfile, LPPROFSECT lpSect, LPSBinary lpSectBin)
{
	if (!fProfile || !lpSect) return;

	HRESULT hRes = S_OK;
	LPSPropValue lpAllProps = NULL;
	ULONG cValues = 0L;

	WC_H_GETPROPS(GetPropsNULL(lpSect,
		fMapiUnicode,
		&cValues,
		&lpAllProps));
	if (FAILED(hRes))
	{
		OutputToFilef(fProfile, L"<properties error=\"0x%08X\" />\n", hRes);
	}
	else if (lpAllProps)
	{
		LPTSTR szBin = NULL;
		if (lpSectBin)
		{
			MyHexFromBin(lpSectBin->lpb, lpSectBin->cb, false, &szBin);
		}
		OutputToFilef(fProfile, L"<properties listtype=\"profilesection\" profilesection=\"%ws\">\n", LPCTSTRToWstring(szBin).c_str());

		delete[] szBin;

		OutputPropertiesToFile(fProfile, cValues, lpAllProps, NULL, false);

		OutputToFile(fProfile, L"</properties>\n");

		MAPIFreeBuffer(lpAllProps);
	}
}

void ExportProfileProvider(FILE* fProfile, int iRow, LPPROVIDERADMIN lpProviderAdmin, LPSRow lpRow)
{
	if (!fProfile || !lpRow) return;

	Outputf(DBGNoDebug, fProfile, true, L"<provider index = \"0x%08X\">\n", iRow);

	OutputToFile(fProfile, L"<properties listtype=\"row\">\n");
	OutputSRowToFile(fProfile, lpRow, NULL);
	OutputToFile(fProfile, L"</properties>\n");

	HRESULT hRes = S_OK;
	LPSPropValue lpProviderUID = NULL;

	lpProviderUID = PpropFindProp(
		lpRow->lpProps,
		lpRow->cValues,
		PR_PROVIDER_UID);

	if (lpProviderUID)
	{
		LPPROFSECT lpSect = NULL;
		EC_H(OpenProfileSection(
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
	OutputSRowToFile(fProfile, lpRow, NULL);
	OutputToFile(fProfile, L"</properties>\n");

	HRESULT hRes = S_OK;
	LPSPropValue lpServiceUID = NULL;

	lpServiceUID = PpropFindProp(
		lpRow->lpProps,
		lpRow->cValues,
		PR_SERVICE_UID);

	if (lpServiceUID)
	{
		LPPROFSECT lpSect = NULL;
		EC_H(OpenProfileSection(
			lpServiceAdmin,
			&lpServiceUID->Value.bin,
			&lpSect));
		if (lpSect)
		{
			ExportProfileSection(fProfile, lpSect, &lpServiceUID->Value.bin);
			lpSect->Release();
		}

		LPPROVIDERADMIN lpProviderAdmin = NULL;

		EC_MAPI(lpServiceAdmin->AdminProviders(
			(LPMAPIUID)lpServiceUID->Value.bin.lpb,
			0, // fMapiUnicode is not supported
			&lpProviderAdmin));

		if (lpProviderAdmin)
		{
			LPMAPITABLE lpProviderTable = NULL;
			EC_MAPI(lpProviderAdmin->GetProviderTable(
				0, // fMapiUnicode is not supported
				&lpProviderTable));

			if (lpProviderTable)
			{
				LPSRowSet lpRowSet = NULL;
				EC_MAPI(HrQueryAllRows(lpProviderTable, NULL, NULL, NULL, 0, &lpRowSet));
				if (lpRowSet && lpRowSet->cRows >= 1)
				{
					for (ULONG i = 0; i < lpRowSet->cRows; i++)
					{
						ExportProfileProvider(fProfile, i, lpProviderAdmin, &lpRowSet->aRow[i]);
					}
				}

				FreeProws(lpRowSet);
				lpProviderTable->Release();
				lpProviderTable = NULL;
			}

			lpProviderAdmin->Release();
			lpProviderAdmin = NULL;
		}
	}

	OutputToFile(fProfile, L"</service>\n");
}

void ExportProfile(_In_z_ LPCSTR szProfile, _In_z_ LPCWSTR szProfileSection, bool bByteSwapped, wstring szFileName)
{
	if (!szProfile) return;

	DebugPrint(DBGGeneric, L"ExportProfile: Saving profile \"%hs\" to \"%ws\"\n", szProfile, szFileName.c_str());
	if (szProfileSection)
	{
		DebugPrint(DBGGeneric, L"ExportProfile: Restricting to \"%ws\"\n", szProfileSection);
	}

	HRESULT hRes = S_OK;
	LPPROFADMIN lpProfAdmin = NULL;
	FILE* fProfile = NULL;

	if (!szFileName.empty())
	{
		fProfile = MyOpenFile(szFileName, true);
	}

	OutputToFile(fProfile, g_szXMLHeader);
	Outputf(DBGNoDebug, fProfile, true, L"<profile profilename= \"%hs\">\n", szProfile);

	EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));

	if (lpProfAdmin)
	{
		LPSERVICEADMIN lpServiceAdmin = NULL;
#pragma warning(push)
#pragma warning(disable:4616)
#pragma warning(disable:6276)
		EC_MAPI(lpProfAdmin->AdminServices(
			(TCHAR*)szProfile,
			(TCHAR*)"",
			NULL,
			MAPI_DIALOG,
			&lpServiceAdmin));
#pragma warning(pop)
		if (lpServiceAdmin)
		{
			if (szProfileSection)
			{
				LPCGUID lpGuid = GUIDNameToGUID(szProfileSection, bByteSwapped);

				if (lpGuid)
				{
					LPPROFSECT lpSect = NULL;
					SBinary sBin = { 0 };
					sBin.cb = sizeof(GUID);
					sBin.lpb = (LPBYTE)lpGuid;

					EC_H(OpenProfileSection(
						lpServiceAdmin,
						&sBin,
						&lpSect));

					ExportProfileSection(fProfile, lpSect, &sBin);
					delete[] lpGuid;
				}
			}
			else
			{
				LPMAPITABLE lpServiceTable = NULL;

				EC_MAPI(lpServiceAdmin->GetMsgServiceTable(
					0, // fMapiUnicode is not supported
					&lpServiceTable));

				if (lpServiceTable)
				{
					LPSRowSet lpRowSet = NULL;
					EC_MAPI(HrQueryAllRows(lpServiceTable, NULL, NULL, NULL, 0, &lpRowSet));
					if (lpRowSet && lpRowSet->cRows >= 1)
					{
						ULONG i = 0;
						for (i = 0; i < lpRowSet->cRows; i++)
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

		if (fProfile)
		{
			CloseFile(fProfile);
		}
	}

	OutputToFile(fProfile, L"</profile>");
}