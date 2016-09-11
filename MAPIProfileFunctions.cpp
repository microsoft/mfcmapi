// MAPIProfileFunctions.cpp : Collection of useful MAPI functions

#include "stdafx.h"
#include "MAPIProfileFunctions.h"
#include "ImportProcs.h"
#include "Editor.h"
#include "MAPIFunctions.h"
#include "ExtraPropTags.h"

#ifndef MRMAPI
// This declaration is missing from the MAPI headers
STDAPI STDAPICALLTYPE LaunchWizard(HWND hParentWnd,
	ULONG ulFlags,
	LPCSTR* lppszServiceNameToAdd,
	ULONG cchBufferMax,
	_Out_cap_(cchBufferMax) LPSTR lpszNewProfileName);

wstring LaunchProfileWizard(
	_In_ HWND hParentWnd,
	ULONG ulFlags,
	_In_ string szServiceNameToAdd)
{
	auto hRes = S_OK;
	CHAR szProfName[80] = { 0 };
	ULONG cchProfName = _countof(szProfName);
	LPCSTR szServices[] = { szServiceNameToAdd.c_str(), nullptr };

	DebugPrint(DBGGeneric, L"LaunchProfileWizard: Using LAUNCHWIZARDENTRY to launch wizard API.\n");

	// Call LaunchWizard to add the service.
	WC_MAPI(LaunchWizard(
		hParentWnd,
		ulFlags,
		szServices,
		cchProfName,
		szProfName));
	if (MAPI_E_CALL_FAILED == hRes)
	{
		CHECKHRESMSG(hRes, IDS_LAUNCHWIZARDFAILED);
	}
	else CHECKHRES(hRes);

	if (SUCCEEDED(hRes))
	{
		DebugPrint(DBGGeneric, L"LaunchProfileWizard: Profile \"%hs\" configured.\n", szProfName);
	}

	return LPCSTRToWstring(szProfName);
}

void DisplayMAPISVCPath(_In_ CWnd* pParentWnd)
{
	auto hRes = S_OK;

	DebugPrint(DBGGeneric, L"DisplayMAPISVCPath()\n");

	CEditor MyData(
		pParentWnd,
		IDS_MAPISVCTITLE,
		IDS_MAPISVCTEXT,
		1,
		CEDITOR_BUTTON_OK);
	MyData.InitPane(0, CreateSingleLinePane(IDS_FILEPATH, true));
	MyData.SetStringW(0, GetMAPISVCPath());

	WC_H(MyData.DisplayDialog());
}

// Function name : GetMAPISVCPath
// Description : This will get the correct path to the MAPISVC.INF file.
// Return type : wstring
wstring GetMAPISVCPath()
{
	auto hRes = S_OK;

	auto szMAPIDir = GetMAPIPath(L"Microsoft Outlook"); // STRING_OK

	// We got the path to MAPI - need to strip it
	if (!szMAPIDir.empty())
	{
		szMAPIDir.erase(szMAPIDir.find_last_of(L'\\'), string::npos);
	}
	else
	{
		// Fall back on System32
		UINT cchSystemDir = 0;
		WCHAR szSystemDir[MAX_PATH] = { 0 };
		EC_D(cchSystemDir, GetSystemDirectoryW(szSystemDir, _countof(szSystemDir)));

		if (cchSystemDir < _countof(szSystemDir))
		{
			szMAPIDir = szSystemDir;
		}
	}

	if (SUCCEEDED(hRes) && !szMAPIDir.empty())
	{
		szMAPIDir += L"\\MAPISVC.INF";
	}

	return szMAPIDir;
}

struct SERVICESINIREC
{
	LPWSTR lpszSection;
	LPWSTR lpszKey;
	ULONG ulKey;
	LPWSTR lpszValue;
};

static SERVICESINIREC aEMSServicesIni[] =
{
	{ L"Default Services", L"MSEMS", 0L, L"Microsoft Exchange Server" }, // STRING_OK
	{ L"Services", L"MSEMS", 0L, L"Microsoft Exchange Server" }, // STRING_OK

	{ L"MSEMS", L"PR_DISPLAY_NAME", 0L, L"Microsoft Exchange Server" }, // STRING_OK
	{ L"MSEMS", L"Sections", 0L, L"MSEMS_MSMail_Section" }, // STRING_OK
	{ L"MSEMS", L"PR_SERVICE_DLL_NAME", 0L, L"emsui.dll" }, // STRING_OK
	{ L"MSEMS", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D26A-C2AC-11D1-AD3E-10A0C911C9C0}" }, // STRING_OK
	{ L"MSEMS", L"PR_SERVICE_ENTRY_NAME", 0L, L"EMSCfg" }, // STRING_OK
	{ L"MSEMS", L"PR_RESOURCE_FLAGS", 0L, L"SERVICE_SINGLE_COPY" }, // STRING_OK
	{ L"MSEMS", L"WIZARD_ENTRY_NAME", 0L, L"EMSWizardEntry" }, // STRING_OK
	{ L"MSEMS", L"Providers", 0L, L"EMS_DSA, EMS_MDB_public, EMS_MDB_private, EMS_RXP, EMS_MSX, EMS_Hook" }, // STRING_OK
	{ L"MSEMS", L"PR_SERVICE_SUPPORT_FILES", 0L, L"emsui.dll,emsabp.dll,emsmdb.dll" }, // STRING_OK

	{ L"EMS_MDB_public", L"PR_RESOURCE_TYPE", 0L, L"MAPI_STORE_PROVIDER" }, // STRING_OK
	{ L"EMS_MDB_public", L"PR_PROVIDER_DLL_NAME", 0L, L"EMSMDB.DLL" }, // STRING_OK
	{ L"EMS_MDB_public", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D26A-C2AC-11D1-AD3E-10A0C911C9C0}" }, // STRING_OK
	{ L"EMS_MDB_public", L"PR_RESOURCE_FLAGS", 0L, L"STATUS_NO_DEFAULT_STORE" }, // STRING_OK
	{ L"EMS_MDB_public", nullptr, PR_PROFILE_OPEN_FLAGS, L"06000000" }, // STRING_OK
	{ L"EMS_MDB_public", nullptr, PR_PROFILE_TYPE, L"03000000" }, // STRING_OK
	{ L"EMS_MDB_public", nullptr, PR_MDB_PROVIDER, L"78b2fa70aff711cd9bc800aa002fc45a" }, // STRING_OK
	{ L"EMS_MDB_public", L"PR_DISPLAY_NAME", 0L, L"Public Folders" }, // STRING_OK
	{ L"EMS_MDB_public", L"PR_PROVIDER_DISPLAY", 0L, L"Microsoft Exchange Message Store" }, // STRING_OK

	{ L"EMS_MDB_private", L"PR_PROVIDER_DLL_NAME", 0L, L"EMSMDB.DLL" }, // STRING_OK
	{ L"EMS_MDB_private", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D26A-C2AC-11D1-AD3E-10A0C911C9C0}" }, // STRING_OK
	{ L"EMS_MDB_private", L"PR_RESOURCE_TYPE", 0L, L"MAPI_STORE_PROVIDER" }, // STRING_OK
	{ L"EMS_MDB_private", L"PR_RESOURCE_FLAGS", 0L, L"STATUS_PRIMARY_IDENTITY|STATUS_DEFAULT_STORE|STATUS_PRIMARY_STORE" }, // STRING_OK
	{ L"EMS_MDB_private", nullptr, PR_PROFILE_OPEN_FLAGS, L"0C000000" }, // STRING_OK
	{ L"EMS_MDB_private", nullptr, PR_PROFILE_TYPE, L"01000000" }, // STRING_OK
	{ L"EMS_MDB_private", nullptr, PR_MDB_PROVIDER, L"5494A1C0297F101BA58708002B2A2517" }, // STRING_OK
	{ L"EMS_MDB_private", L"PR_DISPLAY_NAME", 0L, L"Private Folders" }, // STRING_OK
	{ L"EMS_MDB_private", L"PR_PROVIDER_DISPLAY", 0L, L"Microsoft Exchange Message Store" }, // STRING_OK

	{ L"EMS_DSA", L"PR_DISPLAY_NAME", 0L, L"Microsoft Exchange Directory Service" }, // STRING_OK
	{ L"EMS_DSA", L"PR_PROVIDER_DISPLAY", 0L, L"Microsoft Exchange Directory Service" }, // STRING_OK
	{ L"EMS_DSA", L"PR_PROVIDER_DLL_NAME", 0L, L"EMSABP.DLL" }, // STRING_OK
	{ L"EMS_DSA", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D26A-C2AC-11D1-AD3E-10A0C911C9C0}" }, // STRING_OK
	{ L"EMS_DSA", L"PR_RESOURCE_TYPE", 0L, L"MAPI_AB_PROVIDER" }, // STRING_OK

	{ L"MSEMS_MSMail_Section", L"UID", 0L, L"13DBB0C8AA05101A9BB000AA002FC45A" }, // STRING_OK
	{ L"MSEMS_MSMail_Section", nullptr, PR_PROFILE_VERSION, L"01050000" }, // STRING_OK
	{ L"MSEMS_MSMail_Section", nullptr, PR_PROFILE_CONFIG_FLAGS, L"04000000" }, // STRING_OK
	{ L"MSEMS_MSMail_Section", nullptr, PR_PROFILE_TRANSPORT_FLAGS, L"03000000" }, // STRING_OK
	{ L"MSEMS_MSMail_Section", nullptr, PR_PROFILE_CONNECT_FLAGS, L"02000000" }, // STRING_OK

	{ L"EMS_RXP", L"PR_DISPLAY_NAME", 0L, L"Microsoft Exchange Remote Transport" }, // STRING_OK
	{ L"EMS_RXP", L"PR_PROVIDER_DISPLAY", 0L, L"Microsoft Exchange Remote Transport" }, // STRING_OK
	{ L"EMS_RXP", L"PR_PROVIDER_DLL_NAME", 0L, L"EMSUI.DLL" }, // STRING_OK
	{ L"EMS_RXP", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D26A-C2AC-11D1-AD3E-10A0C911C9C0}" }, // STRING_OK
	{ L"EMS_RXP", L"PR_RESOURCE_TYPE", 0L, L"MAPI_TRANSPORT_PROVIDER" }, // STRING_OK
	{ L"EMS_RXP", nullptr, PR_PROFILE_OPEN_FLAGS, L"40000000" }, // STRING_OK
	{ L"EMS_RXP", nullptr, PR_PROFILE_TYPE, L"0A000000" }, // STRING_OK

	{ L"EMS_MSX", L"PR_DISPLAY_NAME", 0L, L"Microsoft Exchange Transport" }, // STRING_OK
	{ L"EMS_MSX", L"PR_PROVIDER_DISPLAY", 0L, L"Microsoft Exchange Transport" }, // STRING_OK
	{ L"EMS_MSX", L"PR_PROVIDER_DLL_NAME", 0L, L"EMSMDB.DLL" }, // STRING_OK
	{ L"EMS_MSX", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D26A-C2AC-11D1-AD3E-10A0C911C9C0}" }, // STRING_OK
	{ L"EMS_MSX", L"PR_RESOURCE_TYPE", 0L, L"MAPI_TRANSPORT_PROVIDER" }, // STRING_OK
	{ L"EMS_MSX", nullptr, PR_PROFILE_OPEN_FLAGS, L"00000000" }, // STRING_OK

	{ L"EMS_Hook", L"PR_DISPLAY_NAME", 0L, L"Microsoft Exchange Hook" }, // STRING_OK
	{ L"EMS_Hook", L"PR_PROVIDER_DISPLAY", 0L, L"Microsoft Exchange Hook" }, // STRING_OK
	{ L"EMS_Hook", L"PR_PROVIDER_DLL_NAME", 0L, L"EMSMDB.DLL" }, // STRING_OK
	{ L"EMS_Hook", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D26A-C2AC-11D1-AD3E-10A0C911C9C0}" }, // STRING_OK
	{ L"EMS_Hook", L"PR_RESOURCE_TYPE", 0L, L"MAPI_HOOK_PROVIDER" }, // STRING_OK
	{ L"EMS_Hook", L"PR_RESOURCE_FLAGS", 0L, L"HOOK_INBOUND" }, // STRING_OK

	{ nullptr, nullptr, 0L, nullptr }
};

// Here's an example of the array to use to remove a service
static SERVICESINIREC aREMOVE_MSEMSServicesIni[] =
{
	{ L"Default Services", L"MSEMS", 0L, nullptr }, // STRING_OK
	{ L"Services", L"MSEMS", 0L, nullptr }, // STRING_OK

	{ L"MSEMS", nullptr, 0L, nullptr }, // STRING_OK

	{ L"EMS_MDB_public", nullptr, 0L, nullptr }, // STRING_OK

	{ L"EMS_MDB_private", nullptr, 0L, nullptr }, // STRING_OK

	{ L"EMS_DSA", nullptr, 0L, nullptr }, // STRING_OK

	{ L"MSEMS_MSMail_Section", nullptr, 0L, nullptr }, // STRING_OK

	{ L"EMS_RXP", nullptr, 0L, nullptr }, // STRING_OK

	{ L"EMS_MSX", nullptr, 0L, nullptr }, // STRING_OK

	{ L"EMS_Hook", nullptr, 0L, nullptr }, // STRING_OK
	{ L"EMSDelegate", nullptr, 0L, nullptr }, // STRING_OK


	{ nullptr, nullptr, 0L, nullptr }
};

static SERVICESINIREC aPSTServicesIni[] =
{
	{ L"Services", L"MSPST MS", 0L, L"Personal Folders File (.pst)" }, // STRING_OK
	{ L"Services", L"MSPST AB", 0L, L"Personal Address Book" }, // STRING_OK

	{ L"MSPST AB", L"PR_DISPLAY_NAME", 0L, L"Personal Address Book" }, // STRING_OK
	{ L"MSPST AB", L"Providers", 0L, L"MSPST ABP" }, // STRING_OK
	{ L"MSPST AB", L"PR_SERVICE_DLL_NAME", 0L, L"MSPST.DLL" }, // STRING_OK
	{ L"MSPST AB", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D262-C2AC-11D1-AD3E-10A0C911C9C0}" }, // STRING_OK
	{ L"MSPST AB", L"PR_SERVICE_SUPPORT_FILES", 0L, L"MSPST.DLL" }, // STRING_OK
	{ L"MSPST AB", L"PR_SERVICE_ENTRY_NAME", 0L, L"PABServiceEntry" }, // STRING_OK
	{ L"MSPST AB", L"PR_RESOURCE_FLAGS", 0L, L"SERVICE_SINGLE_COPY|SERVICE_NO_PRIMARY_IDENTITY" }, // STRING_OK

	{ L"MSPST ABP", L"PR_PROVIDER_DLL_NAME", 0L, L"MSPST.DLL" }, // STRING_OK
	{ L"MSPST ABP", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D262-C2AC-11D1-AD3E-10A0C911C9C0}" }, // STRING_OK
	{ L"MSPST ABP", L"PR_RESOURCE_TYPE", 0L, L"MAPI_AB_PROVIDER" }, // STRING_OK
	{ L"MSPST ABP", L"PR_DISPLAY_NAME", 0L, L"Personal Address Book" }, // STRING_OK
	{ L"MSPST ABP", L"PR_PROVIDER_DISPLAY", 0L, L"Personal Address Book" }, // STRING_OK
	{ L"MSPST ABP", L"PR_SERVICE_DLL_NAME", 0L, L"MSPST.DLL" }, // STRING_OK

	{ L"MSPST MS", L"Providers", 0L, L"MSPST MSP" }, // STRING_OK
	{ L"MSPST MS", L"PR_SERVICE_DLL_NAME", 0L, L"mspst.dll" }, // STRING_OK
	{ L"MSPST MS", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D262-C2AC-11D1-AD3E-10A0C911C9C0}" }, // STRING_OK
	{ L"MSPST MS", L"PR_SERVICE_SUPPORT_FILES", 0L, L"mspst.dll" }, // STRING_OK
	{ L"MSPST MS", L"PR_SERVICE_ENTRY_NAME", 0L, L"PSTServiceEntry" }, // STRING_OK
	{ L"MSPST MS", L"PR_RESOURCE_FLAGS", 0L, L"SERVICE_NO_PRIMARY_IDENTITY" }, // STRING_OK

	{ L"MSPST MSP", nullptr, PR_MDB_PROVIDER, L"4e495441f9bfb80100aa0037d96e0000" }, // STRING_OK
	{ L"MSPST MSP", L"PR_PROVIDER_DLL_NAME", 0L, L"mspst.dll" }, // STRING_OK
	{ L"MSPST MSP", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D262-C2AC-11D1-AD3E-10A0C911C9C0}" }, // STRING_OK
	{ L"MSPST MSP", L"PR_RESOURCE_TYPE", 0L, L"MAPI_STORE_PROVIDER" }, // STRING_OK
	{ L"MSPST MSP", L"PR_RESOURCE_FLAGS", 0L, L"STATUS_DEFAULT_STORE" }, // STRING_OK
	{ L"MSPST MSP", L"PR_DISPLAY_NAME", 0L, L"Personal Folders" }, // STRING_OK
	{ L"MSPST MSP", L"PR_PROVIDER_DISPLAY", 0L, L"Personal Folders File (.pst)" }, // STRING_OK
	{ L"MSPST MSP", L"PR_SERVICE_DLL_NAME", 0L, L"mspst.dll" }, // STRING_OK


	{ nullptr, nullptr, 0L, nullptr }
};

static SERVICESINIREC aREMOVE_MSPSTServicesIni[] =
{
	{ L"Default Services", L"MSPST MS", 0L, nullptr }, // STRING_OK
	{ L"Default Services", L"MSPST AB", 0L, nullptr }, // STRING_OK
	{ L"Services", L"MSPST MS", 0L, nullptr }, // STRING_OK
	{ L"Services", L"MSPST AB", 0L, nullptr }, // STRING_OK

	{ L"MSPST AB", nullptr, 0L, nullptr }, // STRING_OK

	{ L"MSPST ABP", nullptr, 0L, nullptr }, // STRING_OK

	{ L"MSPST MS", nullptr, 0L, nullptr }, // STRING_OK

	{ L"MSPST MSP", nullptr, 0L, nullptr }, // STRING_OK

	{ nullptr, nullptr, 0L, nullptr }
};

// Add values to MAPISVC.INF
_Check_return_ HRESULT HrSetProfileParameters(_In_ SERVICESINIREC *lpServicesIni)
{
	auto hRes = S_OK;

	DebugPrint(DBGGeneric, L"HrSetProfileParameters()\n");

	if (!lpServicesIni) return MAPI_E_INVALID_PARAMETER;

	auto szServicesIni = GetMAPISVCPath();

	if (szServicesIni.empty())
	{
		hRes = MAPI_E_NOT_FOUND;
	}
	else
	{
		DebugPrint(DBGGeneric, L"Writing to this file: \"%ws\"\n", szServicesIni.c_str());

		// Loop through and add items to MAPISVC.INF
		auto n = 0;

		while (lpServicesIni[n].lpszSection != nullptr)
		{
			wstring lpszProp = lpServicesIni[n].lpszKey;
			wstring lpszValue = lpServicesIni[n].lpszValue;

			// Switch the property if necessary
			if (lpszProp.empty() && lpServicesIni[n].ulKey != 0)
			{
				lpszProp = format(
					L"%lx", // STRING_OK
					lpServicesIni[n].ulKey);
			}

			// Write the item to MAPISVC.INF
			DebugPrint(DBGGeneric, L"\tWriting: \"%ws\"::\"%ws\"::\"%ws\"\n",
				lpServicesIni[n].lpszSection,
				lpszProp,
				lpszValue);

			EC_B(WritePrivateProfileStringW(
				lpServicesIni[n].lpszSection,
				lpszProp.c_str(),
				lpszValue.c_str(),
				szServicesIni.c_str()));
			n++;
		}

		// Flush the information - we can ignore the return code
		WritePrivateProfileStringW(nullptr, nullptr, nullptr, szServicesIni.c_str());
	}

	return hRes;
}

void AddServicesToMapiSvcInf()
{
	auto hRes = S_OK;
	CEditor MyData(
		nullptr,
		IDS_ADDSERVICESTOINF,
		IDS_ADDSERVICESTOINFPROMPT,
		2,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, CreateCheckPane(IDS_EXCHANGE, false, false));
	MyData.InitPane(1, CreateCheckPane(IDS_PST, false, false));

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		if (MyData.GetCheck(0))
		{
			EC_H(HrSetProfileParameters(aEMSServicesIni));
			hRes = S_OK;
		}
		if (MyData.GetCheck(1))
		{
			EC_H(HrSetProfileParameters(aPSTServicesIni));
		}
	}
}

void RemoveServicesFromMapiSvcInf()
{
	auto hRes = S_OK;
	CEditor MyData(
		nullptr,
		IDS_REMOVEFROMINF,
		IDS_REMOVEFROMINFPROMPT,
		2,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, CreateCheckPane(IDS_EXCHANGE, false, false));
	MyData.InitPane(1, CreateCheckPane(IDS_PST, false, false));

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		if (MyData.GetCheck(0))
		{
			EC_H(HrSetProfileParameters(aREMOVE_MSEMSServicesIni));
			hRes = S_OK;
		}
		if (MyData.GetCheck(1))
		{
			EC_H(HrSetProfileParameters(aREMOVE_MSPSTServicesIni));
		}
	}
}

#define PR_MARKER PR_BODY_A
#define MARKER_STRING "MFCMAPI Existing Provider Marker" // STRING_OK
// Walk through providers and add/remove our tag
// bAddMark of true will add tag, bAddMark of false will remove it
_Check_return_ HRESULT HrMarkExistingProviders(_In_ LPSERVICEADMIN lpServiceAdmin, bool bAddMark)
{
	auto hRes = S_OK;
	LPMAPITABLE lpProviderTable = nullptr;

	if (!lpServiceAdmin) return MAPI_E_INVALID_PARAMETER;

	static const SizedSPropTagArray(1, pTagUID) =
	{
	1,
	PR_SERVICE_UID
	};

	EC_MAPI(lpServiceAdmin->GetMsgServiceTable(0, &lpProviderTable));

	if (lpProviderTable)
	{
		LPSRowSet lpRowSet = nullptr;

		EC_MAPI(HrQueryAllRows(lpProviderTable, LPSPropTagArray(&pTagUID), NULL, NULL, 0, &lpRowSet));

		if (lpRowSet) DebugPrintSRowSet(DBGGeneric, lpRowSet, NULL);

		if (lpRowSet && lpRowSet->cRows >= 1)
		{
			for (ULONG i = 0; i < lpRowSet->cRows; i++)
			{
				hRes = S_OK;
				auto lpCurRow = &lpRowSet->aRow[i];

				auto lpServiceUID = PpropFindProp(
					lpCurRow->lpProps,
					lpCurRow->cValues,
					PR_SERVICE_UID);

				if (lpServiceUID)
				{
					LPPROFSECT lpSect = nullptr;
					EC_H(OpenProfileSection(
						lpServiceAdmin,
						&lpServiceUID->Value.bin,
						&lpSect));
					if (lpSect)
					{
						if (bAddMark)
						{
							SPropValue PropVal;
							PropVal.ulPropTag = PR_MARKER;
							PropVal.Value.lpszA = MARKER_STRING;
							EC_MAPI(lpSect->SetProps(1, &PropVal, NULL));
						}
						else
						{
							SPropTagArray pTagArray = { 1, PR_MARKER };
							WC_MAPI(lpSect->DeleteProps(&pTagArray, NULL));
						}
						hRes = S_OK;
						EC_MAPI(lpSect->SaveChanges(0));
						lpSect->Release();
					}
				}
			}

		}
		FreeProws(lpRowSet);
		lpProviderTable->Release();
	}
	return hRes;
}

// Returns first provider without our mark on it
_Check_return_ HRESULT HrFindUnmarkedProvider(_In_ LPSERVICEADMIN lpServiceAdmin, _Deref_out_opt_ LPSRowSet* lpRowSet)
{
	auto hRes = S_OK;
	LPMAPITABLE lpProviderTable = nullptr;
	LPPROFSECT lpSect = nullptr;

	if (!lpServiceAdmin || !lpRowSet) return MAPI_E_INVALID_PARAMETER;

	*lpRowSet = nullptr;

	static const SizedSPropTagArray(1, pTagUID) =
	{
	1,
	PR_SERVICE_UID
	};

	EC_MAPI(lpServiceAdmin->GetMsgServiceTable(0, &lpProviderTable));

	if (lpProviderTable)
	{
		EC_MAPI(lpProviderTable->SetColumns(LPSPropTagArray(&pTagUID), TBL_BATCH));
		for (;;)
		{
			EC_MAPI(lpProviderTable->QueryRows(1, 0, lpRowSet));
			if (S_OK == hRes && *lpRowSet && 1 == (*lpRowSet)->cRows)
			{
				auto lpCurRow = &(*lpRowSet)->aRow[0];

				auto lpServiceUID = PpropFindProp(
					lpCurRow->lpProps,
					lpCurRow->cValues,
					PR_SERVICE_UID);

				if (lpServiceUID)
				{
					EC_H(OpenProfileSection(
						lpServiceAdmin,
						&lpServiceUID->Value.bin,
						&lpSect));
					if (lpSect)
					{
						SPropTagArray pTagArray = { 1, PR_MARKER };
						ULONG ulPropVal = 0;
						LPSPropValue lpsPropVal = nullptr;
						EC_H_GETPROPS(lpSect->GetProps(&pTagArray, NULL, &ulPropVal, &lpsPropVal));
						if (!(CheckStringProp(lpsPropVal, PROP_TYPE(PR_MARKER)) &&
							!strcmp(lpsPropVal->Value.lpszA, MARKER_STRING)))
						{
							// got an unmarked provider - this is our hit
							// Don't free *lpRowSet - we're returning it
							hRes = S_OK; // wipe any error from the GetProps - it was expected
							MAPIFreeBuffer(lpsPropVal);
							break;
						}
						MAPIFreeBuffer(lpsPropVal);
						lpSect->Release();
						lpSect = nullptr;
					}
				}
				// go on to next one in the loop
				FreeProws(*lpRowSet);
				*lpRowSet = nullptr;
			}
			else
			{
				// no more hits - get out of the loop
				FreeProws(*lpRowSet);
				*lpRowSet = nullptr;
				break;
			}
		}
		if (lpSect) lpSect->Release();

		lpProviderTable->Release();
	}
	return hRes;
}

_Check_return_ HRESULT HrAddServiceToProfile(
	_In_ string lpszServiceName, // Service Name
	_In_ ULONG_PTR ulUIParam, // hwnd for CreateMsgService
	ULONG ulFlags, // Flags for CreateMsgService
	ULONG cPropVals, // Count of properties for ConfigureMsgService
	_In_opt_ LPSPropValue lpPropVals, // Properties for ConfigureMsgService
	_In_ string lpszProfileName) // profile name
{
	auto hRes = S_OK;
	LPPROFADMIN lpProfAdmin = nullptr;
	LPSERVICEADMIN lpServiceAdmin = nullptr;
	LPSRowSet lpRowSet = nullptr;

	DebugPrint(DBGGeneric, L"HrAddServiceToProfile(%hs,%hs)\n", lpszServiceName.c_str(), lpszProfileName.c_str());

	if (lpszServiceName.empty() || lpszProfileName.empty()) return MAPI_E_INVALID_PARAMETER;

	// Connect to Profile Admin interface.
	EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));
	if (!lpProfAdmin) return hRes;

	EC_MAPI(lpProfAdmin->AdminServices(
		reinterpret_cast<LPTSTR>(const_cast<LPSTR>(lpszProfileName.c_str())),
		LPTSTR(""),
		0,
		0,
		&lpServiceAdmin));

	if (lpServiceAdmin)
	{
		MAPIUID uidService = { 0 };
		auto lpuidService = &uidService;

		LPSERVICEADMIN2 lpServiceAdmin2 = nullptr;
		WC_MAPI(lpServiceAdmin->QueryInterface(IID_IMsgServiceAdmin2, reinterpret_cast<LPVOID*>(&lpServiceAdmin2)));

		if (SUCCEEDED(hRes) && lpServiceAdmin2)
		{
			EC_H_MSG(lpServiceAdmin2->CreateMsgServiceEx(
				reinterpret_cast<LPTSTR>(const_cast<LPSTR>(lpszServiceName.c_str())),
				reinterpret_cast<LPTSTR>(const_cast<LPSTR>(lpszServiceName.c_str())),
				ulUIParam,
				ulFlags,
				&uidService),
				IDS_CREATEMSGSERVICEFAILED);
		}
		else
		{
			hRes = S_OK;
			// Only need to mark if we plan on calling ConfigureMsgService
			if (lpPropVals)
			{
				// Add a dummy prop to the current providers
				EC_H(HrMarkExistingProviders(lpServiceAdmin, true));
			}
			EC_H_MSG(lpServiceAdmin->CreateMsgService(
				reinterpret_cast<LPTSTR>(const_cast<LPSTR>(lpszServiceName.c_str())),
				reinterpret_cast<LPTSTR>(const_cast<LPSTR>(lpszServiceName.c_str())),
				ulUIParam,
				ulFlags),
				IDS_CREATEMSGSERVICEFAILED);
			if (lpPropVals)
			{
				// Look for a provider without our dummy prop
				EC_H(HrFindUnmarkedProvider(lpServiceAdmin, &lpRowSet));

				if (lpRowSet) DebugPrintSRowSet(DBGGeneric, lpRowSet, NULL);

				// should only have one unmarked row
				if (lpRowSet && lpRowSet->cRows == 1)
				{
					auto lpServiceUIDProp = PpropFindProp(
						lpRowSet->aRow[0].lpProps,
						lpRowSet->aRow[0].cValues,
						PR_SERVICE_UID);

					if (lpServiceUIDProp)
					{
						lpuidService = reinterpret_cast<LPMAPIUID>(lpServiceUIDProp->Value.bin.lpb);
					}
				}

				hRes = S_OK;
				// Strip out the dummy prop
				EC_H(HrMarkExistingProviders(lpServiceAdmin, false));
			}
		}

		if (lpPropVals)
		{
			EC_H_CANCEL(lpServiceAdmin->ConfigureMsgService(
				lpuidService,
				NULL,
				0,
				cPropVals,
				lpPropVals));
		}

		FreeProws(lpRowSet);
		if (lpServiceAdmin2) lpServiceAdmin2->Release();
		lpServiceAdmin->Release();
	}

	lpProfAdmin->Release();

	return hRes;
}

_Check_return_ HRESULT HrAddExchangeToProfile(
	_In_ ULONG_PTR ulUIParam, // hwnd for CreateMsgService
	_In_ string& lpszServerName,
	_In_ string& lpszMailboxName,
	_In_ string lpszProfileName)
{
	auto hRes = S_OK;

	DebugPrint(DBGGeneric, L"HrAddExchangeToProfile(%hs,%hs,%hs)\n", lpszServerName.c_str(), lpszMailboxName.c_str(), lpszProfileName.c_str());

	if (lpszServerName.empty() || lpszMailboxName.empty() || lpszProfileName.empty()) return MAPI_E_INVALID_PARAMETER;

#define NUMEXCHANGEPROPS 2
	SPropValue PropVal[NUMEXCHANGEPROPS];
	PropVal[0].ulPropTag = PR_PROFILE_UNRESOLVED_SERVER;
	PropVal[0].Value.lpszA = const_cast<LPSTR>(lpszServerName.c_str());
	PropVal[1].ulPropTag = PR_PROFILE_UNRESOLVED_NAME;
	PropVal[1].Value.lpszA = const_cast<LPSTR>(lpszMailboxName.c_str());
	EC_H(HrAddServiceToProfile("MSEMS", ulUIParam, NULL, NUMEXCHANGEPROPS, PropVal, lpszProfileName)); // STRING_OK

	return hRes;
}

_Check_return_ HRESULT HrAddPSTToProfile(
	_In_ ULONG_PTR ulUIParam, // hwnd for CreateMsgService
	bool bUnicodePST,
	_In_ wstring& lpszPSTPath, // PST name
	_In_ string lpszProfileName, // profile name
	bool bPasswordSet, // whether or not to include a password
	_In_ string& lpszPassword) // password to include
{
	auto hRes = S_OK;

	DebugPrint(DBGGeneric, L"HrAddPSTToProfile(0x%X,%ws,%hs,0x%X,%hs)\n", bUnicodePST, lpszPSTPath.c_str(), lpszProfileName, bPasswordSet, lpszPassword.c_str());

	if (lpszPSTPath.empty() || lpszProfileName.empty()) return MAPI_E_INVALID_PARAMETER;

	SPropValue PropVal[2];
	PropVal[0].ulPropTag = CHANGE_PROP_TYPE(PR_PST_PATH, PT_UNICODE);
	PropVal[0].Value.lpszW = const_cast<LPWSTR>(lpszPSTPath.c_str());
	PropVal[1].ulPropTag = PR_PST_PW_SZ_OLD;
	PropVal[1].Value.lpszA = const_cast<LPSTR>(lpszPassword.c_str());

	if (bUnicodePST)
	{
		EC_H(HrAddServiceToProfile("MSUPST MS", ulUIParam, NULL, bPasswordSet ? 2 : 1, PropVal, lpszProfileName)); // STRING_OK
	}
	else
	{
		EC_H(HrAddServiceToProfile("MSPST MS", ulUIParam, NULL, bPasswordSet ? 2 : 1, PropVal, lpszProfileName)); // STRING_OK
	}

	return hRes;
}

// Creates an empty profile.
_Check_return_ HRESULT HrCreateProfile(
	_In_ string& lpszProfileName) // profile name
{
	auto hRes = S_OK;
	LPPROFADMIN lpProfAdmin = nullptr;

	DebugPrint(DBGGeneric, L"HrCreateProfile(%hs)\n", lpszProfileName.c_str());

	if (lpszProfileName.empty()) return MAPI_E_INVALID_PARAMETER;

	// Connect to Profile Admin interface.
	EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));
	if (!lpProfAdmin) return hRes;

	// Create the profile
	WC_MAPI(lpProfAdmin->CreateProfile(
		reinterpret_cast<LPTSTR>(const_cast<LPSTR>(lpszProfileName.c_str())),
		NULL,
		0,
		NULL)); // fMapiUnicode is not supported!
	if (S_OK != hRes)
	{
		// Did it fail because a profile of this name already exists?
		auto hResCheck = HrMAPIProfileExists(lpProfAdmin, lpszProfileName);
		CHECKHRESMSG(hResCheck, IDS_DUPLICATEPROFILE);
	}

	lpProfAdmin->Release();

	return hRes;
}

// Removes a profile.
_Check_return_ HRESULT HrRemoveProfile(
	_In_ string lpszProfileName)
{
	auto hRes = S_OK;
	LPPROFADMIN lpProfAdmin = nullptr;

	DebugPrint(DBGGeneric, L"HrRemoveProfile(%hs)\n", lpszProfileName.c_str());
	if (lpszProfileName.empty()) return MAPI_E_INVALID_PARAMETER;

	EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));
	if (!lpProfAdmin) return hRes;

	EC_MAPI(lpProfAdmin->DeleteProfile(reinterpret_cast<LPTSTR>(const_cast<LPSTR>(lpszProfileName.c_str())), 0));

	lpProfAdmin->Release();

	RegFlushKey(HKEY_LOCAL_MACHINE);
	RegFlushKey(HKEY_CURRENT_USER);

	return hRes;
}

// Set a profile as default.
_Check_return_ HRESULT HrSetDefaultProfile(
	_In_ string lpszProfileName)
{
	auto hRes = S_OK;
	LPPROFADMIN lpProfAdmin = nullptr;

	DebugPrint(DBGGeneric, L"HrRemoveProfile(%hs)\n", lpszProfileName.c_str());
	if (lpszProfileName.empty()) return MAPI_E_INVALID_PARAMETER;

	EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));
	if (!lpProfAdmin) return hRes;

	EC_MAPI(lpProfAdmin->SetDefaultProfile(reinterpret_cast<LPTSTR>(const_cast<LPSTR>(lpszProfileName.c_str())), 0));

	lpProfAdmin->Release();

	RegFlushKey(HKEY_LOCAL_MACHINE);
	RegFlushKey(HKEY_CURRENT_USER);

	return hRes;
}

// Checks for an existing profile.
_Check_return_ HRESULT HrMAPIProfileExists(
	_In_ LPPROFADMIN lpProfAdmin,
	_In_ string& lpszProfileName)
{
	auto hRes = S_OK;
	LPMAPITABLE lpTable = nullptr;
	LPSRowSet lpRows = nullptr;
	LPSPropValue lpProp = nullptr;

	static const SizedSPropTagArray(1, rgPropTag) =
	{
	1,
	PR_DISPLAY_NAME_A
	};

	DebugPrint(DBGGeneric, L"HrMAPIProfileExists()\n");
	if (!lpProfAdmin || lpszProfileName.empty()) return MAPI_E_INVALID_PARAMETER;

	// Get a table of existing profiles

	EC_MAPI(lpProfAdmin->GetProfileTable(
		0,
		&lpTable));
	if (!lpTable) return hRes;

	EC_MAPI(HrQueryAllRows(
		lpTable,
		LPSPropTagArray(&rgPropTag),
		NULL,
		NULL,
		0,
		&lpRows));

	if (lpRows)
	{
		if (lpRows->cRows == 0)
		{
			// If table is empty then profile doesn't exist
			hRes = S_OK;
		}
		else
		{
			// Search rows for the folder in question

			if (!FAILED(hRes)) for (ULONG i = 0; i < lpRows->cRows; i++)
			{
				hRes = S_OK;
				lpProp = lpRows->aRow[i].lpProps;

				ULONG ulComp = NULL;

				EC_D(ulComp, CompareStringA(
					g_lcid, // LOCALE_INVARIANT,
					NORM_IGNORECASE,
					lpProp[0].Value.lpszA,
					-1,
					lpszProfileName.c_str(),
					-1));

				if (CSTR_EQUAL == ulComp)
				{
					hRes = E_ACCESSDENIED;
					break;
				}
			}
		}
	}

	if (lpRows) FreeProws(lpRows);

	lpTable->Release();
	return hRes;
}

_Check_return_ HRESULT GetProfileServiceVersion(
	_In_ string lpszProfileName,
	_Out_ ULONG* lpulServerVersion,
	_Out_ EXCHANGE_STORE_VERSION_NUM* lpStoreVersion,
	_Out_ bool* lpbFoundServerVersion,
	_Out_ bool* lpbFoundServerFullVersion)
{
	if (lpszProfileName.empty()
		|| !lpulServerVersion
		|| !lpStoreVersion
		|| !lpbFoundServerVersion
		|| !lpbFoundServerFullVersion) return MAPI_E_INVALID_PARAMETER;
	*lpulServerVersion = NULL;
	memset(lpStoreVersion, 0, sizeof(EXCHANGE_STORE_VERSION_NUM));
	*lpbFoundServerVersion = false;
	*lpbFoundServerFullVersion = false;

	auto hRes = S_OK;
	LPPROFADMIN lpProfAdmin = nullptr;
	LPSERVICEADMIN lpServiceAdmin = nullptr;

	DebugPrint(DBGGeneric, L"GetProfileServiceVersion(%hs)\n", lpszProfileName.c_str());

	EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));
	if (!lpProfAdmin) return hRes;

	EC_MAPI(lpProfAdmin->AdminServices(
		reinterpret_cast<LPTSTR>(const_cast<LPSTR>(lpszProfileName.c_str())),
		LPTSTR(""),
		0,
		0,
		&lpServiceAdmin));

	if (lpServiceAdmin)
	{
		LPPROFSECT lpProfSect = nullptr;
		EC_MAPI(lpServiceAdmin->OpenProfileSection(
			LPMAPIUID(pbGlobalProfileSectionGuid),
			NULL,
			0,
			&lpProfSect));
		if (lpProfSect)
		{
			LPSPropValue lpServerVersion = nullptr;
			WC_MAPI(HrGetOneProp(lpProfSect, PR_PROFILE_SERVER_VERSION, &lpServerVersion));

			if (SUCCEEDED(hRes) && lpServerVersion && PR_PROFILE_SERVER_VERSION == lpServerVersion->ulPropTag)
			{
				*lpbFoundServerVersion = true;
				*lpulServerVersion = lpServerVersion->Value.l;
			}
			MAPIFreeBuffer(lpServerVersion);
			hRes = S_OK;

			LPSPropValue lpServerFullVersion = nullptr;
			WC_MAPI(HrGetOneProp(lpProfSect, PR_PROFILE_SERVER_FULL_VERSION, &lpServerFullVersion));

			if (SUCCEEDED(hRes) &&
				lpServerFullVersion &&
				PR_PROFILE_SERVER_FULL_VERSION == lpServerFullVersion->ulPropTag &&
				sizeof(EXCHANGE_STORE_VERSION_NUM) == lpServerFullVersion->Value.bin.cb)
			{
				DebugPrint(DBGGeneric, L"PR_PROFILE_SERVER_FULL_VERSION = ");
				DebugPrintBinary(DBGGeneric, &lpServerFullVersion->Value.bin);
				DebugPrint(DBGGeneric, L"\n");

				memcpy(lpStoreVersion, lpServerFullVersion->Value.bin.lpb, sizeof(EXCHANGE_STORE_VERSION_NUM));
				*lpbFoundServerFullVersion = true;
			}
			MAPIFreeBuffer(lpServerFullVersion);

			lpProfSect->Release();
		}

		lpServiceAdmin->Release();
	}

	lpProfAdmin->Release();

	// If we found any server version, consider the call a success
	if (*lpbFoundServerVersion || *lpbFoundServerFullVersion) hRes = S_OK;

	return hRes;
}

// Copies a profile.
_Check_return_ HRESULT HrCopyProfile(
	_In_ string lpszOldProfileName,
	_In_ string lpszNewProfileName)
{
	auto hRes = S_OK;
	LPPROFADMIN lpProfAdmin = nullptr;

	DebugPrint(DBGGeneric, L"HrCopyProfile(%hs, %hs)\n", lpszOldProfileName.c_str(), lpszNewProfileName.c_str());
	if (lpszOldProfileName.empty() || lpszNewProfileName.empty()) return MAPI_E_INVALID_PARAMETER;

	EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));
	if (!lpProfAdmin) return hRes;

	EC_MAPI(lpProfAdmin->CopyProfile(
		reinterpret_cast<LPTSTR>(const_cast<LPSTR>(lpszOldProfileName.c_str())),
		NULL,
		reinterpret_cast<LPTSTR>(const_cast<LPSTR>(lpszNewProfileName.c_str())),
		NULL,
		NULL));

	lpProfAdmin->Release();

	return hRes;
}
#endif

#define MAPI_FORCE_ACCESS 0x00080000

_Check_return_ HRESULT OpenProfileSection(_In_ LPSERVICEADMIN lpServiceAdmin, _In_ LPSBinary lpServiceUID, _Deref_out_opt_ LPPROFSECT* lppProfSect)
{
	auto hRes = S_OK;

	DebugPrint(DBGOpenItemProp, L"OpenProfileSection opening lpServiceUID = ");
	DebugPrintBinary(DBGOpenItemProp, lpServiceUID);
	DebugPrint(DBGOpenItemProp, L"\n");

	if (!lpServiceUID || !lpServiceAdmin || !lppProfSect) return MAPI_E_INVALID_PARAMETER;
	*lppProfSect = nullptr;

	// First, we try the normal way of opening the profile section:
	WC_MAPI(lpServiceAdmin->OpenProfileSection(
		reinterpret_cast<LPMAPIUID>(lpServiceUID->lpb),
		NULL,
		MAPI_MODIFY | MAPI_FORCE_ACCESS, // passing this flag might actually work with Outlook 2000 and XP
		lppProfSect));

	if (!*lppProfSect)
	{
		hRes = S_OK;
		///////////////////////////////////////////////////////////////////
		// HACK CENTRAL. This is a MAJOR hack. MAPI will always return E_ACCESSDENIED
		// when we open a profile section on the service if we are a client. The workaround
		// (HACK) is to call into one of MAPI's internal functions that bypasses
		// the security check. We build a Interface to it and then point to it from our
		// offset of 0x48. USE AT YOUR OWN RISK! NOT SUPPORTED!
		interface IOpenSectionHack : public IUnknown
		{
		public:
			virtual HRESULT STDMETHODCALLTYPE OpenSection(LPMAPIUID, ULONG, LPPROFSECT*) = 0;
		};

		IOpenSectionHack** ppProfile;

		ppProfile = reinterpret_cast<IOpenSectionHack**>((reinterpret_cast<BYTE*>(lpServiceAdmin) + 0x48));

		// Now, we want to get open the Services Profile Section and store that
		// interface with the Object

		if (ppProfile && *ppProfile)
		{
			EC_MAPI((*ppProfile)->OpenSection(
				reinterpret_cast<LPMAPIUID>(lpServiceUID->lpb),
				MAPI_MODIFY,
				lppProfSect));
		}
		else
		{
			hRes = MAPI_E_NOT_FOUND;
		}
		// END OF HACK. I'm amazed that this works....
		///////////////////////////////////////////////////////////////////
	}
	return hRes;
}

_Check_return_ HRESULT OpenProfileSection(_In_ LPPROVIDERADMIN lpProviderAdmin, _In_ LPSBinary lpProviderUID, _Deref_out_ LPPROFSECT* lppProfSect)
{
	auto hRes = S_OK;

	DebugPrint(DBGOpenItemProp, L"OpenProfileSection opening lpServiceUID = ");
	DebugPrintBinary(DBGOpenItemProp, lpProviderUID);
	DebugPrint(DBGOpenItemProp, L"\n");

	if (!lpProviderUID || !lpProviderAdmin || !lppProfSect) return MAPI_E_INVALID_PARAMETER;
	*lppProfSect = nullptr;

	WC_MAPI(lpProviderAdmin->OpenProfileSection(
		reinterpret_cast<LPMAPIUID>(lpProviderUID->lpb),
		NULL,
		MAPI_MODIFY | MAPI_FORCE_ACCESS,
		lppProfSect));
	if (!*lppProfSect)
	{
		hRes = S_OK;

		// We only do this hack as a last resort - it can crash some versions of Outlook, but is required for Exchange
		*(reinterpret_cast<BYTE*>(lpProviderAdmin) + 0x60) = 0x2; // Use at your own risk! NOT SUPPORTED!

		WC_MAPI(lpProviderAdmin->OpenProfileSection(
			reinterpret_cast<LPMAPIUID>(lpProviderUID->lpb),
			NULL,
			MAPI_MODIFY,
			lppProfSect));
	}
	return hRes;
}