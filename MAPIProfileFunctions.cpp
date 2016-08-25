// MAPIProfileFunctions.cpp : Collection of useful MAPI functions

#include "stdafx.h"
#include "MAPIProfileFunctions.h"
#include "ImportProcs.h"
#include "Editor.h"
#include "MAPIFunctions.h"
#include "ExtraPropTags.h"
#include "Guids.h"

#ifndef MRMAPI
// This declaration is missing from the MAPI headers
STDAPI STDAPICALLTYPE LaunchWizard(HWND hParentWnd,
	ULONG ulFlags,
	LPCSTR* lppszServiceNameToAdd,
	ULONG cchBufferMax,
	_Out_cap_(cchBufferMax) LPSTR lpszNewProfileName);

void LaunchProfileWizard(
	_In_ HWND hParentWnd,
	ULONG ulFlags,
	_In_z_ LPCSTR* lppszServiceNameToAdd,
	ULONG cchBufferMax,
	_Out_cap_(cchBufferMax) LPSTR lpszNewProfileName)
{
	HRESULT hRes = S_OK;

	DebugPrint(DBGGeneric, L"LaunchProfileWizard: Using LAUNCHWIZARDENTRY to launch wizard API.\n");

	// Call LaunchWizard to add the service.
	WC_MAPI(LaunchWizard(
		hParentWnd,
		ulFlags,
		lppszServiceNameToAdd,
		cchBufferMax,
		lpszNewProfileName));
	if (MAPI_E_CALL_FAILED == hRes)
	{
		CHECKHRESMSG(hRes, IDS_LAUNCHWIZARDFAILED);
	}
	else CHECKHRES(hRes);

	if (SUCCEEDED(hRes))
	{
		DebugPrint(DBGGeneric, L"LaunchProfileWizard: Profile \"%hs\" configured.\n", lpszNewProfileName);
	}
} // LaunchProfileWizard

void DisplayMAPISVCPath(_In_ CWnd* pParentWnd)
{
	HRESULT hRes = S_OK;
	TCHAR   szServicesIni[MAX_PATH + 12] = { 0 }; // 12 = space for 'MAPISVC.INF'

	DebugPrint(DBGGeneric, L"DisplayMAPISVCPath()\n");

	GetMAPISVCPath(szServicesIni, _countof(szServicesIni));

	CEditor MyData(
		pParentWnd,
		IDS_MAPISVCTITLE,
		IDS_MAPISVCTEXT,
		1,
		CEDITOR_BUTTON_OK);
	MyData.InitPane(0, CreateSingleLinePane(IDS_FILEPATH, true));
	MyData.SetString(0, szServicesIni);

	WC_H(MyData.DisplayDialog());
} // DisplayMAPISVCPath

///////////////////////////////////////////////////////////////////////////////
// Function name : GetMAPISVCPath
// Description   : This will get the correct path to the MAPISVC.INF file.
// Return type   : void
// Argument      : LPTSTR szMAPIDir - Buffer to hold the path to the MAPISVC file.
//                 ULONG cchMAPIDir - size of the buffer
void GetMAPISVCPath(_Inout_z_count_(cchMAPIDir) LPTSTR szMAPIDir, ULONG cchMAPIDir)
{
	HRESULT hRes = S_OK;

	GetMAPIPath(_T("Microsoft Outlook"), szMAPIDir, cchMAPIDir); // STRING_OK

	// We got the path to MAPI - need to strip it
	if (szMAPIDir[0])
	{
		LPTSTR lpszSlash = NULL;
		LPTSTR lpszCur = szMAPIDir;

		for (lpszSlash = lpszCur; *lpszCur; lpszCur = lpszCur++)
		{
			if (*lpszCur == _T('\\')) lpszSlash = lpszCur;
		}
		*lpszSlash = _T('\0');
	}
	else
	{
		// Fall back on System32
		UINT cchSystemDir = 0;
		TCHAR szSystemDir[MAX_PATH] = { 0 };
		EC_D(cchSystemDir, GetSystemDirectory(szSystemDir, _countof(szSystemDir)));

		if (cchSystemDir < _countof(szSystemDir))
		{
			EC_H(StringCchCopy(szMAPIDir, cchMAPIDir, szSystemDir));
		}
	}

	if (SUCCEEDED(hRes) && szMAPIDir[0])
	{
		EC_H(StringCchPrintf(
			szMAPIDir,
			cchMAPIDir,
			_T("%s\\%s"), // STRING_OK
			szMAPIDir,
			_T("MAPISVC.INF"))); // STRING_OK
	}
} // GetMAPISVCPath

struct SERVICESINIREC
{
	LPTSTR lpszSection;
	LPTSTR lpszKey;
	ULONG ulKey;
	LPTSTR lpszValue;
};

static SERVICESINIREC aEMSServicesIni[] =
{
	{ _T("Default Services"), _T("MSEMS"), 0L, _T("Microsoft Exchange Server") }, // STRING_OK
	{ _T("Services"), _T("MSEMS"), 0L, _T("Microsoft Exchange Server") }, // STRING_OK

	{ _T("MSEMS"), _T("PR_DISPLAY_NAME"), 0L, _T("Microsoft Exchange Server") }, // STRING_OK
	{ _T("MSEMS"), _T("Sections"), 0L, _T("MSEMS_MSMail_Section") }, // STRING_OK
	{ _T("MSEMS"), _T("PR_SERVICE_DLL_NAME"), 0L, _T("emsui.dll") }, // STRING_OK
	{ _T("MSEMS"), _T("PR_SERVICE_INSTALL_ID"), 0L, _T("{6485D26A-C2AC-11D1-AD3E-10A0C911C9C0}") }, // STRING_OK
	{ _T("MSEMS"), _T("PR_SERVICE_ENTRY_NAME"), 0L, _T("EMSCfg") }, // STRING_OK
	{ _T("MSEMS"), _T("PR_RESOURCE_FLAGS"), 0L, _T("SERVICE_SINGLE_COPY") }, // STRING_OK
	{ _T("MSEMS"), _T("WIZARD_ENTRY_NAME"), 0L, _T("EMSWizardEntry") }, // STRING_OK
	{ _T("MSEMS"), _T("Providers"), 0L, _T("EMS_DSA, EMS_MDB_public, EMS_MDB_private, EMS_RXP, EMS_MSX, EMS_Hook") }, // STRING_OK
	{ _T("MSEMS"), _T("PR_SERVICE_SUPPORT_FILES"), 0L, _T("emsui.dll,emsabp.dll,emsmdb.dll") }, // STRING_OK

	{ _T("EMS_MDB_public"), _T("PR_RESOURCE_TYPE"), 0L, _T("MAPI_STORE_PROVIDER") }, // STRING_OK
	{ _T("EMS_MDB_public"), _T("PR_PROVIDER_DLL_NAME"), 0L, _T("EMSMDB.DLL") }, // STRING_OK
	{ _T("EMS_MDB_public"), _T("PR_SERVICE_INSTALL_ID"), 0L, _T("{6485D26A-C2AC-11D1-AD3E-10A0C911C9C0}") }, // STRING_OK
	{ _T("EMS_MDB_public"), _T("PR_RESOURCE_FLAGS"), 0L, _T("STATUS_NO_DEFAULT_STORE") }, // STRING_OK
	{ _T("EMS_MDB_public"), NULL, PR_PROFILE_OPEN_FLAGS, _T("06000000") }, // STRING_OK
	{ _T("EMS_MDB_public"), NULL, PR_PROFILE_TYPE, _T("03000000") }, // STRING_OK
	{ _T("EMS_MDB_public"), NULL, PR_MDB_PROVIDER, _T("78b2fa70aff711cd9bc800aa002fc45a") }, // STRING_OK
	{ _T("EMS_MDB_public"), _T("PR_DISPLAY_NAME"), 0L, _T("Public Folders") }, // STRING_OK
	{ _T("EMS_MDB_public"), _T("PR_PROVIDER_DISPLAY"), 0L, _T("Microsoft Exchange Message Store") }, // STRING_OK

	{ _T("EMS_MDB_private"), _T("PR_PROVIDER_DLL_NAME"), 0L, _T("EMSMDB.DLL") }, // STRING_OK
	{ _T("EMS_MDB_private"), _T("PR_SERVICE_INSTALL_ID"), 0L, _T("{6485D26A-C2AC-11D1-AD3E-10A0C911C9C0}") }, // STRING_OK
	{ _T("EMS_MDB_private"), _T("PR_RESOURCE_TYPE"), 0L, _T("MAPI_STORE_PROVIDER") }, // STRING_OK
	{ _T("EMS_MDB_private"), _T("PR_RESOURCE_FLAGS"), 0L, _T("STATUS_PRIMARY_IDENTITY|STATUS_DEFAULT_STORE|STATUS_PRIMARY_STORE") }, // STRING_OK
	{ _T("EMS_MDB_private"), NULL, PR_PROFILE_OPEN_FLAGS, _T("0C000000") }, // STRING_OK
	{ _T("EMS_MDB_private"), NULL, PR_PROFILE_TYPE, _T("01000000") }, // STRING_OK
	{ _T("EMS_MDB_private"), NULL, PR_MDB_PROVIDER, _T("5494A1C0297F101BA58708002B2A2517") }, // STRING_OK
	{ _T("EMS_MDB_private"), _T("PR_DISPLAY_NAME"), 0L, _T("Private Folders") }, // STRING_OK
	{ _T("EMS_MDB_private"), _T("PR_PROVIDER_DISPLAY"), 0L, _T("Microsoft Exchange Message Store") }, // STRING_OK

	{ _T("EMS_DSA"), _T("PR_DISPLAY_NAME"), 0L, _T("Microsoft Exchange Directory Service") }, // STRING_OK
	{ _T("EMS_DSA"), _T("PR_PROVIDER_DISPLAY"), 0L, _T("Microsoft Exchange Directory Service") }, // STRING_OK
	{ _T("EMS_DSA"), _T("PR_PROVIDER_DLL_NAME"), 0L, _T("EMSABP.DLL") }, // STRING_OK
	{ _T("EMS_DSA"), _T("PR_SERVICE_INSTALL_ID"), 0L, _T("{6485D26A-C2AC-11D1-AD3E-10A0C911C9C0}") }, // STRING_OK
	{ _T("EMS_DSA"), _T("PR_RESOURCE_TYPE"), 0L, _T("MAPI_AB_PROVIDER") }, // STRING_OK

	{ _T("MSEMS_MSMail_Section"), _T("UID"), 0L, _T("13DBB0C8AA05101A9BB000AA002FC45A") }, // STRING_OK
	{ _T("MSEMS_MSMail_Section"), NULL, PR_PROFILE_VERSION, _T("01050000") }, // STRING_OK
	{ _T("MSEMS_MSMail_Section"), NULL, PR_PROFILE_CONFIG_FLAGS, _T("04000000") }, // STRING_OK
	{ _T("MSEMS_MSMail_Section"), NULL, PR_PROFILE_TRANSPORT_FLAGS, _T("03000000") }, // STRING_OK
	{ _T("MSEMS_MSMail_Section"), NULL, PR_PROFILE_CONNECT_FLAGS, _T("02000000") }, // STRING_OK

	{ _T("EMS_RXP"), _T("PR_DISPLAY_NAME"), 0L, _T("Microsoft Exchange Remote Transport") }, // STRING_OK
	{ _T("EMS_RXP"), _T("PR_PROVIDER_DISPLAY"), 0L, _T("Microsoft Exchange Remote Transport") }, // STRING_OK
	{ _T("EMS_RXP"), _T("PR_PROVIDER_DLL_NAME"), 0L, _T("EMSUI.DLL") }, // STRING_OK
	{ _T("EMS_RXP"), _T("PR_SERVICE_INSTALL_ID"), 0L, _T("{6485D26A-C2AC-11D1-AD3E-10A0C911C9C0}") }, // STRING_OK
	{ _T("EMS_RXP"), _T("PR_RESOURCE_TYPE"), 0L, _T("MAPI_TRANSPORT_PROVIDER") }, // STRING_OK
	{ _T("EMS_RXP"), NULL, PR_PROFILE_OPEN_FLAGS, _T("40000000") }, // STRING_OK
	{ _T("EMS_RXP"), NULL, PR_PROFILE_TYPE, _T("0A000000") }, // STRING_OK

	{ _T("EMS_MSX"), _T("PR_DISPLAY_NAME"), 0L, _T("Microsoft Exchange Transport") }, // STRING_OK
	{ _T("EMS_MSX"), _T("PR_PROVIDER_DISPLAY"), 0L, _T("Microsoft Exchange Transport") }, // STRING_OK
	{ _T("EMS_MSX"), _T("PR_PROVIDER_DLL_NAME"), 0L, _T("EMSMDB.DLL") }, // STRING_OK
	{ _T("EMS_MSX"), _T("PR_SERVICE_INSTALL_ID"), 0L, _T("{6485D26A-C2AC-11D1-AD3E-10A0C911C9C0}") }, // STRING_OK
	{ _T("EMS_MSX"), _T("PR_RESOURCE_TYPE"), 0L, _T("MAPI_TRANSPORT_PROVIDER") }, // STRING_OK
	{ _T("EMS_MSX"), NULL, PR_PROFILE_OPEN_FLAGS, _T("00000000") }, // STRING_OK

	{ _T("EMS_Hook"), _T("PR_DISPLAY_NAME"), 0L, _T("Microsoft Exchange Hook") }, // STRING_OK
	{ _T("EMS_Hook"), _T("PR_PROVIDER_DISPLAY"), 0L, _T("Microsoft Exchange Hook") }, // STRING_OK
	{ _T("EMS_Hook"), _T("PR_PROVIDER_DLL_NAME"), 0L, _T("EMSMDB.DLL") }, // STRING_OK
	{ _T("EMS_Hook"), _T("PR_SERVICE_INSTALL_ID"), 0L, _T("{6485D26A-C2AC-11D1-AD3E-10A0C911C9C0}") }, // STRING_OK
	{ _T("EMS_Hook"), _T("PR_RESOURCE_TYPE"), 0L, _T("MAPI_HOOK_PROVIDER") }, // STRING_OK
	{ _T("EMS_Hook"), _T("PR_RESOURCE_FLAGS"), 0L, _T("HOOK_INBOUND") }, // STRING_OK

	{ NULL, NULL, 0L, NULL }
};

// Here's an example of the array to use to remove a service
static SERVICESINIREC aREMOVE_MSEMSServicesIni[] =
{
	{ _T("Default Services"), _T("MSEMS"), 0L, NULL }, // STRING_OK
	{ _T("Services"), _T("MSEMS"), 0L, NULL }, // STRING_OK

	{ _T("MSEMS"), NULL, 0L, NULL }, // STRING_OK

	{ _T("EMS_MDB_public"), NULL, 0L, NULL }, // STRING_OK

	{ _T("EMS_MDB_private"), NULL, 0L, NULL }, // STRING_OK

	{ _T("EMS_DSA"), NULL, 0L, NULL }, // STRING_OK

	{ _T("MSEMS_MSMail_Section"), NULL, 0L, NULL }, // STRING_OK

	{ _T("EMS_RXP"), NULL, 0L, NULL }, // STRING_OK

	{ _T("EMS_MSX"), NULL, 0L, NULL }, // STRING_OK

	{ _T("EMS_Hook"), NULL, 0L, NULL }, // STRING_OK
	{ _T("EMSDelegate"), NULL, 0L, NULL }, // STRING_OK


	{ NULL, NULL, 0L, NULL }
};

static SERVICESINIREC aPSTServicesIni[] =
{
	{ _T("Services"), _T("MSPST MS"), 0L, _T("Personal Folders File (.pst)") }, // STRING_OK
	{ _T("Services"), _T("MSPST AB"), 0L, _T("Personal Address Book") }, // STRING_OK

	{ _T("MSPST AB"), _T("PR_DISPLAY_NAME"), 0L, _T("Personal Address Book") }, // STRING_OK
	{ _T("MSPST AB"), _T("Providers"), 0L, _T("MSPST ABP") }, // STRING_OK
	{ _T("MSPST AB"), _T("PR_SERVICE_DLL_NAME"), 0L, _T("MSPST.DLL") }, // STRING_OK
	{ _T("MSPST AB"), _T("PR_SERVICE_INSTALL_ID"), 0L, _T("{6485D262-C2AC-11D1-AD3E-10A0C911C9C0}") }, // STRING_OK
	{ _T("MSPST AB"), _T("PR_SERVICE_SUPPORT_FILES"), 0L, _T("MSPST.DLL") }, // STRING_OK
	{ _T("MSPST AB"), _T("PR_SERVICE_ENTRY_NAME"), 0L, _T("PABServiceEntry") }, // STRING_OK
	{ _T("MSPST AB"), _T("PR_RESOURCE_FLAGS"), 0L, _T("SERVICE_SINGLE_COPY|SERVICE_NO_PRIMARY_IDENTITY") }, // STRING_OK

	{ _T("MSPST ABP"), _T("PR_PROVIDER_DLL_NAME"), 0L, _T("MSPST.DLL") }, // STRING_OK
	{ _T("MSPST ABP"), _T("PR_SERVICE_INSTALL_ID"), 0L, _T("{6485D262-C2AC-11D1-AD3E-10A0C911C9C0}") }, // STRING_OK
	{ _T("MSPST ABP"), _T("PR_RESOURCE_TYPE"), 0L, _T("MAPI_AB_PROVIDER") }, // STRING_OK
	{ _T("MSPST ABP"), _T("PR_DISPLAY_NAME"), 0L, _T("Personal Address Book") }, // STRING_OK
	{ _T("MSPST ABP"), _T("PR_PROVIDER_DISPLAY"), 0L, _T("Personal Address Book") }, // STRING_OK
	{ _T("MSPST ABP"), _T("PR_SERVICE_DLL_NAME"), 0L, _T("MSPST.DLL") }, // STRING_OK

	{ _T("MSPST MS"), _T("Providers"), 0L, _T("MSPST MSP") }, // STRING_OK
	{ _T("MSPST MS"), _T("PR_SERVICE_DLL_NAME"), 0L, _T("mspst.dll") }, // STRING_OK
	{ _T("MSPST MS"), _T("PR_SERVICE_INSTALL_ID"), 0L, _T("{6485D262-C2AC-11D1-AD3E-10A0C911C9C0}") }, // STRING_OK
	{ _T("MSPST MS"), _T("PR_SERVICE_SUPPORT_FILES"), 0L, _T("mspst.dll") }, // STRING_OK
	{ _T("MSPST MS"), _T("PR_SERVICE_ENTRY_NAME"), 0L, _T("PSTServiceEntry") }, // STRING_OK
	{ _T("MSPST MS"), _T("PR_RESOURCE_FLAGS"), 0L, _T("SERVICE_NO_PRIMARY_IDENTITY") }, // STRING_OK

	{ _T("MSPST MSP"), NULL, PR_MDB_PROVIDER, _T("4e495441f9bfb80100aa0037d96e0000") }, // STRING_OK
	{ _T("MSPST MSP"), _T("PR_PROVIDER_DLL_NAME"), 0L, _T("mspst.dll") }, // STRING_OK
	{ _T("MSPST MSP"), _T("PR_SERVICE_INSTALL_ID"), 0L, _T("{6485D262-C2AC-11D1-AD3E-10A0C911C9C0}") }, // STRING_OK
	{ _T("MSPST MSP"), _T("PR_RESOURCE_TYPE"), 0L, _T("MAPI_STORE_PROVIDER") }, // STRING_OK
	{ _T("MSPST MSP"), _T("PR_RESOURCE_FLAGS"), 0L, _T("STATUS_DEFAULT_STORE") }, // STRING_OK
	{ _T("MSPST MSP"), _T("PR_DISPLAY_NAME"), 0L, _T("Personal Folders") }, // STRING_OK
	{ _T("MSPST MSP"), _T("PR_PROVIDER_DISPLAY"), 0L, _T("Personal Folders File (.pst)") }, // STRING_OK
	{ _T("MSPST MSP"), _T("PR_SERVICE_DLL_NAME"), 0L, _T("mspst.dll") }, // STRING_OK


	{ NULL, NULL, 0L, NULL }
};

static SERVICESINIREC aREMOVE_MSPSTServicesIni[] =
{
	{ _T("Default Services"), _T("MSPST MS"), 0L, NULL }, // STRING_OK
	{ _T("Default Services"), _T("MSPST AB"), 0L, NULL }, // STRING_OK
	{ _T("Services"), _T("MSPST MS"), 0L, NULL }, // STRING_OK
	{ _T("Services"), _T("MSPST AB"), 0L, NULL }, // STRING_OK

	{ _T("MSPST AB"), NULL, 0L, NULL }, // STRING_OK

	{ _T("MSPST ABP"), NULL, 0L, NULL }, // STRING_OK

	{ _T("MSPST MS"), NULL, 0L, NULL }, // STRING_OK

	{ _T("MSPST MSP"), NULL, 0L, NULL }, // STRING_OK

	{ NULL, NULL, 0L, NULL }
};

// $--HrSetProfileParameters----------------------------------------------
// Add values to MAPISVC.INF
// -----------------------------------------------------------------------------
_Check_return_ HRESULT HrSetProfileParameters(_In_ SERVICESINIREC *lpServicesIni)
{
	HRESULT	hRes = S_OK;
	TCHAR	szServicesIni[MAX_PATH + 12] = { 0 }; // 12 = space for 'MAPISVC.INF'
	UINT	n = 0;
	TCHAR	szPropNum[10] = { 0 };

	DebugPrint(DBGGeneric, L"HrSetProfileParameters()\n");

	if (!lpServicesIni) return MAPI_E_INVALID_PARAMETER;

	GetMAPISVCPath(szServicesIni, _countof(szServicesIni));

	if (!szServicesIni[0])
	{
		hRes = MAPI_E_NOT_FOUND;
	}
	else
	{
		DebugPrint(DBGGeneric, L"Writing to this file: \"%ws\"\n", LPCTSTRToWstring(szServicesIni).c_str());

		//
		// Loop through and add items to MAPISVC.INF
		//

		n = 0;

		while (lpServicesIni[n].lpszSection != NULL)
		{
			LPTSTR lpszProp = lpServicesIni[n].lpszKey;
			LPTSTR lpszValue = lpServicesIni[n].lpszValue;

			// Switch the property if necessary

			if ((lpszProp == NULL) && (lpServicesIni[n].ulKey != 0))
			{
				EC_H(StringCchPrintf(
					szPropNum,
					_countof(szPropNum),
					_T("%lx"), // STRING_OK
					lpServicesIni[n].ulKey));

				lpszProp = szPropNum;
			}

			//
			// Write the item to MAPISVC.INF
			//

			DebugPrint(DBGGeneric, L"\tWriting: \"%ws\"::\"%ws\"::\"%ws\"\n",
				LPCTSTRToWstring(lpServicesIni[n].lpszSection).c_str(),
				LPCTSTRToWstring(lpszProp).c_str(),
				LPCTSTRToWstring(lpszValue).c_str());

			EC_B(WritePrivateProfileString(
				lpServicesIni[n].lpszSection,
				lpszProp,
				lpszValue,
				szServicesIni));
			n++;
		}

		// Flush the information - we can ignore the return code
		WritePrivateProfileString(NULL, NULL, NULL, szServicesIni);
	}

	return hRes;
} // HrSetProfileParameters

void AddServicesToMapiSvcInf()
{
	HRESULT hRes = S_OK;
	CEditor MyData(
		NULL,
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
			hRes = S_OK;
		}
	}
} // AddServicesToMapiSvcInf

void RemoveServicesFromMapiSvcInf()
{
	HRESULT hRes = S_OK;
	CEditor MyData(
		NULL,
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
			hRes = S_OK;
		}
	}
} // RemoveServicesFromMapiSvcInf

#define PR_MARKER PR_BODY_A
#define MARKER_STRING "MFCMAPI Existing Provider Marker" // STRING_OK
// Walk through providers and add/remove our tag
// bAddMark of true will add tag, bAddMark of false will remove it
_Check_return_ HRESULT HrMarkExistingProviders(_In_ LPSERVICEADMIN lpServiceAdmin, bool bAddMark)
{
	HRESULT			hRes = S_OK;
	LPMAPITABLE		lpProviderTable = NULL;

	if (!lpServiceAdmin) return MAPI_E_INVALID_PARAMETER;

	static const SizedSPropTagArray(1, pTagUID) =
	{
		1,
		PR_SERVICE_UID
	};

	EC_MAPI(lpServiceAdmin->GetMsgServiceTable(0, &lpProviderTable));

	if (lpProviderTable)
	{
		LPSRowSet lpRowSet = NULL;

		EC_MAPI(HrQueryAllRows(lpProviderTable, (LPSPropTagArray)&pTagUID, NULL, NULL, 0, &lpRowSet));

		if (lpRowSet) DebugPrintSRowSet(DBGGeneric, lpRowSet, NULL);

		if (lpRowSet && lpRowSet->cRows >= 1)
		{
			for (ULONG i = 0; i < lpRowSet->cRows; i++)
			{
				LPSRow		 lpCurRow = NULL;
				LPSPropValue lpServiceUID = NULL;

				hRes = S_OK;
				lpCurRow = &lpRowSet->aRow[i];

				lpServiceUID = PpropFindProp(
					lpCurRow->lpProps,
					lpCurRow->cValues,
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
						if (bAddMark)
						{
							SPropValue	PropVal;
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
} // HrMarkExistingProviders

// Returns first provider without our mark on it
_Check_return_ HRESULT HrFindUnmarkedProvider(_In_ LPSERVICEADMIN lpServiceAdmin, _Deref_out_opt_ LPSRowSet* lpRowSet)
{
	HRESULT			hRes = S_OK;
	LPMAPITABLE		lpProviderTable = NULL;
	LPPROFSECT		lpSect = NULL;

	if (!lpServiceAdmin || !lpRowSet) return MAPI_E_INVALID_PARAMETER;

	*lpRowSet = NULL;

	static const SizedSPropTagArray(1, pTagUID) =
	{
		1,
		PR_SERVICE_UID
	};

	EC_MAPI(lpServiceAdmin->GetMsgServiceTable(0, &lpProviderTable));

	if (lpProviderTable)
	{
		EC_MAPI(lpProviderTable->SetColumns((LPSPropTagArray)&pTagUID, TBL_BATCH));
		for (;;)
		{
			EC_MAPI(lpProviderTable->QueryRows(1, 0, lpRowSet));
			if (S_OK == hRes && *lpRowSet && 1 == (*lpRowSet)->cRows)
			{
				LPSRow		 lpCurRow = NULL;
				LPSPropValue lpServiceUID = NULL;

				lpCurRow = &(*lpRowSet)->aRow[0];

				lpServiceUID = PpropFindProp(
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
						SPropTagArray	pTagArray = { 1, PR_MARKER };
						ULONG			ulPropVal = 0;
						LPSPropValue	lpsPropVal = NULL;
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
						lpSect = NULL;
					}
				}
				// go on to next one in the loop
				FreeProws(*lpRowSet);
				*lpRowSet = NULL;
			}
			else
			{
				// no more hits - get out of the loop
				FreeProws(*lpRowSet);
				*lpRowSet = NULL;
				break;
			}
		}
		if (lpSect) lpSect->Release();

		lpProviderTable->Release();
	}
	return hRes;
} // HrFindUnmarkedProvider

_Check_return_ HRESULT HrAddServiceToProfile(
	_In_z_ LPCSTR lpszServiceName, // Service Name
	_In_ ULONG_PTR ulUIParam, // hwnd for CreateMsgService
	ULONG ulFlags, // Flags for CreateMsgService
	ULONG cPropVals, // Count of properties for ConfigureMsgService
	_In_opt_ LPSPropValue lpPropVals, // Properties for ConfigureMsgService
	_In_z_ LPCSTR lpszProfileName) // profile name
{
	HRESULT			hRes = S_OK;
	LPPROFADMIN		lpProfAdmin = NULL;
	LPSERVICEADMIN	lpServiceAdmin = NULL;
	LPSRowSet		lpRowSet = NULL;

	DebugPrint(DBGGeneric, L"HrAddServiceToProfile(%hs,%hs)\n", lpszServiceName, lpszProfileName);

	if (!lpszServiceName || !lpszProfileName) return MAPI_E_INVALID_PARAMETER;

	// Connect to Profile Admin interface.
	EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));
	if (!lpProfAdmin) return hRes;

	EC_MAPI(lpProfAdmin->AdminServices(
		(LPTSTR)lpszProfileName,
		(LPTSTR)_T(""),
		0,
		0,
		&lpServiceAdmin));

	if (lpServiceAdmin)
	{
		MAPIUID uidService = { 0 };
		LPMAPIUID lpuidService = &uidService;

		LPSERVICEADMIN2 lpServiceAdmin2 = NULL;
		WC_MAPI(lpServiceAdmin->QueryInterface(IID_IMsgServiceAdmin2, (LPVOID*)&lpServiceAdmin2));

		if (SUCCEEDED(hRes) && lpServiceAdmin2)
		{
			EC_H_MSG(lpServiceAdmin2->CreateMsgServiceEx(
				(LPTSTR)lpszServiceName,
				(LPTSTR)lpszServiceName,
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
				(LPTSTR)lpszServiceName,
				(LPTSTR)lpszServiceName,
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
					LPSPropValue lpServiceUIDProp = NULL;

					lpServiceUIDProp = PpropFindProp(
						lpRowSet->aRow[0].lpProps,
						lpRowSet->aRow[0].cValues,
						PR_SERVICE_UID);

					if (lpServiceUIDProp)
					{
						lpuidService = (LPMAPIUID)lpServiceUIDProp->Value.bin.lpb;
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
} // HrAddServiceToProfile

_Check_return_ HRESULT HrAddExchangeToProfile(
	_In_ ULONG_PTR ulUIParam, // hwnd for CreateMsgService
	_In_z_ LPCSTR lpszServerName,
	_In_z_ LPCSTR lpszMailboxName,
	_In_z_ LPCSTR lpszProfileName)
{
	HRESULT			hRes = S_OK;

	DebugPrint(DBGGeneric, L"HrAddExchangeToProfile(%hs,%hs,%hs)\n", lpszServerName, lpszMailboxName, lpszProfileName);

	if (!lpszServerName || !lpszMailboxName || !lpszProfileName) return MAPI_E_INVALID_PARAMETER;

#define NUMEXCHANGEPROPS 2
	SPropValue		PropVal[NUMEXCHANGEPROPS];
	PropVal[0].ulPropTag = PR_PROFILE_UNRESOLVED_SERVER;
	PropVal[0].Value.lpszA = (LPSTR)lpszServerName;
	PropVal[1].ulPropTag = PR_PROFILE_UNRESOLVED_NAME;
	PropVal[1].Value.lpszA = (LPSTR)lpszMailboxName;
	EC_H(HrAddServiceToProfile("MSEMS", ulUIParam, NULL, NUMEXCHANGEPROPS, PropVal, lpszProfileName)); // STRING_OK

	return hRes;
} // HrAddExchangeToProfile

_Check_return_ HRESULT HrAddPSTToProfile(
	_In_ ULONG_PTR ulUIParam, // hwnd for CreateMsgService
	bool bUnicodePST,
	_In_z_ LPCTSTR lpszPSTPath, // PST name
	_In_z_ LPCSTR lpszProfileName, // profile name
	bool bPasswordSet, // whether or not to include a password
	_In_z_ LPCSTR lpszPassword) // password to include
{
	HRESULT			hRes = S_OK;

	DebugPrint(DBGGeneric, L"HrAddPSTToProfile(0x%X,%ws,%hs,0x%X,%hs)\n", bUnicodePST, LPCTSTRToWstring(lpszPSTPath).c_str(), lpszProfileName, bPasswordSet, lpszPassword);

	if (!lpszPSTPath || !lpszProfileName) return MAPI_E_INVALID_PARAMETER;

	SPropValue PropVal[2];

	PropVal[0].ulPropTag = CHANGE_PROP_TYPE(PR_PST_PATH, PT_TSTRING);
	PropVal[0].Value.LPSZ = (LPTSTR)lpszPSTPath;
	PropVal[1].ulPropTag = PR_PST_PW_SZ_OLD;
	PropVal[1].Value.lpszA = (LPSTR)lpszPassword;

	if (bUnicodePST)
	{
		EC_H(HrAddServiceToProfile("MSUPST MS", ulUIParam, NULL, bPasswordSet ? 2 : 1, PropVal, lpszProfileName)); // STRING_OK
	}
	else
	{
		EC_H(HrAddServiceToProfile("MSPST MS", ulUIParam, NULL, bPasswordSet ? 2 : 1, PropVal, lpszProfileName)); // STRING_OK
	}

	return hRes;
} // HrAddPSTToProfile

// $--HrCreateProfile---------------------------------------------
// Creates an empty profile.
// -----------------------------------------------------------------------------
_Check_return_ HRESULT HrCreateProfile(
	_In_z_ LPCSTR lpszProfileName) // profile name
{
	HRESULT			hRes = S_OK;
	LPPROFADMIN		lpProfAdmin = NULL;

	DebugPrint(DBGGeneric, L"HrCreateProfile(%hs)\n", lpszProfileName);

	if (!lpszProfileName) return MAPI_E_INVALID_PARAMETER;

	// Connect to Profile Admin interface.
	EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));
	if (!lpProfAdmin) return hRes;

	// Create the profile
	WC_MAPI(lpProfAdmin->CreateProfile(
		(LPTSTR)lpszProfileName,
		NULL,
		0,
		NULL)); // fMapiUnicode is not supported!
	if (S_OK != hRes)
	{
		// Did it fail because a profile of this name already exists?
		EC_H_MSG(HrMAPIProfileExists(lpProfAdmin, lpszProfileName),
			IDS_DUPLICATEPROFILE);
	}

	lpProfAdmin->Release();

	return hRes;
} // HrCreateProfile

// $--HrRemoveProfile---------------------------------------------------------
// Removes a profile.
// ------------------------------------------------------------------------------
_Check_return_ HRESULT HrRemoveProfile(
	_In_z_ LPCSTR lpszProfileName)
{
	HRESULT		hRes = S_OK;
	LPPROFADMIN	lpProfAdmin = NULL;

	DebugPrint(DBGGeneric, L"HrRemoveProfile(%hs)\n", lpszProfileName);
	if (!lpszProfileName) return MAPI_E_INVALID_PARAMETER;

	EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));
	if (!lpProfAdmin) return hRes;

	EC_MAPI(lpProfAdmin->DeleteProfile((LPTSTR)lpszProfileName, 0));

	lpProfAdmin->Release();

	RegFlushKey(HKEY_LOCAL_MACHINE);
	RegFlushKey(HKEY_CURRENT_USER);

	return hRes;
} // HrRemoveProfile

// $--HrSetDefaultProfile---------------------------------------------------------
// Set a profile as default.
// ------------------------------------------------------------------------------
_Check_return_ HRESULT HrSetDefaultProfile(
	_In_z_ LPCSTR lpszProfileName)
{
	HRESULT hRes = S_OK;
	LPPROFADMIN lpProfAdmin = NULL;

	DebugPrint(DBGGeneric, L"HrRemoveProfile(%hs)\n", lpszProfileName);
	if (!lpszProfileName) return MAPI_E_INVALID_PARAMETER;

	EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));
	if (!lpProfAdmin) return hRes;

	EC_MAPI(lpProfAdmin->SetDefaultProfile((LPTSTR)lpszProfileName, 0));

	lpProfAdmin->Release();

	RegFlushKey(HKEY_LOCAL_MACHINE);
	RegFlushKey(HKEY_CURRENT_USER);

	return hRes;
} // HrSetDefaultProfile

// $--HrMAPIProfileExists---------------------------------------------------------
// Checks for an existing profile.
// -----------------------------------------------------------------------------
_Check_return_ HRESULT HrMAPIProfileExists(
	_In_ LPPROFADMIN lpProfAdmin,
	_In_z_ LPCSTR lpszProfileName)
{
	HRESULT hRes = S_OK;
	LPMAPITABLE lpTable = NULL;
	LPSRowSet lpRows = NULL;
	LPSPropValue lpProp = NULL;
	ULONG i = 0;

	static const SizedSPropTagArray(1, rgPropTag) =
	{
		1,
		PR_DISPLAY_NAME_A
	};

	DebugPrint(DBGGeneric, L"HrMAPIProfileExists()\n");
	if (!lpProfAdmin || !lpszProfileName) return MAPI_E_INVALID_PARAMETER;

	// Get a table of existing profiles

	EC_MAPI(lpProfAdmin->GetProfileTable(
		0,
		&lpTable));
	if (!lpTable) return hRes;

	EC_MAPI(HrQueryAllRows(
		lpTable,
		(LPSPropTagArray)&rgPropTag,
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

			if (!FAILED(hRes)) for (i = 0; i < lpRows->cRows; i++)
			{
				hRes = S_OK;
				lpProp = lpRows->aRow[i].lpProps;

				ULONG ulComp = NULL;

				EC_D(ulComp, CompareStringA(
					g_lcid, // LOCALE_INVARIANT,
					NORM_IGNORECASE,
					lpProp[0].Value.lpszA,
					-1,
					lpszProfileName,
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
} // HrMAPIProfileExists

_Check_return_ HRESULT GetProfileServiceVersion(_In_z_ LPCSTR lpszProfileName,
	_Out_ ULONG* lpulServerVersion,
	_Out_ EXCHANGE_STORE_VERSION_NUM* lpStoreVersion,
	_Out_ bool* lpbFoundServerVersion,
	_Out_ bool* lpbFoundServerFullVersion)
{
	if (!lpszProfileName
		|| !lpulServerVersion
		|| !lpStoreVersion
		|| !lpbFoundServerVersion
		|| !lpbFoundServerFullVersion) return MAPI_E_INVALID_PARAMETER;
	*lpulServerVersion = NULL;
	memset(lpStoreVersion, 0, sizeof(EXCHANGE_STORE_VERSION_NUM));
	*lpbFoundServerVersion = false;
	*lpbFoundServerFullVersion = false;

	HRESULT        hRes = S_OK;
	LPPROFADMIN    lpProfAdmin = NULL;
	LPSERVICEADMIN lpServiceAdmin = NULL;

	DebugPrint(DBGGeneric, L"GetProfileServiceVersion(%hs)\n", lpszProfileName);

	EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));
	if (!lpProfAdmin) return hRes;

	EC_MAPI(lpProfAdmin->AdminServices(
		(LPTSTR)lpszProfileName,
		(LPTSTR)_T(""),
		0,
		0,
		&lpServiceAdmin));

	if (lpServiceAdmin)
	{
		LPPROFSECT lpProfSect = NULL;
		EC_MAPI(lpServiceAdmin->OpenProfileSection(
			(LPMAPIUID)pbGlobalProfileSectionGuid,
			NULL,
			0,
			&lpProfSect));
		if (lpProfSect)
		{
			LPSPropValue lpServerVersion = NULL;
			WC_MAPI(HrGetOneProp(lpProfSect, PR_PROFILE_SERVER_VERSION, &lpServerVersion));

			if (SUCCEEDED(hRes) && lpServerVersion && PR_PROFILE_SERVER_VERSION == lpServerVersion->ulPropTag)
			{
				*lpbFoundServerVersion = true;
				*lpulServerVersion = lpServerVersion->Value.l;
			}
			MAPIFreeBuffer(lpServerVersion);
			hRes = S_OK;

			LPSPropValue lpServerFullVersion = NULL;
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
} // GetProfileServiceVersion

// $--HrCopyProfile---------------------------------------------------------
// Copies a profile.
// ------------------------------------------------------------------------------
_Check_return_ HRESULT HrCopyProfile(_In_z_ LPCSTR lpszOldProfileName,
	_In_z_ LPCSTR lpszNewProfileName)
{
	HRESULT hRes = S_OK;
	LPPROFADMIN lpProfAdmin = NULL;

	DebugPrint(DBGGeneric, L"HrCopyProfile(%hs, %hs)\n", lpszOldProfileName, lpszNewProfileName);
	if (!lpszOldProfileName || !lpszNewProfileName) return MAPI_E_INVALID_PARAMETER;

	EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));
	if (!lpProfAdmin) return hRes;

	EC_MAPI(lpProfAdmin->CopyProfile((LPTSTR)lpszOldProfileName, NULL, (LPTSTR)lpszNewProfileName, NULL, NULL));

	lpProfAdmin->Release();

	return hRes;
} // HrCopyProfile
#endif

#define MAPI_FORCE_ACCESS 0x00080000

_Check_return_ HRESULT OpenProfileSection(_In_ LPSERVICEADMIN lpServiceAdmin, _In_ LPSBinary lpServiceUID, _Deref_out_opt_ LPPROFSECT* lppProfSect)
{
	HRESULT		hRes = S_OK;

	DebugPrint(DBGOpenItemProp, L"OpenProfileSection opening lpServiceUID = ");
	DebugPrintBinary(DBGOpenItemProp, lpServiceUID);
	DebugPrint(DBGOpenItemProp, L"\n");

	if (!lpServiceUID || !lpServiceAdmin || !lppProfSect) return MAPI_E_INVALID_PARAMETER;
	*lppProfSect = NULL;

	// First, we try the normal way of opening the profile section:
	WC_MAPI(lpServiceAdmin->OpenProfileSection(
		(LPMAPIUID)lpServiceUID->lpb,
		NULL,
		MAPI_MODIFY | MAPI_FORCE_ACCESS, // passing this flag might actually work with Outlook 2000 and XP
		lppProfSect));

	if (!*lppProfSect)
	{
		hRes = S_OK;
		///////////////////////////////////////////////////////////////////
		// HACK CENTRAL.  This is a MAJOR hack.  MAPI will always return E_ACCESSDENIED
		// when we open a profile section on the service if we are a client.  The workaround
		// (HACK) is to call into one of MAPI's internal functions that bypasses
		// the security check.  We build a Interface to it and then point to it from our
		// offset of 0x48.  USE AT YOUR OWN RISK!  NOT SUPPORTED!
		interface IOpenSectionHack : public IUnknown
		{
		public:
			virtual HRESULT STDMETHODCALLTYPE OpenSection(LPMAPIUID, ULONG, LPPROFSECT*) = 0;
		};

		IOpenSectionHack** ppProfile;

		ppProfile = (IOpenSectionHack**)((((BYTE*)lpServiceAdmin) + 0x48));

		// Now, we want to get open the Services Profile Section and store that
		// interface with the Object

		if (ppProfile && *ppProfile)
		{
			EC_MAPI((*ppProfile)->OpenSection(
				(LPMAPIUID)lpServiceUID->lpb,
				MAPI_MODIFY,
				lppProfSect));
		}
		else
		{
			hRes = MAPI_E_NOT_FOUND;
		}
		// END OF HACK.  I'm amazed that this works....
		///////////////////////////////////////////////////////////////////
	}
	return hRes;
} // OpenProfileSection

_Check_return_ HRESULT OpenProfileSection(_In_ LPPROVIDERADMIN lpProviderAdmin, _In_ LPSBinary lpProviderUID, _Deref_out_ LPPROFSECT* lppProfSect)
{
	HRESULT			hRes = S_OK;

	DebugPrint(DBGOpenItemProp, L"OpenProfileSection opening lpServiceUID = ");
	DebugPrintBinary(DBGOpenItemProp, lpProviderUID);
	DebugPrint(DBGOpenItemProp, L"\n");

	if (!lpProviderUID || !lpProviderAdmin || !lppProfSect) return MAPI_E_INVALID_PARAMETER;
	*lppProfSect = NULL;

	WC_MAPI(lpProviderAdmin->OpenProfileSection(
		(LPMAPIUID)lpProviderUID->lpb,
		NULL,
		MAPI_MODIFY | MAPI_FORCE_ACCESS,
		lppProfSect));
	if (!*lppProfSect)
	{
		hRes = S_OK;

		// We only do this hack as a last resort - it can crash some versions of Outlook, but is required for Exchange
		*(((BYTE*)lpProviderAdmin) + 0x60) = 0x2; // Use at your own risk! NOT SUPPORTED!

		WC_MAPI(lpProviderAdmin->OpenProfileSection(
			(LPMAPIUID)lpProviderUID->lpb,
			NULL,
			MAPI_MODIFY,
			lppProfSect));
	}
	return hRes;
} // OpenProfileSection