// Collection of useful MAPI functions

#include <StdAfx.h>
#include <UI/profile.h>
#include <core/utility/import.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/utility/file.h>
#include <core/addin/mfcmapi.h>
#include <core/mapi/extraPropTags.h>

namespace ui::profile
{
	std::wstring LaunchProfileWizard(_In_ HWND hParentWnd, ULONG ulFlags, _In_ const std::wstring& szServiceNameToAdd)
	{
		CHAR szProfName[80] = {0};
		constexpr ULONG cchProfName = _countof(szProfName);
		const auto szServiceNameToAddA = strings::wstringTostring(szServiceNameToAdd);
		LPCSTR szServices[] = {szServiceNameToAddA.c_str(), nullptr};

		output::DebugPrint(
			output::dbgLevel::Generic, L"LaunchProfileWizard: Using LAUNCHWIZARDENTRY to launch wizard API.\n");

		// Call LaunchWizard to add the service.
		const auto hRes = WC_MAPI(LaunchWizard(hParentWnd, ulFlags, szServices, cchProfName, szProfName));
		if (hRes == MAPI_E_CALL_FAILED)
		{
			CHECKHRESMSG(hRes, IDS_LAUNCHWIZARDFAILED);
		}
		else
			CHECKHRES(hRes);

		if (SUCCEEDED(hRes))
		{
			output::DebugPrint(
				output::dbgLevel::Generic, L"LaunchProfileWizard: Profile \"%hs\" configured.\n", szProfName);
		}

		return strings::LPCSTRToWstring(szProfName);
	}

	void DisplayMAPISVCPath(_In_ CWnd* pParentWnd)
	{
		output::DebugPrint(output::dbgLevel::Generic, L"DisplayMAPISVCPath()\n");

		dialog::editor::CEditor MyData(pParentWnd, IDS_MAPISVCTITLE, IDS_MAPISVCTEXT, CEDITOR_BUTTON_OK);
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_FILEPATH, true));
		MyData.SetStringW(0, GetMAPISVCPath());

		static_cast<void>(MyData.DisplayDialog());
	}

	struct SERVICESINIREC
	{
		LPCWSTR lpszSection;
		LPCWSTR lpszKey;
		ULONG ulKey;
		LPCWSTR lpszValue;
	};

	static SERVICESINIREC aEMSServicesIni[] = {
		// clang-format off
			{L"Default Services", L"MSEMS", 0L, L"Microsoft Exchange Server"}, // STRING_OK
			{L"Services", L"MSEMS", 0L, L"Microsoft Exchange Server"}, // STRING_OK

			{L"MSEMS", L"PR_DISPLAY_NAME", 0L, L"Microsoft Exchange Server"}, // STRING_OK
			{L"MSEMS", L"Sections", 0L, L"MSEMS_MSMail_Section"}, // STRING_OK
			{L"MSEMS", L"PR_SERVICE_DLL_NAME", 0L, L"emsui.dll"}, // STRING_OK
			{L"MSEMS", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D26A-C2AC-11D1-AD3E-10A0C911C9C0}"}, // STRING_OK
			{L"MSEMS", L"PR_SERVICE_ENTRY_NAME", 0L, L"EMSCfg"}, // STRING_OK
			{L"MSEMS", L"PR_RESOURCE_FLAGS", 0L, L"SERVICE_SINGLE_COPY"}, // STRING_OK
			{L"MSEMS", L"WIZARD_ENTRY_NAME", 0L, L"EMSWizardEntry"}, // STRING_OK
			{L"MSEMS", L"Providers", 0L, L"EMS_DSA, EMS_MDB_public, EMS_MDB_private, EMS_RXP, EMS_MSX, EMS_Hook"}, // STRING_OK
			{L"MSEMS", L"PR_SERVICE_SUPPORT_FILES", 0L, L"emsui.dll,emsabp.dll,emsmdb.dll"}, // STRING_OK

			{L"EMS_MDB_public", L"PR_RESOURCE_TYPE", 0L, L"MAPI_STORE_PROVIDER"}, // STRING_OK
			{L"EMS_MDB_public", L"PR_PROVIDER_DLL_NAME", 0L, L"EMSMDB.DLL"}, // STRING_OK
			{L"EMS_MDB_public", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D26A-C2AC-11D1-AD3E-10A0C911C9C0}"}, // STRING_OK
			{L"EMS_MDB_public", L"PR_RESOURCE_FLAGS", 0L, L"STATUS_NO_DEFAULT_STORE"}, // STRING_OK
			{L"EMS_MDB_public", nullptr, PR_PROFILE_OPEN_FLAGS, L"06000000"}, // STRING_OK
			{L"EMS_MDB_public", nullptr, PR_PROFILE_TYPE, L"03000000"}, // STRING_OK
			{L"EMS_MDB_public", nullptr, PR_MDB_PROVIDER, L"78b2fa70aff711cd9bc800aa002fc45a"}, // STRING_OK
			{L"EMS_MDB_public", L"PR_DISPLAY_NAME", 0L, L"Public Folders"}, // STRING_OK
			{L"EMS_MDB_public", L"PR_PROVIDER_DISPLAY", 0L, L"Microsoft Exchange Message Store"}, // STRING_OK

			{L"EMS_MDB_private", L"PR_PROVIDER_DLL_NAME", 0L, L"EMSMDB.DLL"}, // STRING_OK
			{L"EMS_MDB_private", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D26A-C2AC-11D1-AD3E-10A0C911C9C0}"}, // STRING_OK
			{L"EMS_MDB_private", L"PR_RESOURCE_TYPE", 0L, L"MAPI_STORE_PROVIDER"}, // STRING_OK
			{L"EMS_MDB_private", L"PR_RESOURCE_FLAGS", 0L, L"STATUS_PRIMARY_IDENTITY|STATUS_DEFAULT_STORE|STATUS_PRIMARY_STORE"}, // STRING_OK
			{L"EMS_MDB_private", nullptr, PR_PROFILE_OPEN_FLAGS, L"0C000000"}, // STRING_OK
			{L"EMS_MDB_private", nullptr, PR_PROFILE_TYPE, L"01000000"}, // STRING_OK
			{L"EMS_MDB_private", nullptr, PR_MDB_PROVIDER, L"5494A1C0297F101BA58708002B2A2517"}, // STRING_OK
			{L"EMS_MDB_private", L"PR_DISPLAY_NAME", 0L, L"Private Folders"}, // STRING_OK
			{L"EMS_MDB_private", L"PR_PROVIDER_DISPLAY", 0L, L"Microsoft Exchange Message Store"}, // STRING_OK

			{L"EMS_DSA", L"PR_DISPLAY_NAME", 0L, L"Microsoft Exchange Directory Service"}, // STRING_OK
			{L"EMS_DSA", L"PR_PROVIDER_DISPLAY", 0L, L"Microsoft Exchange Directory Service"}, // STRING_OK
			{L"EMS_DSA", L"PR_PROVIDER_DLL_NAME", 0L, L"EMSABP.DLL"}, // STRING_OK
			{L"EMS_DSA", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D26A-C2AC-11D1-AD3E-10A0C911C9C0}"}, // STRING_OK
			{L"EMS_DSA", L"PR_RESOURCE_TYPE", 0L, L"MAPI_AB_PROVIDER"}, // STRING_OK

			{L"MSEMS_MSMail_Section", L"UID", 0L, L"13DBB0C8AA05101A9BB000AA002FC45A"}, // STRING_OK
			{L"MSEMS_MSMail_Section", nullptr, PR_PROFILE_VERSION, L"01050000"}, // STRING_OK
			{L"MSEMS_MSMail_Section", nullptr, PR_PROFILE_CONFIG_FLAGS, L"04000000"}, // STRING_OK
			{L"MSEMS_MSMail_Section", nullptr, PR_PROFILE_TRANSPORT_FLAGS, L"03000000"}, // STRING_OK
			{L"MSEMS_MSMail_Section", nullptr, PR_PROFILE_CONNECT_FLAGS, L"02000000"}, // STRING_OK

			{L"EMS_RXP", L"PR_DISPLAY_NAME", 0L, L"Microsoft Exchange Remote Transport"}, // STRING_OK
			{L"EMS_RXP", L"PR_PROVIDER_DISPLAY", 0L, L"Microsoft Exchange Remote Transport"}, // STRING_OK
			{L"EMS_RXP", L"PR_PROVIDER_DLL_NAME", 0L, L"EMSUI.DLL"}, // STRING_OK
			{L"EMS_RXP", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D26A-C2AC-11D1-AD3E-10A0C911C9C0}"}, // STRING_OK
			{L"EMS_RXP", L"PR_RESOURCE_TYPE", 0L, L"MAPI_TRANSPORT_PROVIDER"}, // STRING_OK
			{L"EMS_RXP", nullptr, PR_PROFILE_OPEN_FLAGS, L"40000000"}, // STRING_OK
			{L"EMS_RXP", nullptr, PR_PROFILE_TYPE, L"0A000000"}, // STRING_OK

			{L"EMS_MSX", L"PR_DISPLAY_NAME", 0L, L"Microsoft Exchange Transport"}, // STRING_OK
			{L"EMS_MSX", L"PR_PROVIDER_DISPLAY", 0L, L"Microsoft Exchange Transport"}, // STRING_OK
			{L"EMS_MSX", L"PR_PROVIDER_DLL_NAME", 0L, L"EMSMDB.DLL"}, // STRING_OK
			{L"EMS_MSX", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D26A-C2AC-11D1-AD3E-10A0C911C9C0}"}, // STRING_OK
			{L"EMS_MSX", L"PR_RESOURCE_TYPE", 0L, L"MAPI_TRANSPORT_PROVIDER"}, // STRING_OK
			{L"EMS_MSX", nullptr, PR_PROFILE_OPEN_FLAGS, L"00000000"}, // STRING_OK

			{L"EMS_Hook", L"PR_DISPLAY_NAME", 0L, L"Microsoft Exchange Hook"}, // STRING_OK
			{L"EMS_Hook", L"PR_PROVIDER_DISPLAY", 0L, L"Microsoft Exchange Hook"}, // STRING_OK
			{L"EMS_Hook", L"PR_PROVIDER_DLL_NAME", 0L, L"EMSMDB.DLL"}, // STRING_OK
			{L"EMS_Hook", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D26A-C2AC-11D1-AD3E-10A0C911C9C0}"}, // STRING_OK
			{L"EMS_Hook", L"PR_RESOURCE_TYPE", 0L, L"MAPI_HOOK_PROVIDER"}, // STRING_OK
			{L"EMS_Hook", L"PR_RESOURCE_FLAGS", 0L, L"HOOK_INBOUND"}, // STRING_OK

			{nullptr, nullptr, 0L, nullptr}
		// clang-format on
	};

	// Here's an example of the array to use to remove a service
	static SERVICESINIREC aREMOVE_MSEMSServicesIni[] = {
		// clang-format off
			{L"Default Services", L"MSEMS", 0L, nullptr}, // STRING_OK
			{L"Services", L"MSEMS", 0L, nullptr}, // STRING_OK
			{L"MSEMS", nullptr, 0L, nullptr}, // STRING_OK
			{L"EMS_MDB_public", nullptr, 0L, nullptr}, // STRING_OK
			{L"EMS_MDB_private", nullptr, 0L, nullptr}, // STRING_OK
			{L"EMS_DSA", nullptr, 0L, nullptr}, // STRING_OK
			{L"MSEMS_MSMail_Section", nullptr, 0L, nullptr}, // STRING_OK
			{L"EMS_RXP", nullptr, 0L, nullptr}, // STRING_OK
			{L"EMS_MSX", nullptr, 0L, nullptr}, // STRING_OK
			{L"EMS_Hook", nullptr, 0L, nullptr}, // STRING_OK
			{L"EMSDelegate", nullptr, 0L, nullptr}, // STRING_OK
			{nullptr, nullptr, 0L, nullptr}
		// clang-format on
	};

	static SERVICESINIREC aPSTServicesIni[] = {
		// clang-format off
			{L"Services", L"MSPST MS", 0L, L"Personal Folders File (.pst)"}, // STRING_OK
			{L"Services", L"MSPST AB", 0L, L"Personal Address Book"}, // STRING_OK

			{L"MSPST AB", L"PR_DISPLAY_NAME", 0L, L"Personal Address Book"}, // STRING_OK
			{L"MSPST AB", L"Providers", 0L, L"MSPST ABP"}, // STRING_OK
			{L"MSPST AB", L"PR_SERVICE_DLL_NAME", 0L, L"MSPST.DLL"}, // STRING_OK
			{L"MSPST AB", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D262-C2AC-11D1-AD3E-10A0C911C9C0}"}, // STRING_OK
			{L"MSPST AB", L"PR_SERVICE_SUPPORT_FILES", 0L, L"MSPST.DLL"}, // STRING_OK
			{L"MSPST AB", L"PR_SERVICE_ENTRY_NAME", 0L, L"PABServiceEntry"}, // STRING_OK
			{L"MSPST AB", L"PR_RESOURCE_FLAGS", 0L, L"SERVICE_SINGLE_COPY|SERVICE_NO_PRIMARY_IDENTITY"}, // STRING_OK

			{L"MSPST ABP", L"PR_PROVIDER_DLL_NAME", 0L, L"MSPST.DLL"}, // STRING_OK
			{L"MSPST ABP", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D262-C2AC-11D1-AD3E-10A0C911C9C0}"}, // STRING_OK
			{L"MSPST ABP", L"PR_RESOURCE_TYPE", 0L, L"MAPI_AB_PROVIDER"}, // STRING_OK
			{L"MSPST ABP", L"PR_DISPLAY_NAME", 0L, L"Personal Address Book"}, // STRING_OK
			{L"MSPST ABP", L"PR_PROVIDER_DISPLAY", 0L, L"Personal Address Book"}, // STRING_OK
			{L"MSPST ABP", L"PR_SERVICE_DLL_NAME", 0L, L"MSPST.DLL"}, // STRING_OK

			{L"MSPST MS", L"Providers", 0L, L"MSPST MSP"}, // STRING_OK
			{L"MSPST MS", L"PR_SERVICE_DLL_NAME", 0L, L"mspst.dll"}, // STRING_OK
			{L"MSPST MS", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D262-C2AC-11D1-AD3E-10A0C911C9C0}"}, // STRING_OK
			{L"MSPST MS", L"PR_SERVICE_SUPPORT_FILES", 0L, L"mspst.dll"}, // STRING_OK
			{L"MSPST MS", L"PR_SERVICE_ENTRY_NAME", 0L, L"PSTServiceEntry"}, // STRING_OK
			{L"MSPST MS", L"PR_RESOURCE_FLAGS", 0L, L"SERVICE_NO_PRIMARY_IDENTITY"}, // STRING_OK

			{L"MSPST MSP", nullptr, PR_MDB_PROVIDER, L"4e495441f9bfb80100aa0037d96e0000"}, // STRING_OK
			{L"MSPST MSP", L"PR_PROVIDER_DLL_NAME", 0L, L"mspst.dll"}, // STRING_OK
			{L"MSPST MSP", L"PR_SERVICE_INSTALL_ID", 0L, L"{6485D262-C2AC-11D1-AD3E-10A0C911C9C0}"}, // STRING_OK
			{L"MSPST MSP", L"PR_RESOURCE_TYPE", 0L, L"MAPI_STORE_PROVIDER"}, // STRING_OK
			{L"MSPST MSP", L"PR_RESOURCE_FLAGS", 0L, L"STATUS_DEFAULT_STORE"}, // STRING_OK
			{L"MSPST MSP", L"PR_DISPLAY_NAME", 0L, L"Personal Folders"}, // STRING_OK
			{L"MSPST MSP", L"PR_PROVIDER_DISPLAY", 0L, L"Personal Folders File (.pst)"}, // STRING_OK
			{L"MSPST MSP", L"PR_SERVICE_DLL_NAME", 0L, L"mspst.dll"}, // STRING_OK

			{nullptr, nullptr, 0L, nullptr}
		// clang-format on
	};

	static SERVICESINIREC aREMOVE_MSPSTServicesIni[] = {
		// clang-format off
			{L"Default Services", L"MSPST MS", 0L, nullptr}, // STRING_OK
			{L"Default Services", L"MSPST AB", 0L, nullptr}, // STRING_OK
			{L"Services", L"MSPST MS", 0L, nullptr}, // STRING_OK
			{L"Services", L"MSPST AB", 0L, nullptr}, // STRING_OK

			{L"MSPST AB", nullptr, 0L, nullptr}, // STRING_OK

			{L"MSPST ABP", nullptr, 0L, nullptr}, // STRING_OK

			{L"MSPST MS", nullptr, 0L, nullptr}, // STRING_OK

			{L"MSPST MSP", nullptr, 0L, nullptr}, // STRING_OK

			{nullptr, nullptr, 0L, nullptr}
		// clang-format on
	};

	// Function name : GetMAPISVCPath
	// Description : This will get the correct path to the MAPISVC.INF file.
	// Return type : wstring
	std::wstring GetMAPISVCPath()
	{
		auto szMAPIDir = import::GetMAPIPath(L"Microsoft Outlook"); // STRING_OK

		// We got the path to MAPI - need to strip it
		if (!szMAPIDir.empty())
		{
			szMAPIDir.erase(szMAPIDir.find_last_of(L'\\'), std::string::npos);
		}
		else
		{
			// Fall back on System32
			szMAPIDir = file::GetSystemDirectory();
		}

		if (!szMAPIDir.empty())
		{
			szMAPIDir += L"\\MAPISVC.INF";
		}

		return szMAPIDir;
	}

	// Add values to MAPISVC.INF
	_Check_return_ HRESULT HrSetProfileParameters(_In_ SERVICESINIREC* lpServicesIni)
	{
		auto hRes = S_OK;

		output::DebugPrint(output::dbgLevel::Generic, L"HrSetProfileParameters()\n");

		if (!lpServicesIni) return MAPI_E_INVALID_PARAMETER;

		auto szServicesIni = GetMAPISVCPath();

		if (szServicesIni.empty())
		{
			hRes = MAPI_E_NOT_FOUND;
		}
		else
		{
			output::DebugPrint(output::dbgLevel::Generic, L"Writing to this file: \"%ws\"\n", szServicesIni.c_str());

			// Loop through and add items to MAPISVC.INF
			auto n = 0;

			while (lpServicesIni[n].lpszSection != nullptr)
			{
				std::wstring lpszProp = lpServicesIni[n].lpszKey;
				std::wstring lpszValue = lpServicesIni[n].lpszValue;

				// Switch the property if necessary
				if (lpszProp.empty() && lpServicesIni[n].ulKey != 0)
				{
					lpszProp = strings::format(
						L"%lx", // STRING_OK
						lpServicesIni[n].ulKey);
				}

				// Write the item to MAPISVC.INF
				output::DebugPrint(
					output::dbgLevel::Generic,
					L"\tWriting: \"%ws\"::\"%ws\"::\"%ws\"\n",
					lpServicesIni[n].lpszSection,
					lpszProp.c_str(),
					lpszValue.c_str());

				hRes = EC_B(WritePrivateProfileStringW(
					lpServicesIni[n].lpszSection, lpszProp.c_str(), lpszValue.c_str(), szServicesIni.c_str()));
				n++;
			}

			// Flush the information - we can ignore the return code
			WritePrivateProfileStringW(nullptr, nullptr, nullptr, szServicesIni.c_str());
		}

		return hRes;
	}

	void AddServicesToMapiSvcInf()
	{
		dialog::editor::CEditor MyData(
			nullptr, IDS_ADDSERVICESTOINF, IDS_ADDSERVICESTOINFPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.AddPane(viewpane::CheckPane::Create(0, IDS_EXCHANGE, false, false));
		MyData.AddPane(viewpane::CheckPane::Create(1, IDS_PST, false, false));

		if (MyData.DisplayDialog())
		{
			if (MyData.GetCheck(0))
			{
				EC_H_S(HrSetProfileParameters(aEMSServicesIni));
			}

			if (MyData.GetCheck(1))
			{
				EC_H_S(HrSetProfileParameters(aPSTServicesIni));
			}
		}
	}

	void RemoveServicesFromMapiSvcInf()
	{
		dialog::editor::CEditor MyData(
			nullptr, IDS_REMOVEFROMINF, IDS_REMOVEFROMINFPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.AddPane(viewpane::CheckPane::Create(0, IDS_EXCHANGE, false, false));
		MyData.AddPane(viewpane::CheckPane::Create(1, IDS_PST, false, false));

		if (MyData.DisplayDialog())
		{
			if (MyData.GetCheck(0))
			{
				EC_H_S(HrSetProfileParameters(aREMOVE_MSEMSServicesIni));
			}

			if (MyData.GetCheck(1))
			{
				EC_H_S(HrSetProfileParameters(aREMOVE_MSPSTServicesIni));
			}
		}
	}
} // namespace ui::profile