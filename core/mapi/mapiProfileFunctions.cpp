#include <core/stdafx.h>
#include <core/mapi/mapiProfileFunctions.h>
#include <core/mapi/extraPropTags.h>
#include <core/utility/output.h>
#include <core/utility/strings.h>
#include <core/utility/error.h>
#include <core/mapi/mapiFunctions.h>
#include <core/mapi/interfaces.h>
#include <core/mapi/mapiOutput.h>

namespace mapi::profile
{
	constexpr ULONG PR_MARKER = PR_BODY_A;
	constexpr LPSTR MARKER_STRING = const_cast<LPSTR>("MFCMAPI Existing Provider Marker"); // STRING_OK
	// Walk through providers and add/remove our tag
	// bAddMark of true will add tag, bAddMark of false will remove it
	_Check_return_ HRESULT HrMarkExistingProviders(_In_ LPSERVICEADMIN lpServiceAdmin, bool bAddMark)
	{
		LPMAPITABLE lpProviderTable = nullptr;

		if (!lpServiceAdmin) return MAPI_E_INVALID_PARAMETER;

		static const SizedSPropTagArray(1, pTagUID) = {1, {PR_SERVICE_UID}};

		auto hRes = EC_MAPI(lpServiceAdmin->GetMsgServiceTable(0, &lpProviderTable));
		if (lpProviderTable)
		{
			LPSRowSet lpRowSet = nullptr;

			hRes = EC_MAPI(HrQueryAllRows(lpProviderTable, LPSPropTagArray(&pTagUID), nullptr, nullptr, 0, &lpRowSet));

			if (lpRowSet) output::outputSRowSet(output::dbgLevel::Generic, nullptr, lpRowSet, nullptr);

			if (lpRowSet && lpRowSet->cRows >= 1)
			{
				for (ULONG i = 0; i < lpRowSet->cRows; i++)
				{
					const auto lpCurRow = &lpRowSet->aRow[i];

					auto lpServiceUID = PpropFindProp(lpCurRow->lpProps, lpCurRow->cValues, PR_SERVICE_UID);

					if (lpServiceUID)
					{
						auto lpSect = OpenProfileSection(lpServiceAdmin, &mapi::getBin(lpServiceUID));
						if (lpSect)
						{
							if (bAddMark)
							{
								SPropValue PropVal;
								PropVal.ulPropTag = PR_MARKER;
								PropVal.Value.lpszA = MARKER_STRING;
								EC_MAPI_S(lpSect->SetProps(1, &PropVal, nullptr));
							}
							else
							{
								SPropTagArray pTagArray = {1, {PR_MARKER}};
								WC_MAPI_S(lpSect->DeleteProps(&pTagArray, nullptr));
							}

							hRes = EC_MAPI(lpSect->SaveChanges(0));
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
	_Check_return_ LPSRowSet HrFindUnmarkedProvider(_In_ LPSERVICEADMIN lpServiceAdmin)
	{
		if (!lpServiceAdmin) return nullptr;

		LPSRowSet lpRowSet = nullptr;
		static const SizedSPropTagArray(1, pTagUID) = {1, {PR_SERVICE_UID}};

		LPMAPITABLE lpProviderTable = nullptr;
		auto hRes = EC_MAPI(lpServiceAdmin->GetMsgServiceTable(0, &lpProviderTable));
		if (lpProviderTable)
		{
			LPPROFSECT lpSect = nullptr;
			EC_MAPI_S(lpProviderTable->SetColumns(LPSPropTagArray(&pTagUID), TBL_BATCH));
			for (;;)
			{
				hRes = EC_MAPI(lpProviderTable->QueryRows(1, 0, &lpRowSet));
				if (hRes == S_OK && lpRowSet && 1 == lpRowSet->cRows)
				{
					const auto lpCurRow = &lpRowSet->aRow[0];

					auto lpServiceUID = PpropFindProp(lpCurRow->lpProps, lpCurRow->cValues, PR_SERVICE_UID);

					if (lpServiceUID)
					{
						lpSect = OpenProfileSection(lpServiceAdmin, &mapi::getBin(lpServiceUID));
						if (lpSect)
						{
							auto pTagArray = SPropTagArray{1, {PR_MARKER}};
							ULONG ulPropVal = 0;
							LPSPropValue lpsPropVal = nullptr;
							EC_H_GETPROPS_S(lpSect->GetProps(&pTagArray, NULL, &ulPropVal, &lpsPropVal));
							if (!(strings::CheckStringProp(lpsPropVal, PROP_TYPE(PR_MARKER)) &&
								  !strcmp(lpsPropVal->Value.lpszA, MARKER_STRING)))
							{
								// got an unmarked provider - this is our hit
								// Don't free *lpRowSet - we're returning it
								MAPIFreeBuffer(lpsPropVal);
								break;
							}

							MAPIFreeBuffer(lpsPropVal);
							lpSect->Release();
							lpSect = nullptr;
						}
					}

					// go on to next one in the loop
					FreeProws(lpRowSet);
					lpRowSet = nullptr;
				}
				else
				{
					// no more hits - get out of the loop
					FreeProws(lpRowSet);
					lpRowSet = nullptr;
					break;
				}
			}

			if (lpSect) lpSect->Release();

			lpProviderTable->Release();
		}

		return lpRowSet;
	}

	_Check_return_ HRESULT HrAddServiceToProfile(
		_In_ const std::wstring& lpszServiceName, // Service Name
		_In_ ULONG_PTR ulUIParam, // hwnd for CreateMsgService
		ULONG ulFlags, // Flags for CreateMsgService
		ULONG cPropVals, // Count of properties for ConfigureMsgService
		_In_opt_ LPSPropValue lpPropVals, // Properties for ConfigureMsgService
		_In_ const std::wstring& lpszProfileName) // profile name
	{
		if (lpszServiceName.empty() || lpszProfileName.empty()) return MAPI_E_INVALID_PARAMETER;

		output::DebugPrint(
			output::dbgLevel::Generic,
			L"HrAddServiceToProfile(%ws, %ws)\n",
			lpszServiceName.c_str(),
			lpszProfileName.c_str());

		LPPROFADMIN lpProfAdmin = nullptr;
		// Connect to Profile Admin interface.
		auto hRes = EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));
		if (!lpProfAdmin) return hRes;

		LPSERVICEADMIN lpServiceAdmin = nullptr;
		hRes = EC_MAPI(lpProfAdmin->AdminServices(
			LPTSTR(strings::wstringTostring(lpszProfileName).c_str()), LPTSTR(""), 0, 0, &lpServiceAdmin));

		auto lpszServiceNameA = strings::wstringTostring(lpszServiceName);
		if (lpServiceAdmin)
		{
			MAPIUID uidService = {0};
			auto lpuidService = &uidService;

			auto lpServiceAdmin2 = mapi::safe_cast<LPSERVICEADMIN2>(lpServiceAdmin);
			if (lpServiceAdmin2)
			{
				hRes = EC_H_MSG(
					IDS_CREATEMSGSERVICEFAILED,
					lpServiceAdmin2->CreateMsgServiceEx(
						LPTSTR(lpszServiceNameA.c_str()),
						LPTSTR(lpszServiceNameA.c_str()),
						ulUIParam,
						ulFlags,
						&uidService));
			}
			else
			{
				// Only need to mark if we plan on calling ConfigureMsgService
				if (lpPropVals)
				{
					// Add a dummy prop to the current providers
					hRes = EC_H(HrMarkExistingProviders(lpServiceAdmin, true));
				}

				if (SUCCEEDED(hRes))
				{
					hRes = EC_H_MSG(
						IDS_CREATEMSGSERVICEFAILED,
						lpServiceAdmin->CreateMsgService(
							LPTSTR(lpszServiceNameA.c_str()), LPTSTR(lpszServiceNameA.c_str()), ulUIParam, ulFlags));
				}

				if (lpPropVals)
				{
					// Look for a provider without our dummy prop
					const auto lpRowSet = HrFindUnmarkedProvider(lpServiceAdmin);
					if (lpRowSet) output::outputSRowSet(output::dbgLevel::Generic, nullptr, lpRowSet, nullptr);

					// should only have one unmarked row
					if (lpRowSet && lpRowSet->cRows == 1)
					{
						const auto lpServiceUIDProp =
							PpropFindProp(lpRowSet->aRow[0].lpProps, lpRowSet->aRow[0].cValues, PR_SERVICE_UID);

						if (lpServiceUIDProp)
						{
							lpuidService = reinterpret_cast<LPMAPIUID>(mapi::getBin(lpServiceUIDProp).lpb);
						}
					}

					// Strip out the dummy prop
					hRes = EC_H(HrMarkExistingProviders(lpServiceAdmin, false));

					FreeProws(lpRowSet);
				}
			}

			if (lpPropVals)
			{
				hRes = EC_H_CANCEL(lpServiceAdmin->ConfigureMsgService(lpuidService, NULL, 0, cPropVals, lpPropVals));
			}

			if (lpServiceAdmin2) lpServiceAdmin2->Release();
			lpServiceAdmin->Release();
		}

		lpProfAdmin->Release();

		return hRes;
	}

	_Check_return_ HRESULT HrAddExchangeToProfile(
		_In_ ULONG_PTR ulUIParam, // hwnd for CreateMsgService
		_In_ const std::wstring& lpszServerName,
		_In_ const std::wstring& lpszMailboxName,
		_In_ const std::wstring& lpszProfileName)
	{
		output::DebugPrint(
			output::dbgLevel::Generic,
			L"HrAddExchangeToProfile(%ws, %ws, %ws)\n",
			lpszServerName.c_str(),
			lpszMailboxName.c_str(),
			lpszProfileName.c_str());

		if (lpszServerName.empty() || lpszMailboxName.empty() || lpszProfileName.empty())
			return MAPI_E_INVALID_PARAMETER;

#define NUMEXCHANGEPROPS 2
		SPropValue PropVal[NUMEXCHANGEPROPS];
		PropVal[0].ulPropTag = PR_PROFILE_UNRESOLVED_SERVER;
		auto lpszServerNameA = strings::wstringTostring(lpszServerName);
		PropVal[0].Value.lpszA = const_cast<LPSTR>(lpszServerNameA.c_str());
		PropVal[1].ulPropTag = PR_PROFILE_UNRESOLVED_NAME;
		auto lpszMailboxNameA = strings::wstringTostring(lpszMailboxName);
		PropVal[1].Value.lpszA = const_cast<LPSTR>(lpszMailboxNameA.c_str());
		const auto hRes = EC_H(
			HrAddServiceToProfile(L"MSEMS", ulUIParam, NULL, NUMEXCHANGEPROPS, PropVal, lpszProfileName)); // STRING_OK

		return hRes;
	}

	_Check_return_ HRESULT HrAddPSTToProfile(
		_In_ ULONG_PTR ulUIParam, // hwnd for CreateMsgService
		bool bUnicodePST,
		_In_ const std::wstring& lpszPSTPath, // PST name
		_In_ const std::wstring& lpszProfileName, // profile name
		bool bPasswordSet, // whether or not to include a password
		_In_ const std::wstring& lpszPassword) // password to include
	{
		auto hRes = S_OK;

		output::DebugPrint(
			output::dbgLevel::Generic,
			L"HrAddPSTToProfile(0x%X, %ws, %ws, 0x%X, %ws)\n",
			bUnicodePST,
			lpszPSTPath.c_str(),
			lpszProfileName.c_str(),
			bPasswordSet,
			lpszPassword.c_str());

		if (lpszPSTPath.empty() || lpszProfileName.empty()) return MAPI_E_INVALID_PARAMETER;

		auto lpszPasswordA = strings::wstringTostring(lpszPassword);
		if (bUnicodePST)
		{
			SPropValue PropVal[3];
			PropVal[0].ulPropTag = CHANGE_PROP_TYPE(PR_PST_PATH, PT_UNICODE);
			PropVal[0].Value.lpszW = const_cast<LPWSTR>(lpszPSTPath.c_str());
			PropVal[1].ulPropTag = PR_PST_CONFIG_FLAGS;
			PropVal[1].Value.ul = PST_CONFIG_UNICODE;
			PropVal[2].ulPropTag = PR_PST_PW_SZ_OLD;
			PropVal[2].Value.lpszA = const_cast<LPSTR>(lpszPasswordA.c_str());

			hRes = EC_H(HrAddServiceToProfile(
				L"MSUPST MS", ulUIParam, NULL, bPasswordSet ? 3 : 2, PropVal, lpszProfileName)); // STRING_OK
		}
		else
		{
			SPropValue PropVal[2];
			PropVal[0].ulPropTag = CHANGE_PROP_TYPE(PR_PST_PATH, PT_UNICODE);
			PropVal[0].Value.lpszW = const_cast<LPWSTR>(lpszPSTPath.c_str());
			PropVal[1].ulPropTag = PR_PST_PW_SZ_OLD;
			PropVal[1].Value.lpszA = const_cast<LPSTR>(lpszPasswordA.c_str());

			hRes = EC_H(HrAddServiceToProfile(
				L"MSPST MS", ulUIParam, NULL, bPasswordSet ? 2 : 1, PropVal, lpszProfileName)); // STRING_OK
		}

		return hRes;
	}

	// Creates an empty profile.
	_Check_return_ HRESULT HrCreateProfile(_In_ const std::wstring& lpszProfileName) // profile name
	{
		LPPROFADMIN lpProfAdmin = nullptr;

		output::DebugPrint(output::dbgLevel::Generic, L"HrCreateProfile(%ws)\n", lpszProfileName.c_str());

		if (lpszProfileName.empty()) return MAPI_E_INVALID_PARAMETER;

		// Connect to Profile Admin interface.
		auto hRes = EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));
		if (!lpProfAdmin) return hRes;

		// Create the profile
		hRes = WC_MAPI(lpProfAdmin->CreateProfile(
			LPTSTR(strings::wstringTostring(lpszProfileName).c_str()),
			nullptr,
			0,
			NULL)); // fMapiUnicode is not supported!
		if (hRes != S_OK)
		{
			// Did it fail because a profile of this name already exists?
			const auto hResCheck = HrMAPIProfileExists(lpProfAdmin, lpszProfileName);
			CHECKHRESMSG(hResCheck, IDS_DUPLICATEPROFILE);
		}

		lpProfAdmin->Release();

		return hRes;
	}

	// Removes a profile.
	_Check_return_ HRESULT HrRemoveProfile(_In_ const std::wstring& lpszProfileName)
	{
		LPPROFADMIN lpProfAdmin = nullptr;

		output::DebugPrint(output::dbgLevel::Generic, L"HrRemoveProfile(%ws)\n", lpszProfileName.c_str());
		if (lpszProfileName.empty()) return MAPI_E_INVALID_PARAMETER;

		auto hRes = EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));
		if (!lpProfAdmin) return hRes;

		hRes = EC_MAPI(lpProfAdmin->DeleteProfile(LPTSTR(strings::wstringTostring(lpszProfileName).c_str()), 0));

		lpProfAdmin->Release();

		RegFlushKey(HKEY_LOCAL_MACHINE);
		RegFlushKey(HKEY_CURRENT_USER);

		return hRes;
	}

	// Set a profile as default.
	_Check_return_ HRESULT HrSetDefaultProfile(_In_ const std::wstring& lpszProfileName)
	{
		LPPROFADMIN lpProfAdmin = nullptr;

		output::DebugPrint(output::dbgLevel::Generic, L"HrRemoveProfile(%ws)\n", lpszProfileName.c_str());
		if (lpszProfileName.empty()) return MAPI_E_INVALID_PARAMETER;

		auto hRes = EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));
		if (!lpProfAdmin) return hRes;

		hRes = EC_MAPI(lpProfAdmin->SetDefaultProfile(LPTSTR(strings::wstringTostring(lpszProfileName).c_str()), 0));

		lpProfAdmin->Release();

		RegFlushKey(HKEY_LOCAL_MACHINE);
		RegFlushKey(HKEY_CURRENT_USER);

		return hRes;
	}

	// Checks for an existing profile.
	_Check_return_ HRESULT HrMAPIProfileExists(_In_ LPPROFADMIN lpProfAdmin, _In_ const std::wstring& lpszProfileName)
	{
		LPMAPITABLE lpTable = nullptr;
		LPSRowSet lpRows = nullptr;

		static const SizedSPropTagArray(1, rgPropTag) = {1, {PR_DISPLAY_NAME_A}};

		output::DebugPrint(output::dbgLevel::Generic, L"HrMAPIProfileExists()\n");
		if (!lpProfAdmin || lpszProfileName.empty()) return MAPI_E_INVALID_PARAMETER;

		// Get a table of existing profiles

		auto hRes = EC_MAPI(lpProfAdmin->GetProfileTable(0, &lpTable));
		if (!lpTable) return hRes;

		hRes = EC_MAPI(HrQueryAllRows(lpTable, LPSPropTagArray(&rgPropTag), nullptr, nullptr, 0, &lpRows));

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

				if (SUCCEEDED(hRes))
				{
					auto lpszProfileNameA = strings::wstringTostring(lpszProfileName);
					for (ULONG i = 0; i < lpRows->cRows; i++)
					{
						const auto lpProp = lpRows->aRow[i].lpProps;

						const auto ulComp = EC_D(
							ULONG,
							CompareStringA(
								g_lcid, // LOCALE_INVARIANT,
								NORM_IGNORECASE,
								lpProp[0].Value.lpszA,
								-1,
								lpszProfileNameA.c_str(),
								-1));

						if (ulComp == CSTR_EQUAL)
						{
							hRes = E_ACCESSDENIED;
							break;
						}
					}
				}
			}
		}

		if (lpRows) FreeProws(lpRows);

		lpTable->Release();
		return hRes;
	}

	_Check_return_ HRESULT GetProfileServiceVersion(
		_In_ const std::wstring& lpszProfileName,
		_Out_ ULONG* lpulServerVersion,
		_Out_ EXCHANGE_STORE_VERSION_NUM* lpStoreVersion,
		_Out_ bool* lpbFoundServerVersion,
		_Out_ bool* lpbFoundServerFullVersion)
	{
		if (lpszProfileName.empty() || !lpulServerVersion || !lpStoreVersion || !lpbFoundServerVersion ||
			!lpbFoundServerFullVersion)
			return MAPI_E_INVALID_PARAMETER;
		*lpulServerVersion = NULL;
		memset(lpStoreVersion, 0, sizeof(EXCHANGE_STORE_VERSION_NUM));
		*lpbFoundServerVersion = false;
		*lpbFoundServerFullVersion = false;

		LPPROFADMIN lpProfAdmin = nullptr;
		LPSERVICEADMIN lpServiceAdmin = nullptr;

		output::DebugPrint(output::dbgLevel::Generic, L"GetProfileServiceVersion(%ws)\n", lpszProfileName.c_str());

		auto hRes = EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));
		if (!lpProfAdmin) return hRes;

		hRes = EC_MAPI(lpProfAdmin->AdminServices(
			LPTSTR(strings::wstringTostring(lpszProfileName).c_str()), LPTSTR(""), 0, 0, &lpServiceAdmin));

		if (lpServiceAdmin)
		{
			LPPROFSECT lpProfSect = nullptr;
			hRes = EC_MAPI(
				lpServiceAdmin->OpenProfileSection(LPMAPIUID(pbGlobalProfileSectionGuid), nullptr, 0, &lpProfSect));
			if (lpProfSect)
			{
				LPSPropValue lpServerVersion = nullptr;
				hRes = WC_MAPI(HrGetOneProp(lpProfSect, PR_PROFILE_SERVER_VERSION, &lpServerVersion));

				if (SUCCEEDED(hRes) && lpServerVersion && PR_PROFILE_SERVER_VERSION == lpServerVersion->ulPropTag)
				{
					*lpbFoundServerVersion = true;
					*lpulServerVersion = lpServerVersion->Value.l;
				}
				MAPIFreeBuffer(lpServerVersion);

				LPSPropValue lpServerFullVersion = nullptr;
				hRes = WC_MAPI(HrGetOneProp(lpProfSect, PR_PROFILE_SERVER_FULL_VERSION, &lpServerFullVersion));

				if (SUCCEEDED(hRes) && lpServerFullVersion &&
					PR_PROFILE_SERVER_FULL_VERSION == lpServerFullVersion->ulPropTag &&
					sizeof(EXCHANGE_STORE_VERSION_NUM) == mapi::getBin(lpServerFullVersion).cb)
				{
					const auto bin = mapi::getBin(lpServerFullVersion);
					output::DebugPrint(output::dbgLevel::Generic, L"PR_PROFILE_SERVER_FULL_VERSION = ");
					output::outputBinary(output::dbgLevel::Generic, nullptr, bin);
					output::DebugPrint(output::dbgLevel::Generic, L"\n");

					memcpy(lpStoreVersion, bin.lpb, sizeof(EXCHANGE_STORE_VERSION_NUM));
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
	_Check_return_ HRESULT
	HrCopyProfile(_In_ const std::wstring& lpszOldProfileName, _In_ const std::wstring& lpszNewProfileName)
	{
		LPPROFADMIN lpProfAdmin = nullptr;

		output::DebugPrint(
			output::dbgLevel::Generic,
			L"HrCopyProfile(%ws, %ws)\n",
			lpszOldProfileName.c_str(),
			lpszNewProfileName.c_str());
		if (lpszOldProfileName.empty() || lpszNewProfileName.empty()) return MAPI_E_INVALID_PARAMETER;

		auto hRes = EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));
		if (!lpProfAdmin) return hRes;

		hRes = EC_MAPI(lpProfAdmin->CopyProfile(
			LPTSTR(strings::wstringTostring(lpszOldProfileName).c_str()),
			nullptr,
			LPTSTR(strings::wstringTostring(lpszNewProfileName).c_str()),
			NULL,
			NULL));

		lpProfAdmin->Release();

		return hRes;
	}

	_Check_return_ LPPROFSECT OpenProfileSection(_In_ LPSERVICEADMIN lpServiceAdmin, _In_ const SBinary* lpServiceUID)
	{
		if (!lpServiceUID || !lpServiceAdmin) return nullptr;

		output::DebugPrint(output::dbgLevel::OpenItemProp, L"OpenProfileSection opening lpServiceUID = ");
		output::outputBinary(output::dbgLevel::OpenItemProp, nullptr, *lpServiceUID);
		output::DebugPrint(output::dbgLevel::OpenItemProp, L"\n");

		LPPROFSECT lpProfSect = nullptr;
		// First, we try the normal way of opening the profile section:
		WC_MAPI_S(lpServiceAdmin->OpenProfileSection(
			reinterpret_cast<LPMAPIUID>(lpServiceUID->lpb),
			nullptr,
			MAPI_MODIFY | MAPI_FORCE_ACCESS, // passing this flag might actually work with Outlook 2000 and XP
			&lpProfSect));
		if (!lpProfSect)
		{
			///////////////////////////////////////////////////////////////////
			// HACK CENTRAL. This is a MAJOR hack. MAPI will always return E_ACCESSDENIED
			// when we open a profile section on the service if we are a client. The workaround
			// (HACK) is to call into one of MAPI's internal functions that bypasses
			// the security check. We build a Interface to it and then point to it from our
			// offset of 0x48. USE AT YOUR OWN RISK! NOT SUPPORTED!
			interface IOpenSectionHack : IUnknown
			{
				virtual HRESULT STDMETHODCALLTYPE OpenSection(LPMAPIUID, ULONG, LPPROFSECT*) = 0;
			};

			const auto ppProfile =
				reinterpret_cast<IOpenSectionHack**>((reinterpret_cast<BYTE*>(lpServiceAdmin) + 0x48));

			// Now, we want to get open the Services Profile Section and store that
			// interface with the Object

			if (ppProfile && *ppProfile)
			{
				EC_MAPI_S((*ppProfile)
							  ->OpenSection(reinterpret_cast<LPMAPIUID>(lpServiceUID->lpb), MAPI_MODIFY, &lpProfSect));
			}
			// END OF HACK. I'm amazed that this works....
			///////////////////////////////////////////////////////////////////
		}

		return lpProfSect;
	}

	_Check_return_ LPPROFSECT
	OpenProfileSection(_In_ LPPROVIDERADMIN lpProviderAdmin, _In_ const SBinary* lpProviderUID)
	{
		if (!lpProviderUID || !lpProviderAdmin) return nullptr;

		output::DebugPrint(output::dbgLevel::OpenItemProp, L"OpenProfileSection opening lpServiceUID = ");
		output::outputBinary(output::dbgLevel::OpenItemProp, nullptr, *lpProviderUID);
		output::DebugPrint(output::dbgLevel::OpenItemProp, L"\n");

		LPPROFSECT lpProfSect = nullptr;
		WC_MAPI_S(lpProviderAdmin->OpenProfileSection(
			reinterpret_cast<LPMAPIUID>(lpProviderUID->lpb), nullptr, MAPI_MODIFY | MAPI_FORCE_ACCESS, &lpProfSect));
		if (!lpProfSect)
		{
			// We only do this hack as a last resort - it can crash some versions of Outlook, but is required for Exchange
			*(reinterpret_cast<BYTE*>(lpProviderAdmin) + 0x60) = 0x2; // Use at your own risk! NOT SUPPORTED!

			WC_MAPI_S(lpProviderAdmin->OpenProfileSection(
				reinterpret_cast<LPMAPIUID>(lpProviderUID->lpb), nullptr, MAPI_MODIFY, &lpProfSect));
		}

		return lpProfSect;
	}
} // namespace mapi::profile