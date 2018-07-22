// Collection of useful MAPI Store functions

#include <StdAfx.h>
#include <MAPI/MAPIStoreFunctions.h>
#include <MAPI/MAPIFunctions.h>
#include <Interpret/String.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <Interpret/Guids.h>
#include <Interpret/ExtraPropTags.h>
#ifndef MRMAPI
#include <MAPI/MAPIABFunctions.h>
#include <Interpret/InterpretProp.h>
#endif

namespace mapi
{
	namespace store
	{
		_Check_return_ LPMDB
		CallOpenMsgStore(_In_ LPMAPISESSION lpSession, _In_ ULONG_PTR ulUIParam, _In_ LPSBinary lpEID, ULONG ulFlags)
		{
			if (!lpSession || !lpEID) return nullptr;

			if (registry::RegKeys[registry::regkeyMDB_ONLINE].ulCurDWORD)
			{
				ulFlags |= MDB_ONLINE;
			}

			output::DebugPrint(DBGOpenItemProp, L"CallOpenMsgStore ulFlags = 0x%X\n", ulFlags);

			LPMDB lpMDB = nullptr;
			auto hRes = WC_MAPI(lpSession->OpenMsgStore(
				ulUIParam,
				lpEID->cb,
				reinterpret_cast<LPENTRYID>(lpEID->lpb),
				nullptr,
				ulFlags,
				static_cast<LPMDB*>(&lpMDB)));
			if (hRes == MAPI_E_UNKNOWN_FLAGS && ulFlags & MDB_ONLINE)
			{
				// Perhaps this store doesn't know the MDB_ONLINE flag - remove and retry
				ulFlags = ulFlags & ~MDB_ONLINE;
				output::DebugPrint(DBGOpenItemProp, L"CallOpenMsgStore 2nd attempt ulFlags = 0x%X\n", ulFlags);

				WC_MAPI_S(lpSession->OpenMsgStore(
					ulUIParam,
					lpEID->cb,
					reinterpret_cast<LPENTRYID>(lpEID->lpb),
					nullptr,
					ulFlags,
					static_cast<LPMDB*>(&lpMDB)));
			}

			return lpMDB;
		}

		// Build a server DN.
		std::string BuildServerDN(const std::string& szServerName, const std::string& szPost)
		{
			static std::string szPre = "/cn=Configuration/cn=Servers/cn="; // STRING_OK
			return szPre + szServerName + szPost;
		}

		_Check_return_ HRESULT GetMailboxTable1(
			_In_ LPMDB lpMDB,
			const std::string& szServerDN,
			ULONG ulFlags,
			_Deref_out_opt_ LPMAPITABLE* lpMailboxTable)
		{
			if (!lpMDB || !lpMailboxTable || szServerDN.empty()) return MAPI_E_INVALID_PARAMETER;
			*lpMailboxTable = nullptr;

			auto hRes = S_OK;
			auto lpManageStore1 = mapi::safe_cast<LPEXCHANGEMANAGESTORE>(lpMDB);
			if (lpManageStore1)
			{
				hRes = WC_MAPI(lpManageStore1->GetMailboxTable(LPSTR(szServerDN.c_str()), lpMailboxTable, ulFlags));

				lpManageStore1->Release();
			}

			return hRes;
		}

		_Check_return_ HRESULT GetMailboxTable3(
			_In_ LPMDB lpMDB,
			const std::string& szServerDN,
			ULONG ulOffset,
			ULONG ulFlags,
			_Deref_out_opt_ LPMAPITABLE* lpMailboxTable)
		{
			if (!lpMDB || !lpMailboxTable || szServerDN.empty()) return MAPI_E_INVALID_PARAMETER;
			*lpMailboxTable = nullptr;

			auto hRes = S_OK;
			auto lpManageStore3 = mapi::safe_cast<LPEXCHANGEMANAGESTORE3>(lpMDB);
			if (lpManageStore3)
			{
				hRes = WC_MAPI(lpManageStore3->GetMailboxTableOffset(
					LPSTR(szServerDN.c_str()), lpMailboxTable, ulFlags, ulOffset));

				lpManageStore3->Release();
			}

			return hRes;
		}

		_Check_return_ HRESULT GetMailboxTable5(
			_In_ LPMDB lpMDB,
			const std::string& szServerDN,
			ULONG ulOffset,
			ULONG ulFlags,
			_In_opt_ LPGUID lpGuidMDB,
			_Deref_out_opt_ LPMAPITABLE* lpMailboxTable)
		{
			if (!lpMDB || !lpMailboxTable || szServerDN.empty()) return MAPI_E_INVALID_PARAMETER;
			*lpMailboxTable = nullptr;

			auto hRes = S_OK;
			auto lpManageStore5 = mapi::safe_cast<LPEXCHANGEMANAGESTORE5>(lpMDB);
			if (lpManageStore5)
			{
				hRes = EC_MAPI(lpManageStore5->GetMailboxTableEx(
					LPSTR(szServerDN.c_str()), lpGuidMDB, lpMailboxTable, ulFlags, ulOffset));

				lpManageStore5->Release();
			}

			return hRes;
		}

		// lpMDB needs to be an Exchange MDB - OpenMessageStoreGUID(pbExchangeProviderPrimaryUserGuid) can get one if there's one to be had
		// Use GetServerName to get the default server
		// Will try IID_IExchangeManageStore3 first and fail back to IID_IExchangeManageStore
		_Check_return_ HRESULT GetMailboxTable(
			_In_ LPMDB lpMDB,
			const std::string& szServerName,
			ULONG ulOffset,
			_Deref_out_opt_ LPMAPITABLE* lpMailboxTable)
		{
			if (!lpMDB || !StoreSupportsManageStore(lpMDB) || !lpMailboxTable) return MAPI_E_INVALID_PARAMETER;
			*lpMailboxTable = nullptr;

			auto hRes = S_OK;
			LPMAPITABLE lpLocalTable = nullptr;

			auto szServerDN = BuildServerDN(szServerName, "");
			if (!szServerDN.empty())
			{
				hRes = WC_H(GetMailboxTable3(lpMDB, szServerDN, ulOffset, fMapiUnicode, &lpLocalTable));

				if (!lpLocalTable && 0 == ulOffset)
				{
					hRes = WC_H(GetMailboxTable1(lpMDB, szServerDN, fMapiUnicode, &lpLocalTable));
				}
			}

			*lpMailboxTable = lpLocalTable;
			return hRes;
		}

		_Check_return_ HRESULT GetPublicFolderTable1(
			_In_ LPMDB lpMDB,
			const std::string& szServerDN,
			ULONG ulFlags,
			_Deref_out_opt_ LPMAPITABLE* lpPFTable)
		{
			if (!lpMDB || !lpPFTable || szServerDN.empty()) return MAPI_E_INVALID_PARAMETER;
			*lpPFTable = nullptr;

			auto hRes = S_OK;
			auto lpManageStore1 = mapi::safe_cast<LPEXCHANGEMANAGESTORE>(lpMDB);
			if (lpManageStore1)
			{
				hRes = EC_MAPI(lpManageStore1->GetPublicFolderTable(LPSTR(szServerDN.c_str()), lpPFTable, ulFlags));

				lpManageStore1->Release();
			}

			return hRes;
		}

		_Check_return_ HRESULT GetPublicFolderTable4(
			_In_ LPMDB lpMDB,
			const std::string& szServerDN,
			ULONG ulOffset,
			ULONG ulFlags,
			_Deref_out_opt_ LPMAPITABLE* lpPFTable)
		{
			if (!lpMDB || !lpPFTable || szServerDN.empty()) return MAPI_E_INVALID_PARAMETER;
			*lpPFTable = nullptr;

			auto hRes = S_OK;
			auto lpManageStore4 = mapi::safe_cast<LPEXCHANGEMANAGESTORE4>(lpMDB);
			if (lpManageStore4)
			{
				hRes = EC_MAPI(lpManageStore4->GetPublicFolderTableOffset(
					LPSTR(szServerDN.c_str()), lpPFTable, ulFlags, ulOffset));
				lpManageStore4->Release();
			}

			return hRes;
		}

		_Check_return_ HRESULT GetPublicFolderTable5(
			_In_ LPMDB lpMDB,
			const std::string& szServerDN,
			ULONG ulOffset,
			ULONG ulFlags,
			_In_opt_ LPGUID lpGuidMDB,
			_Deref_out_opt_ LPMAPITABLE* lpPFTable)
		{
			if (!lpMDB || !lpPFTable || szServerDN.empty()) return MAPI_E_INVALID_PARAMETER;
			*lpPFTable = nullptr;

			auto hRes = S_OK;
			auto lpManageStore5 = mapi::safe_cast<LPEXCHANGEMANAGESTORE5>(lpMDB);
			if (lpManageStore5)
			{
				hRes = EC_MAPI(lpManageStore5->GetPublicFolderTableEx(
					LPSTR(szServerDN.c_str()), lpGuidMDB, lpPFTable, ulFlags, ulOffset));

				lpManageStore5->Release();
			}

			return hRes;
		}

		// Get server name from the profile
		std::string GetServerName(_In_ LPMAPISESSION lpSession)
		{
			LPSERVICEADMIN pSvcAdmin = nullptr;
			LPPROFSECT pGlobalProfSect = nullptr;
			LPSPropValue lpServerName = nullptr;
			std::string serverName;

			if (!lpSession) return "";

			auto hRes = EC_MAPI(lpSession->AdminServices(0, &pSvcAdmin));

			if (SUCCEEDED(hRes))
			{
				hRes = EC_MAPI(
					pSvcAdmin->OpenProfileSection(LPMAPIUID(pbGlobalProfileSectionGuid), nullptr, 0, &pGlobalProfSect));
			}

			if (SUCCEEDED(hRes))
			{
				WC_MAPI_S(HrGetOneProp(pGlobalProfSect, PR_PROFILE_HOME_SERVER, &lpServerName));
			}

			if (mapi::CheckStringProp(lpServerName, PT_STRING8)) // profiles are ASCII only
			{
				serverName = lpServerName->Value.lpszA;
			}
#ifndef MRMAPI
			else
			{
				// prompt the user to enter a server name
				dialog::editor::CEditor MyData(
					nullptr, IDS_SERVERNAME, IDS_SERVERNAMEMISSINGPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
				MyData.InitPane(0, viewpane::TextPane::CreateSingleLinePane(IDS_SERVERNAME, false));

				if (MyData.DisplayDialog())
				{
					serverName = strings::wstringTostring(MyData.GetStringW(0));
				}
			}
#endif
			MAPIFreeBuffer(lpServerName);
			if (pGlobalProfSect) pGlobalProfSect->Release();
			if (pSvcAdmin) pSvcAdmin->Release();
			return serverName;
		}

		_Check_return_ HRESULT CreateStoreEntryID(
			_In_ LPMDB lpMDB, // open message store
			const std::string& lpszMsgStoreDN, // desired message store DN
			const std::string& lpszMailboxDN, // desired mailbox DN or NULL
			ULONG ulFlags, // desired flags for CreateStoreEntryID
			_Out_opt_ ULONG* lpcbEntryID,
			_Deref_out_opt_ LPENTRYID* lppEntryID)
		{
			auto hRes = S_OK;

			if (!lpMDB || lpszMsgStoreDN.empty() || !StoreSupportsManageStore(lpMDB))
			{
				if (!lpMDB) output::DebugPrint(DBGGeneric, L"CreateStoreEntryID: MDB was NULL\n");
				if (lpszMsgStoreDN.empty())
					output::DebugPrint(DBGGeneric, L"CreateStoreEntryID: lpszMsgStoreDN was missing\n");
				return MAPI_E_INVALID_PARAMETER;
			}

			auto lpXManageStore = mapi::safe_cast<LPEXCHANGEMANAGESTORE>(lpMDB);
			if (lpXManageStore)
			{
				output::DebugPrint(
					DBGGeneric,
					L"CreateStoreEntryID: Creating EntryID. StoreDN = \"%hs\", MailboxDN = \"%hs\", Flags = \"0x%X\"\n",
					lpszMsgStoreDN.c_str(),
					lpszMailboxDN.c_str(),
					ulFlags);

				hRes = EC_MAPI(lpXManageStore->CreateStoreEntryID(
					LPSTR(lpszMsgStoreDN.c_str()),
					lpszMailboxDN.empty() ? nullptr : LPSTR(lpszMailboxDN.c_str()),
					ulFlags,
					lpcbEntryID,
					lppEntryID));

				lpXManageStore->Release();
			}

			return hRes;
		}

		_Check_return_ HRESULT CreateStoreEntryID2(
			_In_ LPMDB lpMDB, // open message store
			const std::string& lpszMsgStoreDN, // desired message store DN
			const std::string& lpszMailboxDN, // desired mailbox DN or NULL
			const std::wstring& smtpAddress,
			ULONG ulFlags, // desired flags for CreateStoreEntryID
			_Out_opt_ ULONG* lpcbEntryID,
			_Deref_out_opt_ LPENTRYID* lppEntryID)
		{
			auto hRes = S_OK;

			if (!lpMDB || lpszMsgStoreDN.empty() || !StoreSupportsManageStoreEx(lpMDB))
			{
				if (!lpMDB) output::DebugPrint(DBGGeneric, L"CreateStoreEntryID2: MDB was NULL\n");
				if (lpszMsgStoreDN.empty())
					output::DebugPrint(DBGGeneric, L"CreateStoreEntryID2: lpszMsgStoreDN was missing\n");
				return MAPI_E_INVALID_PARAMETER;
			}

			auto lpXManageStoreEx = mapi::safe_cast<LPEXCHANGEMANAGESTOREEX>(lpMDB);
			if (lpXManageStoreEx)
			{
				output::DebugPrint(
					DBGGeneric,
					L"CreateStoreEntryID2: Creating EntryID. StoreDN = \"%hs\", MailboxDN = \"%hs\", SmtpAddress = "
					L"\"%ws\", Flags = \"0x%X\"\n",
					lpszMsgStoreDN.c_str(),
					lpszMailboxDN.c_str(),
					smtpAddress.c_str(),
					ulFlags);
				SPropValue sProps[4] = {0};
				sProps[0].ulPropTag = PR_PROFILE_MAILBOX;
				sProps[0].Value.lpszA = LPSTR(lpszMailboxDN.c_str());

				sProps[1].ulPropTag = PR_PROFILE_MDB_DN;
				sProps[1].Value.lpszA = LPSTR(lpszMsgStoreDN.c_str());

				sProps[2].ulPropTag = PR_FORCE_USE_ENTRYID_SERVER;
				sProps[2].Value.b = true;

				sProps[3].ulPropTag = PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W;
				sProps[3].Value.lpszW = smtpAddress.empty() ? nullptr : const_cast<LPWSTR>(smtpAddress.c_str());

				hRes = EC_MAPI(lpXManageStoreEx->CreateStoreEntryID2(
					smtpAddress.empty() ? _countof(sProps) - 1 : _countof(sProps),
					reinterpret_cast<LPSPropValue>(&sProps),
					ulFlags,
					lpcbEntryID,
					lppEntryID));

				lpXManageStoreEx->Release();
			}

			return hRes;
		}

		_Check_return_ HRESULT CreateStoreEntryID(
			_In_ LPMDB lpMDB, // open message store
			const std::string& lpszMsgStoreDN, // desired message store DN
			const std::string& lpszMailboxDN, // desired mailbox DN or NULL
			const std::wstring& smtpAddress,
			ULONG ulFlags, // desired flags for CreateStoreEntryID
			bool bForceServer, // Use CreateStoreEntryID2
			_Out_opt_ ULONG* lpcbEntryID,
			_Deref_out_opt_ LPENTRYID* lppEntryID)
		{
			auto hRes = S_OK;

			// Use an empty MailboxDN to open the public store
			if (lpszMailboxDN.empty())
			{
				ulFlags |= OPENSTORE_PUBLIC;
			}

			if (!bForceServer)
			{
				hRes =
					EC_MAPI(CreateStoreEntryID(lpMDB, lpszMsgStoreDN, lpszMailboxDN, ulFlags, lpcbEntryID, lppEntryID));
			}
			else
			{
				hRes = EC_MAPI(CreateStoreEntryID2(
					lpMDB, lpszMsgStoreDN, lpszMailboxDN, smtpAddress, ulFlags, lpcbEntryID, lppEntryID));
			}

			return hRes;
		}

		// Stolen from MBLogon.c in the EDK to avoid compiling and linking in the entire beast
		// Cleaned up to fit in with other functions
		// $--HrMailboxLogon------------------------------------------------------
		// Logon to a mailbox. Before calling this function do the following:
		// 1) Create a profile that has Exchange administrator privileges.
		// 2) Logon to Exchange using this profile.
		// 3) Open the mailbox using the Message Store DN and Mailbox DN.
		//
		// This version of the function needs the server and mailbox names to be
		// in the form of distinguished names. They would look something like this:
		// /CN=Configuration/CN=Servers/CN=%s/CN=Microsoft Private MDB
		// /CN=Configuration/CN=Servers/CN=%s/CN=Microsoft Public MDB
		// /O=<Organization>/OU=<Site>/CN=<Container>/CN=<MailboxName>
		// where items in <brackets> would need to be set to appropriate values
		//
		// Note1: The message store DN is nearly identical to the server DN, except
		// for the addition of a trailing '/CN=' part. This part is required although
		// its actual value is ignored.
		//
		// Note2: A NULL lpszMailboxDN indicates the public store should be opened.
		// -----------------------------------------------------------------------------

		_Check_return_ HRESULT HrMailboxLogon(
			_In_ LPMAPISESSION lpMAPISession, // MAPI session handle
			_In_ LPMDB lpMDB, // open message store
			const std::string& lpszMsgStoreDN, // desired message store DN
			const std::string& lpszMailboxDN, // desired mailbox DN or NULL
			const std::wstring& smtpAddress,
			ULONG ulFlags, // desired flags for CreateStoreEntryID
			bool bForceServer, // Use CreateStoreEntryID2
			_Deref_out_opt_ LPMDB* lppMailboxMDB) // ptr to mailbox message store ptr
		{
			SBinary sbEID = {0};

			*lppMailboxMDB = nullptr;

			if (!lpMAPISession)
			{
				output::DebugPrint(DBGGeneric, L"HrMailboxLogon: Session was NULL\n");
				return MAPI_E_INVALID_PARAMETER;
			}

			auto hRes = WC_H(CreateStoreEntryID(
				lpMDB,
				lpszMsgStoreDN,
				lpszMailboxDN,
				smtpAddress,
				ulFlags,
				bForceServer,
				&sbEID.cb,
				reinterpret_cast<LPENTRYID*>(&sbEID.lpb)));

			if (SUCCEEDED(hRes))
			{
				*lppMailboxMDB = CallOpenMsgStore(
					lpMAPISession,
					NULL,
					&sbEID,
					MDB_NO_DIALOG | MDB_NO_MAIL | // spooler not notified of our presence
						MDB_TEMPORARY | // message store not added to MAPI profile
						MAPI_BEST_ACCESS); // normally WRITE, but allow access to RO store
			}

			MAPIFreeBuffer(sbEID.lpb);
			return hRes;
		}

		_Check_return_ HRESULT
		OpenDefaultMessageStore(_In_ LPMAPISESSION lpMAPISession, _Deref_out_ LPMDB* lppDefaultMDB)
		{
			LPMAPITABLE pStoresTbl = nullptr;
			LPSRowSet pRow = nullptr;
			static SRestriction sres;
			SPropValue spv;

			enum
			{
				EID,
				NUM_COLS
			};
			static const SizedSPropTagArray(NUM_COLS, sptEIDCol) = {
				NUM_COLS,
				{PR_ENTRYID},
			};
			if (!lpMAPISession) return MAPI_E_INVALID_PARAMETER;

			auto hRes = EC_MAPI(lpMAPISession->GetMsgStoresTable(0, &pStoresTbl));

			// set up restriction for the default store
			sres.rt = RES_PROPERTY; // gonna compare a property
			sres.res.resProperty.relop = RELOP_EQ; // gonna test equality
			sres.res.resProperty.ulPropTag = PR_DEFAULT_STORE; // tag to compare
			sres.res.resProperty.lpProp = &spv; // prop tag to compare against

			spv.ulPropTag = PR_DEFAULT_STORE; // tag type
			spv.Value.b = true; // tag value

			EC_MAPI_S(HrQueryAllRows(
				pStoresTbl, // table to query
				LPSPropTagArray(&sptEIDCol), // columns to get
				&sres, // restriction to use
				nullptr, // sort order
				0, // max number of rows - 0 means ALL
				&pRow));

			if (pRow && pRow->cRows)
			{
				*lppDefaultMDB =
					CallOpenMsgStore(lpMAPISession, NULL, &pRow->aRow[0].lpProps[EID].Value.bin, MDB_WRITE);
			}
			else
				hRes = MAPI_E_NOT_FOUND;

			if (pRow) FreeProws(pRow);
			if (pStoresTbl) pStoresTbl->Release();
			return hRes;
		}

		// Build DNs for call to HrMailboxLogon
		_Check_return_ LPMDB OpenOtherUsersMailbox(
			_In_ LPMAPISESSION lpMAPISession,
			_In_ LPMDB lpMDB,
			const std::string& szServerName,
			const std::string& szMailboxDN,
			const std::wstring& smtpAddress,
			ULONG ulFlags, // desired flags for CreateStoreEntryID
			bool bForceServer) // Use CreateStoreEntryID2
		{
			if (!lpMAPISession || !lpMDB || szMailboxDN.empty() || !StoreSupportsManageStore(lpMDB)) return nullptr;

			output::DebugPrint(
				DBGGeneric,
				L"OpenOtherUsersMailbox called with lpMAPISession = %p, lpMDB = %p, Server = \"%hs\", Mailbox = "
				L"\"%hs\", SmtpAddress = \"%ws\"\n",
				lpMAPISession,
				lpMDB,
				szServerName.c_str(),
				szMailboxDN.c_str(),
				smtpAddress.c_str());

			std::string serverName;
			if (szServerName.empty())
			{
				// If we weren't given a server name, get one from the profile
				serverName = GetServerName(lpMAPISession);
			}
			else
			{
				serverName = szServerName;
			}

			LPMDB lpOtherUserMDB = nullptr;
			if (!serverName.empty())
			{
				auto szServerDN = BuildServerDN(serverName,
												"/cn=Microsoft Private MDB"); // STRING_OK

				if (!szServerDN.empty())
				{
					output::DebugPrint(
						DBGGeneric, L"Calling HrMailboxLogon with Server DN = \"%hs\"\n", szServerDN.c_str());
					WC_H_S(HrMailboxLogon(
						lpMAPISession,
						lpMDB,
						szServerDN,
						szMailboxDN,
						smtpAddress,
						ulFlags,
						bForceServer,
						&lpOtherUserMDB));
				}
			}

			return lpOtherUserMDB;
		}

#ifndef MRMAPI
		_Check_return_ LPMDB OpenMailboxWithPrompt(
			_In_ LPMAPISESSION lpMAPISession,
			_In_ LPMDB lpMDB,
			const std::string& szServerName,
			const std::wstring& szMailboxDN,
			ULONG ulFlags) // desired flags for CreateStoreEntryID
		{
			if (!lpMAPISession) return nullptr;

			dialog::editor::CEditor MyPrompt(
				nullptr, IDS_OPENOTHERUSER, IDS_OPENWITHFLAGSPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyPrompt.SetPromptPostFix(interpretprop::AllFlagsToString(PROP_ID(PR_PROFILE_OPEN_FLAGS), true));
			MyPrompt.InitPane(
				0,
				viewpane::TextPane::CreateSingleLinePane(
					IDS_SERVERNAME, strings::stringTowstring(szServerName), false));
			MyPrompt.InitPane(1, viewpane::TextPane::CreateSingleLinePane(IDS_USERDN, szMailboxDN, false));
			MyPrompt.InitPane(2, viewpane::TextPane::CreateSingleLinePane(IDS_USER_SMTP_ADDRESS, false));
			MyPrompt.InitPane(3, viewpane::TextPane::CreateSingleLinePane(IDS_CREATESTORENTRYIDFLAGS, false));
			MyPrompt.SetHex(3, ulFlags);
			MyPrompt.InitPane(4, viewpane::CheckPane::Create(IDS_FORCESERVER, false, false));
			if (!MyPrompt.DisplayDialog()) return nullptr;

			return OpenOtherUsersMailbox(
				lpMAPISession,
				lpMDB,
				strings::wstringTostring(MyPrompt.GetStringW(0)),
				strings::wstringTostring(MyPrompt.GetStringW(1)),
				MyPrompt.GetStringW(2),
				MyPrompt.GetHex(3),
				MyPrompt.GetCheck(4));
		}

		// Display a UI to select a mailbox, then call OpenOtherUsersMailbox with the mailboxDN
		// May return MAPI_E_CANCEL
		_Check_return_ LPMDB OpenOtherUsersMailboxFromGal(_In_ LPMAPISESSION lpMAPISession, _In_ LPADRBOOK lpAddrBook)
		{
			if (!lpMAPISession || !lpAddrBook) return nullptr;

			LPMDB lpOtherUserMDB = nullptr;

			const auto szServerName = GetServerName(lpMAPISession);

			auto lpPrivateMDB = OpenMessageStoreGUID(lpMAPISession, pbExchangeProviderPrimaryUserGuid);
			if (lpPrivateMDB && StoreSupportsManageStore(lpPrivateMDB))
			{
				auto lpMailUser = ab::SelectUser(lpAddrBook, ::GetDesktopWindow(), nullptr);
				if (lpMailUser)
				{
					LPSPropValue lpEmailAddress = nullptr;
					EC_MAPI_S(HrGetOneProp(lpMailUser, PR_EMAIL_ADDRESS_W, &lpEmailAddress));
					if (mapi::CheckStringProp(lpEmailAddress, PT_UNICODE))
					{
						lpOtherUserMDB = OpenMailboxWithPrompt(
							lpMAPISession,
							lpPrivateMDB,
							szServerName,
							lpEmailAddress->Value.lpszW,
							OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_TAKE_OWNERSHIP);
					}

					MAPIFreeBuffer(lpEmailAddress);
					lpMailUser->Release();
				}
			}

			if (lpPrivateMDB) lpPrivateMDB->Release();
			return lpOtherUserMDB;
		}
#endif

		// Use these guids:
		// pbExchangeProviderPrimaryUserGuid
		// pbExchangeProviderDelegateGuid
		// pbExchangeProviderPublicGuid
		// pbExchangeProviderXportGuid
		_Check_return_ LPMDB OpenMessageStoreGUID(
			_In_ LPMAPISESSION lpMAPISession,
			_In_z_ LPCSTR lpGUID) // Do not migrate this to wstring/string
		{
			if (!lpMAPISession) return nullptr;

			LPMAPITABLE pStoresTbl = nullptr;
			LPMDB lpMDB = nullptr;

			enum
			{
				EID,
				STORETYPE,
				NUM_COLS
			};
			static const SizedSPropTagArray(NUM_COLS, sptCols) = {NUM_COLS, {PR_ENTRYID, PR_MDB_PROVIDER}};

			EC_MAPI_S(lpMAPISession->GetMsgStoresTable(0, &pStoresTbl));
			if (pStoresTbl)
			{
				LPSRowSet pRow = nullptr;
				EC_MAPI_S(HrQueryAllRows(
					pStoresTbl, // table to query
					LPSPropTagArray(&sptCols), // columns to get
					nullptr, // restriction to use
					nullptr, // sort order
					0, // max number of rows
					&pRow));
				if (pRow)
				{
					for (ULONG ulRowNum = 0; ulRowNum < pRow->cRows; ulRowNum++)
					{
						// check to see if we have a folder with a matching GUID
						if (pRow->aRow[ulRowNum].lpProps[STORETYPE].ulPropTag == PR_MDB_PROVIDER &&
							pRow->aRow[ulRowNum].lpProps[EID].ulPropTag == PR_ENTRYID &&
							IsEqualMAPIUID(pRow->aRow[ulRowNum].lpProps[STORETYPE].Value.bin.lpb, lpGUID))
						{
							lpMDB = CallOpenMsgStore(
								lpMAPISession, NULL, &pRow->aRow[ulRowNum].lpProps[EID].Value.bin, MDB_WRITE);
							break;
						}
					}

					FreeProws(pRow);
				}

				pStoresTbl->Release();
			}

			return lpMDB;
		}

		_Check_return_ LPMDB OpenPublicMessageStore(
			_In_ LPMAPISESSION lpMAPISession,
			const std::string& szServerName,
			ULONG ulFlags) // Flags for CreateStoreEntryID
		{
			if (!lpMAPISession) return nullptr;

			auto lpPublicMDBNonAdmin = OpenMessageStoreGUID(lpMAPISession, pbExchangeProviderPublicGuid);

			// If we don't have flags we're done
			if (!ulFlags)
			{
				return lpPublicMDBNonAdmin;
			}

			LPMDB lpPublicMDB = nullptr;
			if (lpPublicMDBNonAdmin && StoreSupportsManageStore(lpPublicMDBNonAdmin))
			{
				LPSPropValue lpServerName = nullptr;
				auto server = szServerName;
				if (server.empty())
				{
					EC_MAPI_S(HrGetOneProp(
						lpPublicMDBNonAdmin, CHANGE_PROP_TYPE(PR_HIERARCHY_SERVER, PT_STRING8), &lpServerName));
					if (mapi::CheckStringProp(lpServerName, PT_STRING8))
					{
						server = lpServerName->Value.lpszA;
					}

					MAPIFreeBuffer(lpServerName);
				}

				if (!server.empty())
				{
					auto szServerDN = BuildServerDN(server,
													"/cn=Microsoft Public MDB"); // STRING_OK
					if (!szServerDN.empty())
					{
						WC_H_S(HrMailboxLogon(
							lpMAPISession,
							lpPublicMDBNonAdmin,
							szServerDN,
							"",
							strings::emptystring,
							ulFlags,
							false,
							&lpPublicMDB));
					}
				}
			}

			if (lpPublicMDBNonAdmin) lpPublicMDBNonAdmin->Release();
			return lpPublicMDB;
		}

		_Check_return_ LPMDB OpenStoreFromMAPIProp(_In_ LPMAPISESSION lpMAPISession, _In_ LPMAPIPROP lpMAPIProp)
		{
			if (!lpMAPISession || !lpMAPIProp) return nullptr;
			LPSPropValue lpProp = nullptr;

			EC_MAPI_S(HrGetOneProp(lpMAPIProp, PR_STORE_ENTRYID, &lpProp));

			LPMDB lpMDB = nullptr;
			if (lpProp && PT_BINARY == PROP_TYPE(lpProp->ulPropTag))
			{
				lpMDB = CallOpenMsgStore(lpMAPISession, NULL, &lpProp->Value.bin, MAPI_BEST_ACCESS);
			}

			MAPIFreeBuffer(lpProp);
			return lpMDB;
		}

		_Check_return_ bool StoreSupportsManageStore(_In_ LPMDB lpMDB)
		{
			if (!lpMDB) return false;

			auto lpIManageStore = mapi::safe_cast<LPEXCHANGEMANAGESTORE>(lpMDB);
			if (lpIManageStore)
			{
				lpIManageStore->Release();
				return true;
			}

			return false;
		}

		_Check_return_ bool StoreSupportsManageStoreEx(_In_ LPMDB lpMDB)
		{
			if (!lpMDB) return false;

			auto lpIManageStore = mapi::safe_cast<LPEXCHANGEMANAGESTOREEX>(lpMDB);
			if (lpIManageStore)
			{
				lpIManageStore->Release();
				return true;
			}

			return false;
		}

		_Check_return_ HRESULT HrUnWrapMDB(_In_ LPMDB lpMDBIn, _Deref_out_ LPMDB* lppMDBOut)
		{
			if (!lpMDBIn || !lppMDBOut) return MAPI_E_INVALID_PARAMETER;
			auto hRes = S_OK;
			auto lpProxyObj = mapi::safe_cast<IProxyStoreObject*>(lpMDBIn);
			if (lpProxyObj)
			{
				LPMDB lpUnwrappedMDB = nullptr;
				hRes = EC_MAPI(lpProxyObj->UnwrapNoRef(reinterpret_cast<LPVOID*>(&lpUnwrappedMDB)));
				if (SUCCEEDED(hRes) && lpUnwrappedMDB)
				{
					// UnwrapNoRef doesn't addref, so we do it here:
					lpUnwrappedMDB->AddRef();
					*lppMDBOut = lpUnwrappedMDB;
				}
			}

			if (lpProxyObj) lpProxyObj->Release();
			return hRes;
		}
	}
}