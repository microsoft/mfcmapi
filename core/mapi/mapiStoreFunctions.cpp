// Collection of useful MAPI Store functions

#include <core/stdafx.h>
#include <core/mapi/mapiStoreFunctions.h>
#include <core/utility/strings.h>
#include <core/mapi/extraPropTags.h>
#include <core/utility/registry.h>
#include <core/utility/output.h>
#include <core/utility/error.h>
#include <core/mapi/mapiFunctions.h>
#include <core/mapi/interfaces.h>

namespace mapi
{
	namespace store
	{
		std::function<std::string()> promptServerName;

		_Check_return_ LPMDB
		CallOpenMsgStore(_In_ LPMAPISESSION lpSession, _In_ ULONG_PTR ulUIParam, _In_ LPSBinary lpEID, ULONG ulFlags)
		{
			if (!lpSession || !lpEID) return nullptr;

			if (registry::forceMDBOnline)
			{
				ulFlags |= MDB_ONLINE;
			}

			output::DebugPrint(output::DBGOpenItemProp, L"CallOpenMsgStore ulFlags = 0x%X\n", ulFlags);

			LPMDB lpMDB = nullptr;
			auto ignore = (ulFlags & MDB_ONLINE) ? MAPI_E_UNKNOWN_FLAGS : S_OK;
			const auto hRes = EC_H_IGNORE(
				ignore,
				lpSession->OpenMsgStore(
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
				output::DebugPrint(output::DBGOpenItemProp, L"CallOpenMsgStore 2nd attempt ulFlags = 0x%X\n", ulFlags);

				EC_H_S(lpSession->OpenMsgStore(
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

		_Check_return_ LPMAPITABLE GetMailboxTable1(_In_ LPMDB lpMDB, const std::string& szServerDN, ULONG ulFlags)
		{
			if (!lpMDB || szServerDN.empty()) return nullptr;

			LPMAPITABLE lpMailboxTable = nullptr;
			auto lpManageStore1 = mapi::safe_cast<LPEXCHANGEMANAGESTORE>(lpMDB);
			if (lpManageStore1)
			{
				WC_MAPI_S(lpManageStore1->GetMailboxTable(LPSTR(szServerDN.c_str()), &lpMailboxTable, ulFlags));

				lpManageStore1->Release();
			}

			return lpMailboxTable;
		}

		_Check_return_ LPMAPITABLE
		GetMailboxTable3(_In_ LPMDB lpMDB, const std::string& szServerDN, ULONG ulOffset, ULONG ulFlags)
		{
			if (!lpMDB || szServerDN.empty()) return nullptr;

			LPMAPITABLE lpMailboxTable = nullptr;
			auto lpManageStore3 = mapi::safe_cast<LPEXCHANGEMANAGESTORE3>(lpMDB);
			if (lpManageStore3)
			{
				WC_MAPI_S(lpManageStore3->GetMailboxTableOffset(
					LPSTR(szServerDN.c_str()), &lpMailboxTable, ulFlags, ulOffset));

				lpManageStore3->Release();
			}

			return lpMailboxTable;
		}

		_Check_return_ LPMAPITABLE GetMailboxTable5(
			_In_ LPMDB lpMDB,
			const std::string& szServerDN,
			ULONG ulOffset,
			ULONG ulFlags,
			_In_opt_ LPGUID lpGuidMDB)
		{
			if (!lpMDB || szServerDN.empty()) return nullptr;

			LPMAPITABLE lpMailboxTable = nullptr;
			auto lpManageStore5 = mapi::safe_cast<LPEXCHANGEMANAGESTORE5>(lpMDB);
			if (lpManageStore5)
			{
				EC_MAPI_S(lpManageStore5->GetMailboxTableEx(
					LPSTR(szServerDN.c_str()), lpGuidMDB, &lpMailboxTable, ulFlags, ulOffset));

				lpManageStore5->Release();
			}

			return lpMailboxTable;
		}

		// lpMDB needs to be an Exchange MDB - OpenMessageStoreGUID(pbExchangeProviderPrimaryUserGuid) can get one if there's one to be had
		// Use GetServerName to get the default server
		// Will try IID_IExchangeManageStore3 first and fail back to IID_IExchangeManageStore
		_Check_return_ LPMAPITABLE GetMailboxTable(_In_ LPMDB lpMDB, const std::string& szServerName, ULONG ulOffset)
		{
			if (!lpMDB || !StoreSupportsManageStore(lpMDB)) return nullptr;

			LPMAPITABLE lpMailboxTable = nullptr;

			auto szServerDN = BuildServerDN(szServerName, "");
			if (!szServerDN.empty())
			{
				lpMailboxTable = GetMailboxTable3(lpMDB, szServerDN, ulOffset, fMapiUnicode);

				if (!lpMailboxTable && 0 == ulOffset)
				{
					lpMailboxTable = GetMailboxTable1(lpMDB, szServerDN, fMapiUnicode);
				}
			}

			return lpMailboxTable;
		}

		_Check_return_ LPMAPITABLE GetPublicFolderTable1(_In_ LPMDB lpMDB, const std::string& szServerDN, ULONG ulFlags)
		{
			if (!lpMDB || szServerDN.empty()) return nullptr;

			LPMAPITABLE lpPFTable = nullptr;
			auto lpManageStore1 = mapi::safe_cast<LPEXCHANGEMANAGESTORE>(lpMDB);
			if (lpManageStore1)
			{
				EC_MAPI_S(lpManageStore1->GetPublicFolderTable(LPSTR(szServerDN.c_str()), &lpPFTable, ulFlags));

				lpManageStore1->Release();
			}

			return lpPFTable;
		}

		_Check_return_ LPMAPITABLE
		GetPublicFolderTable4(_In_ LPMDB lpMDB, const std::string& szServerDN, ULONG ulOffset, ULONG ulFlags)
		{
			if (!lpMDB || szServerDN.empty()) return nullptr;

			LPMAPITABLE lpPFTable = nullptr;
			auto lpManageStore4 = mapi::safe_cast<LPEXCHANGEMANAGESTORE4>(lpMDB);
			if (lpManageStore4)
			{
				EC_MAPI_S(lpManageStore4->GetPublicFolderTableOffset(
					LPSTR(szServerDN.c_str()), &lpPFTable, ulFlags, ulOffset));
				lpManageStore4->Release();
			}

			return lpPFTable;
		}

		_Check_return_ LPMAPITABLE GetPublicFolderTable5(
			_In_ LPMDB lpMDB,
			const std::string& szServerDN,
			ULONG ulOffset,
			ULONG ulFlags,
			_In_opt_ LPGUID lpGuidMDB)
		{
			if (!lpMDB || szServerDN.empty()) return nullptr;

			LPMAPITABLE lpPFTable = nullptr;
			auto lpManageStore5 = mapi::safe_cast<LPEXCHANGEMANAGESTORE5>(lpMDB);
			if (lpManageStore5)
			{
				EC_MAPI_S(lpManageStore5->GetPublicFolderTableEx(
					LPSTR(szServerDN.c_str()), lpGuidMDB, &lpPFTable, ulFlags, ulOffset));

				lpManageStore5->Release();
			}

			return lpPFTable;
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

			if (strings::CheckStringProp(lpServerName, PT_STRING8)) // profiles are ASCII only
			{
				serverName = lpServerName->Value.lpszA;
			}
			else if (promptServerName)
			{
				serverName = promptServerName();
			}

			MAPIFreeBuffer(lpServerName);
			if (pGlobalProfSect) pGlobalProfSect->Release();
			if (pSvcAdmin) pSvcAdmin->Release();
			return serverName;
		}

		_Check_return_ SBinary CreateStoreEntryID(
			_In_ LPMDB lpMDB, // open message store
			const std::string& lpszMsgStoreDN, // desired message store DN
			const std::string& lpszMailboxDN, // desired mailbox DN or NULL
			ULONG ulFlags) // desired flags for CreateStoreEntryID
		{
			if (!lpMDB || lpszMsgStoreDN.empty() || !StoreSupportsManageStore(lpMDB))
			{
				if (!lpMDB) output::DebugPrint(output::DBGGeneric, L"CreateStoreEntryID: MDB was NULL\n");
				if (lpszMsgStoreDN.empty())
					output::DebugPrint(output::DBGGeneric, L"CreateStoreEntryID: lpszMsgStoreDN was missing\n");
				return {};
			}

			auto eid = SBinary{};
			auto lpXManageStore = mapi::safe_cast<LPEXCHANGEMANAGESTORE>(lpMDB);
			if (lpXManageStore)
			{
				output::DebugPrint(
					output::DBGGeneric,
					L"CreateStoreEntryID: Creating EntryID. StoreDN = \"%hs\", MailboxDN = \"%hs\", Flags = \"0x%X\"\n",
					lpszMsgStoreDN.c_str(),
					lpszMailboxDN.c_str(),
					ulFlags);

				EC_MAPI_S(lpXManageStore->CreateStoreEntryID(
					LPSTR(lpszMsgStoreDN.c_str()),
					lpszMailboxDN.empty() ? nullptr : LPSTR(lpszMailboxDN.c_str()),
					ulFlags,
					&eid.cb,
					reinterpret_cast<LPENTRYID*>(&eid.lpb)));

				lpXManageStore->Release();
			}

			return eid;
		}

		_Check_return_ SBinary CreateStoreEntryID2(
			_In_ LPMDB lpMDB, // open message store
			const std::string& lpszMsgStoreDN, // desired message store DN
			const std::string& lpszMailboxDN, // desired mailbox DN or NULL
			const std::wstring& smtpAddress,
			ULONG ulFlags) // desired flags for CreateStoreEntryID
		{
			if (!lpMDB || lpszMsgStoreDN.empty() || !StoreSupportsManageStoreEx(lpMDB))
			{
				if (!lpMDB) output::DebugPrint(output::DBGGeneric, L"CreateStoreEntryID2: MDB was NULL\n");
				if (lpszMsgStoreDN.empty())
					output::DebugPrint(output::DBGGeneric, L"CreateStoreEntryID2: lpszMsgStoreDN was missing\n");
				return {};
			}

			auto eid = SBinary{};
			auto lpXManageStoreEx = mapi::safe_cast<LPEXCHANGEMANAGESTOREEX>(lpMDB);
			if (lpXManageStoreEx)
			{
				output::DebugPrint(
					output::DBGGeneric,
					L"CreateStoreEntryID2: Creating EntryID. StoreDN = \"%hs\", MailboxDN = \"%hs\", SmtpAddress = "
					L"\"%ws\", Flags = \"0x%X\"\n",
					lpszMsgStoreDN.c_str(),
					lpszMailboxDN.c_str(),
					smtpAddress.c_str(),
					ulFlags);
				SPropValue sProps[4] = {};
				sProps[0].ulPropTag = PR_PROFILE_MAILBOX;
				sProps[0].Value.lpszA = LPSTR(lpszMailboxDN.c_str());

				sProps[1].ulPropTag = PR_PROFILE_MDB_DN;
				sProps[1].Value.lpszA = LPSTR(lpszMsgStoreDN.c_str());

				sProps[2].ulPropTag = PR_FORCE_USE_ENTRYID_SERVER;
				sProps[2].Value.b = true;

				sProps[3].ulPropTag = PR_PROFILE_USER_SMTP_EMAIL_ADDRESS_W;
				sProps[3].Value.lpszW = smtpAddress.empty() ? nullptr : const_cast<LPWSTR>(smtpAddress.c_str());

				EC_MAPI_S(lpXManageStoreEx->CreateStoreEntryID2(
					smtpAddress.empty() ? _countof(sProps) - 1 : _countof(sProps),
					reinterpret_cast<LPSPropValue>(&sProps),
					ulFlags,
					&eid.cb,
					reinterpret_cast<LPENTRYID*>(&eid.lpb)));

				lpXManageStoreEx->Release();
			}

			return eid;
		}

		_Check_return_ SBinary CreateStoreEntryID(
			_In_ LPMDB lpMDB, // open message store
			const std::string& lpszMsgStoreDN, // desired message store DN
			const std::string& lpszMailboxDN, // desired mailbox DN or NULL
			const std::wstring& smtpAddress,
			ULONG ulFlags, // desired flags for CreateStoreEntryID
			bool bForceServer) // Use CreateStoreEntryID2
		{
			// Use an empty MailboxDN to open the public store
			if (lpszMailboxDN.empty())
			{
				ulFlags |= OPENSTORE_PUBLIC;
			}

			if (!bForceServer)
			{
				return CreateStoreEntryID(lpMDB, lpszMsgStoreDN, lpszMailboxDN, ulFlags);
			}
			else
			{
				return CreateStoreEntryID2(lpMDB, lpszMsgStoreDN, lpszMailboxDN, smtpAddress, ulFlags);
			}
		}

		// Stolen from MBLogon.c in the EDK to avoid compiling and linking in the entire beast
		// Cleaned up to fit in with other functions
		// $--MailboxLogon------------------------------------------------------
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

		_Check_return_ LPMDB MailboxLogon(
			_In_ LPMAPISESSION lpMAPISession, // MAPI session handle
			_In_ LPMDB lpMDB, // open message store
			const std::string& lpszMsgStoreDN, // desired message store DN
			const std::string& lpszMailboxDN, // desired mailbox DN or NULL
			const std::wstring& smtpAddress,
			ULONG ulFlags, // desired flags for CreateStoreEntryID
			bool bForceServer) // Use CreateStoreEntryID2
		{
			if (!lpMAPISession)
			{
				output::DebugPrint(output::DBGGeneric, L"MailboxLogon: Session was NULL\n");
				return nullptr;
			}

			auto eid = CreateStoreEntryID(lpMDB, lpszMsgStoreDN, lpszMailboxDN, smtpAddress, ulFlags, bForceServer);

			LPMDB lpMailboxMDB = nullptr;
			if (eid.cb && eid.lpb)
			{
				lpMailboxMDB = CallOpenMsgStore(
					lpMAPISession,
					NULL,
					&eid,
					MDB_NO_DIALOG | MDB_NO_MAIL | // spooler not notified of our presence
						MDB_TEMPORARY | // message store not added to MAPI profile
						MAPI_BEST_ACCESS); // normally WRITE, but allow access to RO store
			}

			MAPIFreeBuffer(eid.lpb);
			return lpMailboxMDB;
		}

		_Check_return_ LPMDB OpenDefaultMessageStore(_In_ LPMAPISESSION lpMAPISession)
		{
			if (!lpMAPISession) return nullptr;

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

			EC_MAPI_S(lpMAPISession->GetMsgStoresTable(0, &pStoresTbl));

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

			LPMDB lpDefaultMDB = nullptr;
			if (pRow && pRow->cRows)
			{
				lpDefaultMDB = CallOpenMsgStore(lpMAPISession, NULL, &pRow->aRow[0].lpProps[EID].Value.bin, MDB_WRITE);
			}

			if (pRow) FreeProws(pRow);
			if (pStoresTbl) pStoresTbl->Release();
			return lpDefaultMDB;
		}

		// Build DNs for call to MailboxLogon
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
				output::DBGGeneric,
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
						output::DBGGeneric, L"Calling MailboxLogon with Server DN = \"%hs\"\n", szServerDN.c_str());
					lpOtherUserMDB =
						MailboxLogon(lpMAPISession, lpMDB, szServerDN, szMailboxDN, smtpAddress, ulFlags, bForceServer);
				}
			}

			return lpOtherUserMDB;
		}

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
					if (strings::CheckStringProp(lpServerName, PT_STRING8))
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
						lpPublicMDB = MailboxLogon(
							lpMAPISession, lpPublicMDBNonAdmin, szServerDN, "", strings::emptystring, ulFlags, false);
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

		_Check_return_ LPMDB HrUnWrapMDB(_In_ LPMDB lpMDBIn)
		{
			if (!lpMDBIn) return nullptr;
			LPMDB lpUnwrappedMDB = nullptr;
			auto lpProxyObj = mapi::safe_cast<IProxyStoreObject*>(lpMDBIn);
			if (lpProxyObj)
			{
				EC_MAPI_S(lpProxyObj->UnwrapNoRef(reinterpret_cast<LPVOID*>(&lpUnwrappedMDB)));
				if (lpUnwrappedMDB)
				{
					// UnwrapNoRef doesn't addref, so we do it here:
					lpUnwrappedMDB->AddRef();
				}

				lpProxyObj->Release();
			}

			return lpUnwrappedMDB;
		}
	} // namespace store
} // namespace mapi