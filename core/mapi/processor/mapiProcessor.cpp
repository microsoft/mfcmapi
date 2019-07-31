#include <core/stdafx.h>
#include <core/mapi/processor/mapiProcessor.h>
#include <core/mapi/mapiStoreFunctions.h>
#include <core/mapi/columnTags.h>
#include <core/utility/strings.h>
#include <core/mapi/mapiOutput.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFunctions.h>
#include <core/utility/error.h>

namespace mapi
{
	namespace processor
	{
		mapiProcessor::mapiProcessor()
		{
			m_lpSession = nullptr;
			m_lpMDB = nullptr;
			m_lpFolder = nullptr;
			m_lpResFolderContents = nullptr;
			m_lpSort = nullptr;
			m_ulCount = 0;
		}

		mapiProcessor::~mapiProcessor()
		{
			FreeFolderList();
			if (m_lpFolder) m_lpFolder->Release();
			if (m_lpMDB) m_lpMDB->Release();
			if (m_lpSession) m_lpSession->Release();
		}

		// --------------------------------------------------------------------------------- //

		void mapiProcessor::InitSession(_In_ LPMAPISESSION lpSession)
		{
			if (m_lpSession) m_lpSession->Release();
			m_lpSession = lpSession;
			if (m_lpSession) m_lpSession->AddRef();
		}

		void mapiProcessor::InitMDB(_In_ LPMDB lpMDB)
		{
			if (m_lpMDB) m_lpMDB->Release();
			m_lpMDB = lpMDB;
			if (m_lpMDB) m_lpMDB->AddRef();
		}

		void mapiProcessor::InitFolder(_In_ LPMAPIFOLDER lpFolder)
		{
			if (m_lpFolder) m_lpFolder->Release();
			m_lpFolder = lpFolder;
			if (m_lpFolder) m_lpFolder->AddRef();
			m_szFolderOffset = L"\\"; // STRING_OK
		}

		void mapiProcessor::InitFolderContentsRestriction(_In_opt_ LPSRestriction lpRes)
		{
			// If we ever need to hold this past the scope of the caller we'll need to copy the restriction.
			// For now, just grab a pointer.
			m_lpResFolderContents = lpRes;
		}

		void mapiProcessor::InitMaxOutput(_In_ ULONG ulCount) { m_ulCount = ulCount; }

		void mapiProcessor::InitSortOrder(_In_ const _SSortOrderSet* lpSort)
		{
			// If we ever need to hold this past the scope of the caller we'll need to copy the sort order.
			// For now, just grab a pointer.
			m_lpSort = lpSort;
		}

		// --------------------------------------------------------------------------------- //

		// Server name MUST be passed
		void mapiProcessor::ProcessMailboxTable(_In_ const std::wstring& szExchangeServerName)
		{
			if (!m_lpSession || szExchangeServerName.empty()) return;

			ULONG ulOffset = 0;
			ULONG ulRowNum = 0;
			auto ulFlags = LOGOFF_NO_WAIT;

			BeginMailboxTableWork(szExchangeServerName);

			auto lpPrimaryMDB = store::OpenMessageStoreGUID(m_lpSession, pbExchangeProviderPrimaryUserGuid);

			if (lpPrimaryMDB && store::StoreSupportsManageStore(lpPrimaryMDB)) do
				{
					auto lpMailBoxTable =
						store::GetMailboxTable(lpPrimaryMDB, strings::wstringTostring(szExchangeServerName), ulOffset);
					if (lpMailBoxTable)
					{
						auto hRes = WC_MAPI(lpMailBoxTable->SetColumns(&columns::sptMBXCols.tags, NULL));

						// go to the first row
						if (SUCCEEDED(hRes))
						{
							hRes = WC_MAPI(lpMailBoxTable->SeekRow(BOOKMARK_BEGINNING, 0, nullptr));
						}

						// get each row in turn and process it
						if (SUCCEEDED(hRes))
						{
							LPSRowSet lpRows = nullptr;
							for (ulRowNum = 0;; ulRowNum++)
							{
								if (lpRows) FreeProws(lpRows);
								lpRows = nullptr;

								hRes = WC_MAPI(lpMailBoxTable->QueryRows(1, NULL, &lpRows));
								if (FAILED(hRes) || !lpRows || !lpRows->cRows) break;

								DoMailboxTablePerRowWork(lpPrimaryMDB, lpRows->aRow, ulRowNum + ulOffset + 1);

								const auto lpEmailAddress =
									PpropFindProp(lpRows->aRow->lpProps, lpRows->aRow->cValues, PR_EMAIL_ADDRESS);

								if (!strings::CheckStringProp(lpEmailAddress, PT_TSTRING)) continue;

								if (m_lpMDB)
								{
									WC_MAPI_S(m_lpMDB->StoreLogoff(&ulFlags));
									m_lpMDB->Release();
									m_lpMDB = nullptr;
								}

								m_lpMDB = store::OpenOtherUsersMailbox(
									m_lpSession,
									lpPrimaryMDB,
									strings::wstringTostring(szExchangeServerName),
									strings::wstringTostring(strings::LPCTSTRToWstring(lpEmailAddress->Value.LPSZ)),
									strings::emptystring,
									OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_TAKE_OWNERSHIP,
									false);

								if (m_lpMDB) ProcessStore();
							}

							if (lpRows) FreeProws(lpRows);
							lpRows = nullptr;
						}

						lpMailBoxTable->Release();
						ulOffset += ulRowNum;
					}
				} while (ulRowNum);

			if (m_lpMDB)
			{
				WC_MAPI_S(m_lpMDB->StoreLogoff(&ulFlags));
				m_lpMDB->Release();
				m_lpMDB = nullptr;
			}

			if (lpPrimaryMDB) lpPrimaryMDB->Release();
			EndMailboxTableWork();
		}

		void mapiProcessor::ProcessStore()
		{
			if (!m_lpMDB) return;

			BeginStoreWork();

			AddFolderToFolderList(nullptr, L"\\"); // STRING_OK

			ProcessFolders(true, true, true);

			EndStoreWork();
		}

		void mapiProcessor::ProcessFolders(bool bDoRegular, bool bDoAssociated, bool bDoDescent)
		{
			BeginProcessFoldersWork();

			if (ContinueProcessingFolders())
			{
				if (!m_lpFolder)
				{
					OpenFirstFolderInList();
				}

				while (m_lpFolder)
				{
					DoProcessFoldersPerFolderWork();
					ProcessFolder(bDoRegular, bDoAssociated, bDoDescent);
					if (!ContinueProcessingFolders()) break;
					OpenFirstFolderInList();
				}
			}
			EndProcessFoldersWork();
		}

		bool mapiProcessor::ContinueProcessingFolders() { return true; }

		bool mapiProcessor::ShouldProcessContentsTable() { return true; }

		void mapiProcessor::ProcessFolder(bool bDoRegular, bool bDoAssociated, bool bDoDescent)
		{
			if (!m_lpMDB || !m_lpFolder) return;

			BeginFolderWork();

			if (bDoRegular && ShouldProcessContentsTable())
			{
				ProcessContentsTable(NULL);
			}

			if (bDoAssociated && ShouldProcessContentsTable())
			{
				ProcessContentsTable(MAPI_ASSOCIATED);
			}

			// If we're not processing subfolders, then get outta here
			if (bDoDescent)
			{
				LPMAPITABLE lpHierarchyTable = nullptr;
				// We need to walk down the tree
				// and get the list of kids of the folder
				WC_MAPI_S(m_lpFolder->GetHierarchyTable(fMapiUnicode, &lpHierarchyTable));
				if (lpHierarchyTable)
				{
					enum
					{
						NAME,
						EID,
						SUBFOLDERS,
						FLAGS,
						NUMCOLS
					};
					static SizedSPropTagArray2(NUMCOLS, sptHierarchyCols) = {
						{NUMCOLS, {PR_DISPLAY_NAME, PR_ENTRYID, PR_SUBFOLDERS, PR_CONTAINER_FLAGS}},
					};

					LPSRowSet lpRows = nullptr;
					// If I don't do this, the MSPST provider can blow chunks (MAPI_E_EXTENDED_ERROR) for some folders when I get a row
					// For some reason, this fixes it.
					auto hRes = WC_MAPI(lpHierarchyTable->SetColumns(&sptHierarchyCols.tags, TBL_BATCH));

					if (SUCCEEDED(hRes))
					{ // go to the first row
						hRes = WC_MAPI(lpHierarchyTable->SeekRow(BOOKMARK_BEGINNING, 0, nullptr));
					}

					if (SUCCEEDED(hRes))
					{
						for (;;)
						{
							if (lpRows) FreeProws(lpRows);
							lpRows = nullptr;
							hRes = WC_MAPI(lpHierarchyTable->QueryRows(255, NULL, &lpRows));
							if (FAILED(hRes))
							{
								break;
							}

							if (!lpRows || !lpRows->cRows) break;

							for (ULONG ulRow = 0; ulRow < lpRows->cRows; ulRow++)
							{
								DoFolderPerHierarchyTableRowWork(&lpRows->aRow[ulRow]);
								if (lpRows->aRow[ulRow].lpProps)
								{
									std::wstring szSubFolderOffset; // Holds subfolder name

									const auto lpFolderDisplayName = PpropFindProp(
										lpRows->aRow[ulRow].lpProps, lpRows->aRow[ulRow].cValues, PR_DISPLAY_NAME);

									if (strings::CheckStringProp(lpFolderDisplayName, PT_TSTRING))
									{
										// Clean up the folder name before appending it to the offset
										szSubFolderOffset = m_szFolderOffset +
															strings::SanitizeFileName(strings::LPCTSTRToWstring(
																lpFolderDisplayName->Value.LPSZ)) +
															L"\\"; // STRING_OK
									}
									else
									{
										szSubFolderOffset = m_szFolderOffset + L"UnknownFolder\\"; // STRING_OK
									}

									AddFolderToFolderList(
										&lpRows->aRow[ulRow].lpProps[EID].Value.bin, szSubFolderOffset);
								}
							}
						}
					}

					if (lpRows) FreeProws(lpRows);
					lpHierarchyTable->Release();
				}
			}

			EndFolderWork();
		}

		void mapiProcessor::ProcessContentsTable(ULONG ulFlags)
		{
			if (!m_lpFolder) return;

			enum
			{
				contPR_SUBJECT,
				contPR_MESSAGE_CLASS,
				contPR_MESSAGE_DELIVERY_TIME,
				contPR_HASATTACH,
				contPR_ENTRYID,
				contPR_SEARCH_KEY,
				contPR_RECORD_KEY,
				contPidTagMid,
				contNUM_COLS
			};
			static const SizedSPropTagArray(contNUM_COLS, contCols) = {
				contNUM_COLS,
				{PR_MESSAGE_CLASS,
				 PR_SUBJECT,
				 PR_MESSAGE_DELIVERY_TIME,
				 PR_HASATTACH,
				 PR_ENTRYID,
				 PR_SEARCH_KEY,
				 PR_RECORD_KEY,
				 PidTagMid},
			};

			LPMAPITABLE lpContentsTable = nullptr;

			auto hRes = WC_MAPI(m_lpFolder->GetContentsTable(ulFlags | fMapiUnicode, &lpContentsTable));
			if (SUCCEEDED(hRes) && lpContentsTable)
			{
				hRes = WC_MAPI(lpContentsTable->SetColumns(LPSPropTagArray(&contCols), TBL_BATCH));
			}

			if (SUCCEEDED(hRes) && lpContentsTable && m_lpResFolderContents)
			{
				output::outputRestriction(output::DBGGeneric, nullptr, m_lpResFolderContents, nullptr);
				WC_MAPI_S(lpContentsTable->Restrict(m_lpResFolderContents, TBL_BATCH));
			}

			if (SUCCEEDED(hRes) && lpContentsTable && m_lpSort)
			{
				WC_MAPI_S(lpContentsTable->SortTable(const_cast<LPSSortOrderSet>(m_lpSort), TBL_BATCH));
			}

			if (SUCCEEDED(hRes) && lpContentsTable)
			{
				ULONG ulCountRows = 0;
				WC_MAPI_S(lpContentsTable->GetRowCount(0, &ulCountRows));

				BeginContentsTableWork(ulFlags, ulCountRows);

				LPSRowSet lpRows = nullptr;
				ULONG i = 0;
				for (;;)
				{
					// If we've output enough rows, stop
					if (m_ulCount && i >= m_ulCount) break;
					if (lpRows) FreeProws(lpRows);
					lpRows = nullptr;
					hRes = WC_MAPI(lpContentsTable->QueryRows(255, NULL, &lpRows));
					if (FAILED(hRes))
					{
						break;
					}

					if (!lpRows || !lpRows->cRows) break;

					for (ULONG iRow = 0; iRow < lpRows->cRows; iRow++)
					{
						// If we've output enough rows, stop
						if (m_ulCount && i >= m_ulCount) break;
						if (!DoContentsTablePerRowWork(&lpRows->aRow[iRow], i++)) continue;

						const auto lpMsgEID =
							PpropFindProp(lpRows->aRow[iRow].lpProps, lpRows->aRow[iRow].cValues, PR_ENTRYID);

						if (!lpMsgEID) continue;

						auto lpMessage = mapi::CallOpenEntry<LPMESSAGE>(
							nullptr,
							nullptr,
							m_lpFolder,
							nullptr,
							lpMsgEID->Value.bin.cb,
							reinterpret_cast<LPENTRYID>(lpMsgEID->Value.bin.lpb),
							nullptr,
							MAPI_BEST_ACCESS,
							nullptr);
						if (lpMessage)
						{
							auto bHasAttach = true;
							const auto lpMsgHasAttach =
								PpropFindProp(lpRows->aRow[iRow].lpProps, lpRows->aRow[iRow].cValues, PR_HASATTACH);
							if (lpMsgHasAttach)
							{
								bHasAttach = 0 != lpMsgHasAttach->Value.b;
							}

							ProcessMessage(lpMessage, bHasAttach, nullptr);
							lpMessage->Release();
						}

						lpMessage = nullptr;
					}
				}

				if (lpRows) FreeProws(lpRows);
				lpContentsTable->Release();
			}

			EndContentsTableWork();
		}

		void
		mapiProcessor::ProcessMessage(_In_ LPMESSAGE lpMessage, bool bHasAttach, _In_opt_ LPVOID lpParentMessageData)
		{
			if (!lpMessage) return;

			LPVOID lpData = nullptr;
			if (BeginMessageWork(lpMessage, lpParentMessageData, &lpData))
			{
				ProcessRecipients(lpMessage, lpData);
				ProcessAttachments(lpMessage, bHasAttach, lpData);
				EndMessageWork(lpMessage, lpData);
			}
		}

		void mapiProcessor::ProcessRecipients(_In_ LPMESSAGE lpMessage, _In_opt_ LPVOID lpData)
		{
			if (!lpMessage) return;

			LPMAPITABLE lpRecipTable = nullptr;
			LPSRowSet lpRows = nullptr;

			if (!BeginRecipientWork(lpMessage, lpData)) return;

			// get the recipient table
			WC_MAPI_S(lpMessage->GetRecipientTable(NULL, &lpRecipTable));
			if (lpRecipTable)
			{
				ULONG i = 0;
				for (;;)
				{
					if (lpRows) FreeProws(lpRows);
					lpRows = nullptr;
					const auto hRes = WC_MAPI(lpRecipTable->QueryRows(255, NULL, &lpRows));
					if (FAILED(hRes))
					{
						break;
					}

					if (!lpRows || !lpRows->cRows) break;

					for (ULONG iRow = 0; iRow < lpRows->cRows; iRow++)
					{
						DoMessagePerRecipientWork(lpMessage, lpData, &lpRows->aRow[iRow], i++);
					}
				}

				if (lpRows) FreeProws(lpRows);
				if (lpRecipTable) lpRecipTable->Release();
			}

			EndRecipientWork(lpMessage, lpData);
		}

		void mapiProcessor::ProcessAttachments(_In_ LPMESSAGE lpMessage, bool bHasAttach, _In_opt_ LPVOID lpData)
		{
			if (!lpMessage) return;

			enum
			{
				attPR_ATTACH_NUM,
				atPR_ATTACH_METHOD,
				attPR_ATTACH_FILENAME,
				attPR_DISPLAY_NAME,
				attNUM_COLS
			};
			static const SizedSPropTagArray(attNUM_COLS, attCols) = {
				attNUM_COLS,
				{PR_ATTACH_NUM, PR_ATTACH_METHOD, PR_ATTACH_FILENAME, PR_DISPLAY_NAME},
			};
			auto hRes = S_OK;
			LPMAPITABLE lpAttachTable = nullptr;
			LPSRowSet lpRows = nullptr;

			if (!BeginAttachmentWork(lpMessage, lpData)) return;

			// Only get the attachment table if PR_HASATTACH was true or didn't exist.
			if (bHasAttach)
			{
				// get the attachment table
				hRes = WC_MAPI(lpMessage->GetAttachmentTable(NULL, &lpAttachTable));
			}

			if (SUCCEEDED(hRes) && lpAttachTable)
			{
				// Make sure we have the columns we need
				hRes = WC_MAPI(lpAttachTable->SetColumns(LPSPropTagArray(&attCols), TBL_BATCH));
			}

			if (SUCCEEDED(hRes) && lpAttachTable)
			{
				ULONG i = 0;
				for (;;)
				{
					if (lpRows) FreeProws(lpRows);
					lpRows = nullptr;
					hRes = WC_MAPI(lpAttachTable->QueryRows(255, NULL, &lpRows));
					if (FAILED(hRes))
					{
						break;
					}

					if (!lpRows || !lpRows->cRows) break;

					for (ULONG iRow = 0; iRow < lpRows->cRows; iRow++)
					{
						const auto lpAttachNum =
							PpropFindProp(lpRows->aRow[iRow].lpProps, lpRows->aRow[iRow].cValues, PR_ATTACH_NUM);

						if (lpAttachNum)
						{
							LPATTACH lpAttach = nullptr;
							WC_MAPI_S(lpMessage->OpenAttach(
								lpAttachNum->Value.l, nullptr, MAPI_BEST_ACCESS, static_cast<LPATTACH*>(&lpAttach)));

							DoMessagePerAttachmentWork(lpMessage, lpData, &lpRows->aRow[iRow], lpAttach, i++);
							// Check if the attachment is an embedded message - if it is, parse it as such!
							const auto lpAttachMethod =
								PpropFindProp(lpRows->aRow->lpProps, lpRows->aRow->cValues, PR_ATTACH_METHOD);

							if (lpAttach && lpAttachMethod && ATTACH_EMBEDDED_MSG == lpAttachMethod->Value.l)
							{
								LPMESSAGE lpAttachMsg = nullptr;

								WC_MAPI_S(lpAttach->OpenProperty(
									PR_ATTACH_DATA_OBJ,
									const_cast<LPIID>(&IID_IMessage),
									0,
									NULL, // MAPI_MODIFY,
									reinterpret_cast<LPUNKNOWN*>(&lpAttachMsg)));
								if (lpAttachMsg)
								{
									// Just assume this message might have attachments
									ProcessMessage(lpAttachMsg, true, lpData);
									lpAttachMsg->Release();
								}
							}

							if (lpAttach)
							{
								lpAttach->Release();
								lpAttach = nullptr;
							}
						}
					}
				}

				if (lpRows) FreeProws(lpRows);
				lpAttachTable->Release();
			}

			EndAttachmentWork(lpMessage, lpData);
		}

		// --------------------------------------------------------------------------------- //
		// List Functions
		// --------------------------------------------------------------------------------- //
		void mapiProcessor::AddFolderToFolderList(
			_In_opt_ const _SBinary* lpFolderEID,
			_In_ const std::wstring& szFolderOffsetPath)
		{
			FolderNode newNode{};
			newNode.szFolderOffsetPath = szFolderOffsetPath;
			newNode.lpFolderEID = lpFolderEID ? CopySBinary(lpFolderEID) : nullptr;

			m_List.push_back(newNode);
		}

		// Call OpenEntry on the first folder in the list, remove it from the list
		// If we fail to open a folder, move on to the next item in the list
		void mapiProcessor::OpenFirstFolderInList()
		{
			LPMAPIFOLDER lpFolder = nullptr;
			m_szFolderOffset.clear();

			if (m_lpFolder) m_lpFolder->Release();
			m_lpFolder = nullptr;

			// loop over nodes until we open one or run out
			while (!lpFolder && !m_List.empty())
			{
				auto node = m_List.front();
				lpFolder = mapi::CallOpenEntry<LPMAPIFOLDER>(
					m_lpMDB, nullptr, nullptr, nullptr, node.lpFolderEID, nullptr, MAPI_BEST_ACCESS, nullptr);
				if (!node.szFolderOffsetPath.empty())
				{
					m_szFolderOffset = node.szFolderOffsetPath;
				}

				MAPIFreeBuffer(node.lpFolderEID);
				m_List.pop_front();
			}

			m_lpFolder = lpFolder;
		}

		// Clean up the list
		void mapiProcessor::FreeFolderList() const
		{
			for (auto& node : m_List)
			{
				MAPIFreeBuffer(node.lpFolderEID);
			}
		}

		// --------------------------------------------------------------------------------- //
		// Worker Functions
		// --------------------------------------------------------------------------------- //

		void mapiProcessor::BeginMailboxTableWork(_In_ const std::wstring& /*szExchangeServerName*/) {}

		void
		mapiProcessor::DoMailboxTablePerRowWork(_In_ LPMDB /*lpMDB*/, _In_ const _SRow* /*lpSRow*/, ULONG /*ulCurRow*/)
		{
		}

		void mapiProcessor::EndMailboxTableWork() {}

		void mapiProcessor::BeginStoreWork() {}

		void mapiProcessor::EndStoreWork() {}

		void mapiProcessor::BeginProcessFoldersWork() {}

		void mapiProcessor::DoProcessFoldersPerFolderWork() {}

		void mapiProcessor::EndProcessFoldersWork() {}

		void mapiProcessor::BeginFolderWork() {}

		void mapiProcessor::DoFolderPerHierarchyTableRowWork(_In_ const _SRow* /*lpSRow*/) {}

		void mapiProcessor::EndFolderWork() {}

		void mapiProcessor::BeginContentsTableWork(ULONG /*ulFlags*/, ULONG /*ulCountRows*/) {}

		bool mapiProcessor::DoContentsTablePerRowWork(_In_ const _SRow* /*lpSRow*/, ULONG /*ulCurRow*/)
		{
			return true; // Keep processing
		}

		void mapiProcessor::EndContentsTableWork() {}

		bool mapiProcessor::BeginMessageWork(
			_In_ LPMESSAGE /*lpMessage*/,
			_In_opt_ LPVOID /*lpParentMessageData*/,
			_Deref_out_opt_ LPVOID* lpData)
		{
			if (lpData) *lpData = nullptr;
			return true; // Keep processing
		}

		bool mapiProcessor::BeginRecipientWork(_In_ LPMESSAGE /*lpMessage*/, _In_opt_ LPVOID /*lpData*/)
		{
			return true; // Keep processing
		}

		void mapiProcessor::DoMessagePerRecipientWork(
			_In_opt_ LPMESSAGE /*lpMessage*/,
			_In_opt_ LPVOID /*lpData*/,
			_In_opt_ const _SRow* /*lpSRow*/,
			ULONG /*ulCurRow*/)
		{
		}

		void mapiProcessor::EndRecipientWork(_In_opt_ LPMESSAGE /*lpMessage*/, _In_opt_ LPVOID /*lpData*/) {}

		bool mapiProcessor::BeginAttachmentWork(_In_ LPMESSAGE /*lpMessage*/, _In_opt_ LPVOID /*lpData*/)
		{
			return true; // Keep processing
		}

		void mapiProcessor::DoMessagePerAttachmentWork(
			_In_opt_ LPMESSAGE /*lpMessage*/,
			_In_opt_ LPVOID /*lpData*/,
			_In_opt_ const _SRow* /*lpSRow*/,
			_In_opt_ LPATTACH /*lpAttach*/,
			ULONG /*ulCurRow*/)
		{
		}

		void mapiProcessor::EndAttachmentWork(_In_ LPMESSAGE /*lpMessage*/, _In_opt_ LPVOID /*lpData*/) {}

		void mapiProcessor::EndMessageWork(_In_ LPMESSAGE /*lpMessage*/, _In_opt_ LPVOID /*lpData*/) {}
	} // namespace processor
} // namespace mapi