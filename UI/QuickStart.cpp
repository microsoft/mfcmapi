#include <StdAfx.h>
#include <UI/QuickStart.h>
#include <core/mapi/mapiStoreFunctions.h>
#include <UI/Dialogs/MFCUtilityFunctions.h>
#include <UI/Dialogs/HierarchyTable/ABContDlg.h>
#include <core/mapi/extraPropTags.h>
#include <core/smartview/SmartView.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <core/mapi/mapiABFunctions.h>
#include <UI/ViewPane/CountedTextPane.h>
#include <UI/Dialogs/ContentsTable/MainDlg.h>
#include <UI/Dialogs/Editors/QSSpecialFolders.h>
#include <core/mapi/cache/mapiObjects.h>
#include <core/utility/strings.h>
#include <core/mapi/mapiFunctions.h>

namespace dialog
{
	LPMAPISESSION OpenSessionForQuickStart(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd)
	{
		auto lpMapiObjects = lpHostDlg->GetMapiObjects(); // do not release
		if (!lpMapiObjects) return nullptr;

		const auto lpMAPISession = lpMapiObjects->LogonGetSession(hwnd); // do not release
		if (lpMAPISession)
		{
			// Since we've opened a session, populate the store table in the UI
			lpHostDlg->OnOpenMessageStoreTable();
			return lpMAPISession;
		}

		return nullptr;
	}

	LPMDB OpenStoreForQuickStart(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd)
	{

		auto lpMapiObjects = lpHostDlg->GetMapiObjects(); // do not release
		if (!lpMapiObjects) return nullptr;

		LPMDB lpMDB = nullptr;
		const auto lpMAPISession = OpenSessionForQuickStart(lpHostDlg, hwnd); // do not release
		if (lpMAPISession)
		{
			lpMDB = mapi::store::OpenDefaultMessageStore(lpMAPISession);
			lpMapiObjects->SetMDB(lpMDB);
		}

		return lpMDB;
	}

	_Check_return_ LPADRBOOK OpenABForQuickStart(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd)
	{
		auto lpMapiObjects = lpHostDlg->GetMapiObjects(); // do not release
		if (!lpMapiObjects) return nullptr;

		// ensure we have an AB
		(void) OpenSessionForQuickStart(lpHostDlg, hwnd); // do not release
		auto lpAdrBook = lpMapiObjects->GetAddrBook(true); // do not release

		if (lpAdrBook)
		{
			lpAdrBook->AddRef();
		}

		return lpAdrBook;
	}

	void OnQSDisplayFolder(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd, _In_ ULONG ulFolder)
	{
		auto lpMDB = OpenStoreForQuickStart(lpHostDlg, hwnd);
		if (lpMDB)
		{
			auto lpFolder = mapi::OpenDefaultFolder(ulFolder, lpMDB);
			if (lpFolder)
			{
				WC_H_S(DisplayObject(lpFolder, NULL, objectType::contents, lpHostDlg));

				lpFolder->Release();
			}

			lpMDB->Release();
		}
	}

	void OnQSDisplayTable(
		_In_ CMainDlg* lpHostDlg,
		_In_ HWND hwnd,
		_In_ ULONG ulFolder,
		_In_ ULONG ulProp,
		_In_ objectType tType)
	{
		auto lpMDB = OpenStoreForQuickStart(lpHostDlg, hwnd);
		if (lpMDB)
		{
			auto lpFolder = mapi::OpenDefaultFolder(ulFolder, lpMDB);
			if (lpFolder)
			{
				WC_H_S(DisplayExchangeTable(lpFolder, ulProp, tType, lpHostDlg));
				lpFolder->Release();
			}

			lpMDB->Release();
		}
	}

	void OnQSDisplayDefaultDir(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd)
	{
		auto lpAdrBook = OpenABForQuickStart(lpHostDlg, hwnd);
		if (lpAdrBook)
		{
			ULONG cbEID = NULL;
			LPENTRYID lpEID = nullptr;
			ULONG ulObjType = NULL;
			LPABCONT lpDefaultDir = nullptr;

			const auto hRes = WC_MAPI(lpAdrBook->GetDefaultDir(&cbEID, &lpEID));
			if (SUCCEEDED(hRes))
			{
				lpDefaultDir = mapi::CallOpenEntry<LPABCONT>(
					nullptr, lpAdrBook, nullptr, nullptr, cbEID, lpEID, nullptr, MAPI_MODIFY, &ulObjType);
			}

			if (lpDefaultDir)
			{
				WC_H_S(DisplayObject(lpDefaultDir, ulObjType, objectType::default, lpHostDlg));

				lpDefaultDir->Release();
			}

			MAPIFreeBuffer(lpEID);
			lpAdrBook->Release();
		}
	}

	void OnQSDisplayAB(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd)
	{
		const auto lpMapiObjects = lpHostDlg->GetMapiObjects(); // do not release
		if (!lpMapiObjects) return;

		const auto lpParentWnd = lpHostDlg->GetParentWnd(); // do not release
		if (!lpParentWnd) return;

		auto lpAdrBook = OpenABForQuickStart(lpHostDlg, hwnd);
		if (lpAdrBook)
		{
			// call the dialog
			new CAbContDlg(lpParentWnd, lpMapiObjects);
			lpAdrBook->Release();
		}
	}

	void OnQSDisplayNicknameCache(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd)
	{
		std::wstring szNicknames;
		LPSPropValue lpsProp = nullptr;

		lpHostDlg->UpdateStatusBarText(statusPane::infoText, IDS_STATUSTEXTLOADINGNICKNAME);
		lpHostDlg->SendMessage(WM_PAINT, NULL, NULL); // force paint so we update the status now

		auto lpMDB = OpenStoreForQuickStart(lpHostDlg, hwnd);
		if (lpMDB)
		{
			auto lpFolder = OpenDefaultFolder(mapi::DEFAULT_INBOX, lpMDB);
			if (lpFolder)
			{
				LPMAPITABLE lpTable = nullptr;
				WC_MAPI_S(lpFolder->GetContentsTable(MAPI_ASSOCIATED, &lpTable));
				if (lpTable)
				{
					SRestriction sRes = {};
					SPropValue sPV = {};
					sRes.rt = RES_PROPERTY;
					sRes.res.resProperty.ulPropTag = PR_MESSAGE_CLASS;
					sRes.res.resProperty.relop = RELOP_EQ;
					sRes.res.resProperty.lpProp = &sPV;
					sPV.ulPropTag = sRes.res.resProperty.ulPropTag;
					sPV.Value.LPSZ = const_cast<LPTSTR>(_T("IPM.Configuration.Autocomplete")); // STRING_OK
					auto hRes = WC_MAPI(lpTable->Restrict(&sRes, TBL_BATCH));
					if (SUCCEEDED(hRes))
					{
						enum
						{
							eidPR_ENTRYID,
							eidNUM_COLS
						};
						static const SizedSPropTagArray(eidNUM_COLS, eidCols) = {
							eidNUM_COLS,
							{PR_ENTRYID},
						};

						hRes = WC_MAPI(lpTable->SetColumns(LPSPropTagArray(&eidCols), TBL_BATCH));
						if (SUCCEEDED(hRes))
						{
							LPSRowSet lpRows = nullptr;
							hRes = WC_MAPI(lpTable->QueryRows(1, NULL, &lpRows));

							if (lpRows && 1 == lpRows->cRows &&
								PR_ENTRYID == lpRows->aRow[0].lpProps[eidPR_ENTRYID].ulPropTag)
							{
								auto lpMSG = mapi::CallOpenEntry<LPMESSAGE>(
									lpMDB,
									nullptr,
									nullptr,
									nullptr,
									&lpRows->aRow[0].lpProps[eidPR_ENTRYID].Value.bin,
									nullptr,
									NULL,
									nullptr);
								if (SUCCEEDED(hRes) && lpMSG)
								{
									lpsProp = mapi::GetLargeBinaryProp(lpMSG, PR_ROAMING_BINARYSTREAM);
									if (lpsProp)
									{
										// Get the string interpretation
										szNicknames = smartview::InterpretBinaryAsString(
											lpsProp->Value.bin, parserType::NICKNAMECACHE, lpMSG);
									}

									lpMSG->Release();
								}
							}

							FreeProws(lpRows);
						}
					}

					lpTable->Release();
				}

				lpFolder->Release();
			}

			lpMDB->Release();

			// Display our dialog
			if (!szNicknames.empty() && lpsProp)
			{
				editor::CEditor MyResults(lpHostDlg, IDS_NICKNAME, NULL, CEDITOR_BUTTON_OK);
				MyResults.AddPane(viewpane::TextPane::CreateCollapsibleTextPane(0, NULL, true));
				MyResults.AddPane(viewpane::CountedTextPane::Create(1, IDS_HEX, true, IDS_CB));

				MyResults.SetStringW(0, szNicknames);

				if (lpsProp)
				{
					auto lpPane = std::dynamic_pointer_cast<viewpane::CountedTextPane>(MyResults.GetPane(1));
					if (lpPane) lpPane->SetCount(lpsProp->Value.bin.cb);
					MyResults.SetBinary(1, lpsProp->Value.bin.lpb, lpsProp->Value.bin.cb);
				}

				(void) MyResults.DisplayDialog();
			}

			MAPIFreeBuffer(lpsProp);
		}

		lpHostDlg->UpdateStatusBarText(statusPane::infoText, strings::emptystring);
	}

	enum
	{
		qPR_STORE_SUPPORT_MASK,
		qPR_DISPLAY_NAME_W,
		qPR_MESSAGE_SIZE_EXTENDED,
		qPR_STORAGE_QUOTA_LIMIT,
		qPR_PROHIBIT_SEND_QUOTA,
		qPR_PROHIBIT_RECEIVE_QUOTA,
		qPR_MAX_SUBMIT_MESSAGE_SIZE,
		qPR_QUOTA_WARNING,
		qPR_QUOTA_SEND,
		qPR_QUOTA_RECEIVE,
		qPR_MDB_PROVIDER,
		qNUM_COLS
	};
	static const SizedSPropTagArray(qNUM_COLS, sptaQuota) = {
		qNUM_COLS,
		{PR_STORE_SUPPORT_MASK,
		 PR_DISPLAY_NAME_W,
		 PR_MESSAGE_SIZE_EXTENDED,
		 PR_STORAGE_QUOTA_LIMIT,
		 PR_PROHIBIT_SEND_QUOTA,
		 PR_PROHIBIT_RECEIVE_QUOTA,
		 PR_MAX_SUBMIT_MESSAGE_SIZE,
		 PR_QUOTA_WARNING,
		 PR_QUOTA_SEND,
		 PR_QUOTA_RECEIVE,
		 PR_MDB_PROVIDER},
	};

	std::wstring FormatQuota(LPSPropValue lpProp, ULONG ulPropTag, const std::wstring& szName)
	{
		if (lpProp && lpProp->ulPropTag == ulPropTag)
		{
			return strings::formatmessage(IDS_QUOTAPROP, szName.c_str(), lpProp->Value.l);
		}

		return L"";
	}

#define AddFormattedQuota(__TAG) szQuotaString += FormatQuota(&lpProps[q##__TAG], __TAG, L#__TAG);

	void OnQSDisplayQuota(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd)
	{
		std::wstring szQuotaString;

		lpHostDlg->UpdateStatusBarText(statusPane::infoText, IDS_STATUSTEXTLOADINGQUOTA);
		lpHostDlg->SendMessage(WM_PAINT, NULL, NULL); // force paint so we update the status now

		auto lpMDB = OpenStoreForQuickStart(lpHostDlg, hwnd);
		if (lpMDB)
		{
			ULONG cProps = 0;
			LPSPropValue lpProps = nullptr;

			// Get quota properties
			WC_H_GETPROPS_S(lpMDB->GetProps(LPSPropTagArray(&sptaQuota), fMapiUnicode, &cProps, &lpProps));
			if (lpProps)
			{
				if (lpProps[qPR_DISPLAY_NAME_W].ulPropTag == PR_DISPLAY_NAME_W)
				{
					std::wstring szDisplayName = lpProps[qPR_DISPLAY_NAME_W].Value.lpszW;
					if (szDisplayName.empty())
					{
						szDisplayName = strings::loadstring(IDS_NOTFOUND);
					}

					szQuotaString += strings::formatmessage(IDS_QUOTADISPLAYNAME, szDisplayName.c_str());
				}

				if (lpProps[qPR_MESSAGE_SIZE_EXTENDED].ulPropTag == PR_MESSAGE_SIZE_EXTENDED)
				{
					szQuotaString += strings::formatmessage(
						IDS_QUOTASIZE,
						lpProps[qPR_MESSAGE_SIZE_EXTENDED].Value.li.QuadPart,
						lpProps[qPR_MESSAGE_SIZE_EXTENDED].Value.li.QuadPart / 1024);
				}

				// All of these properties are in kilobytes. Be careful adding a property not in kilobytes.
				AddFormattedQuota(PR_PROHIBIT_SEND_QUOTA);
				AddFormattedQuota(PR_PROHIBIT_RECEIVE_QUOTA);
				AddFormattedQuota(PR_STORAGE_QUOTA_LIMIT);
				AddFormattedQuota(PR_QUOTA_SEND);
				AddFormattedQuota(PR_QUOTA_RECEIVE);
				AddFormattedQuota(PR_QUOTA_WARNING);
				AddFormattedQuota(PR_MAX_SUBMIT_MESSAGE_SIZE);

				if (lpProps[qPR_STORE_SUPPORT_MASK].ulPropTag == PR_STORE_SUPPORT_MASK)
				{
					auto szFlags = smartview::InterpretNumberAsStringProp(
						lpProps[qPR_STORE_SUPPORT_MASK].Value.l, PR_STORE_SUPPORT_MASK);
					szQuotaString +=
						strings::formatmessage(IDS_QUOTAMASK, lpProps[qPR_STORE_SUPPORT_MASK].Value.l, szFlags.c_str());
				}

				if (lpProps[qPR_MDB_PROVIDER].ulPropTag == PR_MDB_PROVIDER)
				{
					szQuotaString += strings::formatmessage(
						IDS_QUOTAPROVIDER, strings::BinToHexString(&lpProps[qPR_MDB_PROVIDER].Value.bin, true).c_str());
				}

				MAPIFreeBuffer(lpProps);
			}

			lpMDB->Release();

			// Display our dialog
			editor::CEditor MyResults(lpHostDlg, IDS_QUOTA, NULL, CEDITOR_BUTTON_OK);
			MyResults.AddPane(viewpane::TextPane::CreateMultiLinePane(0, NULL, true));
			MyResults.SetStringW(0, szQuotaString);

			(void) MyResults.DisplayDialog();
		}

		lpHostDlg->UpdateStatusBarText(statusPane::infoText, strings::emptystring);
	}

	void OnQSOpenUser(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd)
	{
		const auto lpMapiObjects = lpHostDlg->GetMapiObjects(); // do not release
		if (!lpMapiObjects) return;

		const auto lpParentWnd = lpHostDlg->GetParentWnd(); // do not release
		if (!lpParentWnd) return;

		auto lpAdrBook = OpenABForQuickStart(lpHostDlg, hwnd);
		if (lpAdrBook)
		{
			ULONG ulObjType = NULL;
			auto lpMailUser = mapi::ab::SelectUser(lpAdrBook, hwnd, &ulObjType);
			if (lpMailUser)
			{
				EC_H_S(DisplayObject(lpMailUser, ulObjType, objectType::default, lpHostDlg));

				lpMailUser->Release();
			}

			lpAdrBook->Release();
		}
	}

	void OnQSLookupThumbail(_In_ CMainDlg* lpHostDlg, _In_ HWND hwnd)
	{
		LPSPropValue lpThumbnail = nullptr;

		const auto lpMapiObjects = lpHostDlg->GetMapiObjects(); // do not release
		if (!lpMapiObjects) return;

		const auto lpParentWnd = lpHostDlg->GetParentWnd(); // do not release
		if (!lpParentWnd) return;

		auto lpAdrBook = OpenABForQuickStart(lpHostDlg, hwnd);
		if (lpAdrBook)
		{
			auto lpMailUser = mapi::ab::SelectUser(lpAdrBook, hwnd, nullptr);
			if (lpMailUser)
			{
				lpThumbnail = mapi::GetLargeBinaryProp(lpMailUser, PR_EMS_AB_THUMBNAIL_PHOTO);
				lpMailUser->Release();
			}

			lpAdrBook->Release();
		}

		editor::CEditor MyResults(lpHostDlg, IDS_QSTHUMBNAIL, NULL, CEDITOR_BUTTON_OK);

		if (lpThumbnail)
		{
			MyResults.AddPane(viewpane::CountedTextPane::Create(0, IDS_HEX, true, IDS_CB));
			MyResults.AddPane(viewpane::TextPane::CreateCollapsibleTextPane(1, IDS_ANSISTRING, true));

			auto lpPane = std::dynamic_pointer_cast<viewpane::CountedTextPane>(MyResults.GetPane(0));
			if (lpPane) lpPane->SetCount(lpThumbnail->Value.bin.cb);
			MyResults.SetBinary(0, lpThumbnail->Value.bin.lpb, lpThumbnail->Value.bin.cb);

			MyResults.SetStringA(1, std::string(LPCSTR(lpThumbnail->Value.bin.lpb), lpThumbnail->Value.bin.cb));
		}
		else
		{
			MyResults.AddPane(viewpane::TextPane::CreateSingleLinePaneID(0, 0, IDS_QSTHUMBNAILNOTFOUND, true));
		}

		(void) MyResults.DisplayDialog();

		MAPIFreeBuffer(lpThumbnail);
		if (lpAdrBook) lpAdrBook->Release();
	}

	bool HandleQuickStart(_In_ WORD wMenuSelect, _In_ CMainDlg* lpHostDlg, _In_ HWND hwnd)
	{
		switch (wMenuSelect)
		{
		case ID_QSINBOX:
			OnQSDisplayFolder(lpHostDlg, hwnd, mapi::DEFAULT_INBOX);
			return true;
		case ID_QSCALENDAR:
			OnQSDisplayFolder(lpHostDlg, hwnd, mapi::DEFAULT_CALENDAR);
			return true;
		case ID_QSCONTACTS:
			OnQSDisplayFolder(lpHostDlg, hwnd, mapi::DEFAULT_CONTACTS);
			return true;
		case ID_QSJOURNAL:
			OnQSDisplayFolder(lpHostDlg, hwnd, mapi::DEFAULT_JOURNAL);
			return true;
		case ID_QSNOTES:
			OnQSDisplayFolder(lpHostDlg, hwnd, mapi::DEFAULT_NOTES);
			return true;
		case ID_QSTASKS:
			OnQSDisplayFolder(lpHostDlg, hwnd, mapi::DEFAULT_TASKS);
			return true;
		case ID_QSREMINDERS:
			OnQSDisplayFolder(lpHostDlg, hwnd, mapi::DEFAULT_REMINDERS);
			return true;
		case ID_QSDRAFTS:
			OnQSDisplayFolder(lpHostDlg, hwnd, mapi::DEFAULT_DRAFTS);
			return true;
		case ID_QSSENTITEMS:
			OnQSDisplayFolder(lpHostDlg, hwnd, mapi::DEFAULT_SENTITEMS);
			return true;
		case ID_QSOUTBOX:
			OnQSDisplayFolder(lpHostDlg, hwnd, mapi::DEFAULT_OUTBOX);
			return true;
		case ID_QSDELETEDITEMS:
			OnQSDisplayFolder(lpHostDlg, hwnd, mapi::DEFAULT_DELETEDITEMS);
			return true;
		case ID_QSFINDER:
			OnQSDisplayFolder(lpHostDlg, hwnd, mapi::DEFAULT_FINDER);
			return true;
		case ID_QSIPM_SUBTREE:
			OnQSDisplayFolder(lpHostDlg, hwnd, mapi::DEFAULT_IPM_SUBTREE);
			return true;
		case ID_QSLOCALFREEBUSY:
			OnQSDisplayFolder(lpHostDlg, hwnd, mapi::DEFAULT_LOCALFREEBUSY);
			return true;
		case ID_QSCONFLICTS:
			OnQSDisplayFolder(lpHostDlg, hwnd, mapi::DEFAULT_CONFLICTS);
			return true;
		case ID_QSSYNCISSUES:
			OnQSDisplayFolder(lpHostDlg, hwnd, mapi::DEFAULT_SYNCISSUES);
			return true;
		case ID_QSLOCALFAILURES:
			OnQSDisplayFolder(lpHostDlg, hwnd, mapi::DEFAULT_LOCALFAILURES);
			return true;
		case ID_QSSERVERFAILURES:
			OnQSDisplayFolder(lpHostDlg, hwnd, mapi::DEFAULT_SERVERFAILURES);
			return true;
		case ID_QSJUNKMAIL:
			OnQSDisplayFolder(lpHostDlg, hwnd, mapi::DEFAULT_JUNKMAIL);
			return true;
		case ID_QSRULES:
			OnQSDisplayTable(lpHostDlg, hwnd, mapi::DEFAULT_INBOX, PR_RULES_TABLE, objectType::rules);
			return true;
		case ID_QSDEFAULTDIR:
			OnQSDisplayDefaultDir(lpHostDlg, hwnd);
			return true;
		case ID_QSAB:
			OnQSDisplayAB(lpHostDlg, hwnd);
			return true;
		case ID_QSCALPERM:
			OnQSDisplayTable(lpHostDlg, hwnd, mapi::DEFAULT_CALENDAR, PR_ACL_TABLE, objectType::ACL);
			return true;
		case ID_QSNICKNAME:
			OnQSDisplayNicknameCache(lpHostDlg, hwnd);
			return true;
		case ID_QSQUOTA:
			OnQSDisplayQuota(lpHostDlg, hwnd);
			return true;
		case ID_QSCHECKSPECIALFOLDERS:
			editor::OnQSCheckSpecialFolders(lpHostDlg, hwnd);
			return true;
		case ID_QSTHUMBNAIL:
			OnQSLookupThumbail(lpHostDlg, hwnd);
			return true;
		case ID_QSOPENUSER:
			OnQSOpenUser(lpHostDlg, hwnd);
			return true;
		}
		return false;
	}
} // namespace dialog