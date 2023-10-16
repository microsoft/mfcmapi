#include <StdAfx.h>
#include <UI/Dialogs/ContentsTable/MainDlg.h>
#include <UI/Controls/SortList/ContentsTableListCtrl.h>
#include <core/mapi/cache/mapiObjects.h>
#include <UI/Controls/SortList/SingleMAPIPropListCtrl.h>
#include <core/mapi/mapiStoreFunctions.h>
#include <UI/profile.h>
#include <core/mapi/columnTags.h>
#include <UI/Dialogs/MFCUtilityFunctions.h>
#include <UI/Dialogs/HierarchyTable/ABContDlg.h>
#include <UI/Dialogs/Editors/Editor.h>
#include <core/mapi/processor/dumpStore.h>
#include <core/utility/file.h>
#include <UI/Dialogs/ContentsTable/ProfileListDlg.h>
#include <core/utility/import.h>
#include <UI/Dialogs/AboutDlg.h>
#include <UI/Dialogs/ContentsTable/FormContainerDlg.h>
#include <UI/file/FileDialogEx.h>
#include <core/mapi/mapiMime.h>
#include <UI/QuickStart.h>
#include <UI/UIFunctions.h>
#include <core/sortlistdata/contentsData.h>
#include <core/mapi/cache/globalCache.h>
#include <mapistub/library/stubutils.h>
#include <UI/mapiui.h>
#include <UI/addinui.h>
#include <core/utility/registry.h>
#include <core/mapi/mapiOutput.h>
#include <core/utility/strings.h>
#include <core/utility/output.h>
#include <core/mapi/mapiFile.h>
#include <core/interpret/guid.h>
#include <core/interpret/flags.h>
#include <core/interpret/proptags.h>
#include <core/mapi/mapiFunctions.h>
#include <core/property/parseProperty.h>
#include <core/mapi/processor/dumpStore.h>
#include <UI/Dialogs/propList/AccountsDialog.h>
#include <UI/Dialogs/propList/RegistryDialog.h>

namespace dialog
{
	static std::wstring CLASS = L"CMainDlg";

	CMainDlg::CMainDlg(_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects)
		: CContentsTableDlg(
			  lpMapiObjects,
			  ID_PRODUCTNAME,
			  createDialogType::DO_NOT_CALL_CREATE_DIALOG,
			  nullptr,
			  nullptr,
			  &columns::sptSTORECols.tags,
			  columns::STOREColumns,
			  IDR_MENU_MAIN_POPUP,
			  MENU_CONTEXT_MAIN)
	{
		TRACE_CONSTRUCTOR(CLASS);

		CContentsTableDlg::CreateDialogAndMenu(IDR_MENU_MAIN);
		AddLoadMAPIMenus();

		if (registry::displayAboutDialog)
		{
			DisplayAboutDlg(this);
		}
	}

	CMainDlg::~CMainDlg() { TRACE_DESTRUCTOR(CLASS); }

	BEGIN_MESSAGE_MAP(CMainDlg, CContentsTableDlg)
	ON_COMMAND(ID_CLOSEADDRESSBOOK, OnCloseAddressBook)
	ON_COMMAND(ID_COMPUTEGIVENSTOREHASH, OnComputeGivenStoreHash)
	ON_COMMAND(ID_DISPLAYACCOUNTSDIALOG, OnDisplayAccountsDialog)
	ON_COMMAND(ID_DISPLAYREGISTRYDIALOG, OnDisplayRegistryDialog)
	ON_COMMAND(ID_DISPLAYMAILBOXTABLE, OnDisplayMailboxTable)
	ON_COMMAND(ID_DISPLAYPUBLICFOLDERTABLE, OnDisplayPublicFolderTable)
	ON_COMMAND(ID_DUMPSTORECONTENTS, OnDumpStoreContents)
	ON_COMMAND(ID_DUMPSERVERCONTENTSTOTEXT, OnDumpServerContents)
	ON_COMMAND(ID_FASTSHUTDOWN, OnFastShutdown)
	ON_COMMAND(ID_GETCAPABILITIES, OnGetCapabilities)
	ON_COMMAND(ID_ISATTACHMENTBLOCKED, OnIsAttachmentBlocked)
	ON_COMMAND(ID_LOADMAPI, OnLoadMAPI)
	ON_COMMAND(ID_LOGOFF, OnLogoff)
	ON_COMMAND(ID_LOGOFFWITHFLAGS, OnLogoffWithFlags)
	ON_COMMAND(ID_LOGON, OnLogon)
	ON_COMMAND(ID_LOGONANDDISPLAYSTORES, OnLogonAndDisplayStores)
	ON_COMMAND(ID_LOGONWITHFLAGS, OnLogonWithFlags)
	ON_COMMAND(ID_MAPIINITIALIZE, OnMAPIInitialize)
	ON_COMMAND(ID_MAPIOPENLOCALFORMCONTAINER, OnMAPIOpenLocalFormContainer)
	ON_COMMAND(ID_MAPIUNINITIALIZE, OnMAPIUninitialize)
	ON_COMMAND(ID_OPENADDRESSBOOK, OnOpenAddressBook)
	ON_COMMAND(ID_OPENDEFAULTMESSAGESTORE, OnOpenDefaultMessageStore)
	ON_COMMAND(ID_OPENFORMCONTAINER, OnOpenFormContainer)
	ON_COMMAND(ID_OPENMAILBOXWITHDN, OnOpenMailboxWithDN)
	ON_COMMAND(ID_OPENMESSAGESTOREEID, OnOpenMessageStoreEID)
	ON_COMMAND(ID_OPENMESSAGESTORETABLE, OnOpenMessageStoreTable)
	ON_COMMAND(ID_OPENOTHERUSERSMAILBOXFROMGAL, OnOpenOtherUsersMailboxFromGAL)
	ON_COMMAND(ID_OPENPUBLICFOLDERS, OnOpenPublicFolders)
	ON_COMMAND(ID_OPENSELECTEDSTOREDELETEDFOLDERS, OnOpenSelectedStoreDeletedFolders)
	ON_COMMAND(ID_QUERYDEFAULTMESSAGEOPT, OnQueryDefaultMessageOpt)
	ON_COMMAND(ID_QUERYDEFAULTRECIPOPT, OnQueryDefaultRecipOpt)
	ON_COMMAND(ID_QUERYIDENTITY, OnQueryIdentity)
	ON_COMMAND(ID_RESOLVEMESSAGECLASS, OnResolveMessageClass)
	ON_COMMAND(ID_SELECTFORM, OnSelectForm)
	ON_COMMAND(ID_SELECTFORMCONTAINER, OnSelectFormContainer)
	ON_COMMAND(ID_SETDEFAULTSTORE, OnSetDefaultStore)
	ON_COMMAND(ID_UNLOADMAPI, OnUnloadMAPI)
	ON_COMMAND(ID_VIEWABHIERARCHY, OnABHierarchy)
	ON_COMMAND(ID_OPENDEFAULTDIR, OnOpenDefaultDir)
	ON_COMMAND(ID_OPENPAB, OnOpenPAB)
	ON_COMMAND(ID_VIEWSTATUSTABLE, OnStatusTable)
	ON_COMMAND(ID_SHOWPROFILES, OnShowProfiles)
	ON_COMMAND(ID_GETMAPISVCINF, OnGetMAPISVC)
	ON_COMMAND(ID_ADDSERVICESTOMAPISVCINF, OnAddServicesToMAPISVC)
	ON_COMMAND(ID_REMOVESERVICESFROMMAPISVCINF, OnRemoveServicesFromMAPISVC)
	ON_COMMAND(ID_LAUNCHPROFILEWIZARD, OnLaunchProfileWizard)
	ON_COMMAND(ID_VIEWMSGPROPERTIES, OnViewMSGProperties)
	ON_COMMAND(ID_CONVERTMSGTOEML, OnConvertMSGToEML)
	ON_COMMAND(ID_CONVERTEMLTOMSG, OnConvertEMLToMSG)
	ON_COMMAND(ID_CONVERTMSGTOXML, OnConvertMSGToXML)
	ON_COMMAND(ID_DISPLAYMAPIPATH, OnDisplayMAPIPath)
	ON_COMMAND(ID_OPENPUBLICFOLDERWITHDN, OnOpenPublicFolderWithDN)
	END_MESSAGE_MAP()

	void CMainDlg::AddLoadMAPIMenus() const
	{
		output::DebugPrint(output::dbgLevel::LoadMAPI, L"AddLoadMAPIMenus - Extending menus\n");

		// Find the submenu with ID_LOADMAPI on it
		const auto hAddInMenu = ui::LocateSubmenu(::GetMenu(m_hWnd), ID_LOADMAPI);

		// Now add each of the menu entries
		if (hAddInMenu)
		{
			UINT uidCurMenu = ID_LOADMAPIMENUMIN;
			auto paths = mapistub::GetMAPIPaths();
			for (const auto& szPath : paths)
			{
				if (uidCurMenu > ID_LOADMAPIMENUMAX) break;

				output::DebugPrint(output::dbgLevel::LoadMAPI, L"Found MAPI path %ws\n", szPath.c_str());
				const auto lpMenu = ui::CreateMenuEntry(szPath);

				EC_B_S(
					AppendMenu(hAddInMenu, MF_ENABLED | MF_OWNERDRAW, uidCurMenu++, reinterpret_cast<LPCTSTR>(lpMenu)));
			}
		}

		output::DebugPrint(output::dbgLevel::LoadMAPI, L"Done extending menus\n");
	}

	bool CMainDlg::InvokeLoadMAPIMenu(WORD wMenuSelect) const
	{
		if (wMenuSelect < ID_LOADMAPIMENUMIN || wMenuSelect > ID_LOADMAPIMENUMAX) return false;
		output::DebugPrint(output::dbgLevel::LoadMAPI, L"InvokeLoadMAPIMenu - got menu item %u\n", wMenuSelect);

		MENUITEMINFOW subMenu = {0};
		subMenu.cbSize = sizeof(MENUITEMINFO);
		subMenu.fMask = MIIM_DATA;

		WC_B_S(GetMenuItemInfoW(::GetMenu(m_hWnd), wMenuSelect, false, &subMenu));
		if (subMenu.dwItemData)
		{
			const auto lme = reinterpret_cast<ui::LPMENUENTRY>(subMenu.dwItemData);
			output::DebugPrint(output::dbgLevel::LoadMAPI, L"Loading MAPI from %ws\n", lme->m_pName.c_str());
			const auto hMAPI = EC_D(HMODULE, import::MyLoadLibraryW(lme->m_pName));
			mapistub::SetMAPIHandle(hMAPI);
		}

		return false;
	}

	_Check_return_ bool CMainDlg::HandleMenu(WORD wMenuSelect)
	{
		output::DebugPrint(
			output::dbgLevel::Menu, L"CMainDlg::HandleMenu wMenuSelect = 0x%X = %u\n", wMenuSelect, wMenuSelect);
		if (HandleQuickStart(wMenuSelect, this, m_hWnd)) return true;
		if (InvokeLoadMAPIMenu(wMenuSelect)) return true;

		return CContentsTableDlg::HandleMenu(wMenuSelect);
	}

	void CMainDlg::OnInitMenu(_In_ CMenu* pMenu)
	{
		if (pMenu)
		{
			LPMAPISESSION lpMAPISession = nullptr;
			LPADRBOOK lpAddrBook = nullptr;
			const auto bMAPIInitialized = cache::CGlobalCache::getInstance().bMAPIInitialized();
			const auto hMAPI = mapistub::GetMAPIHandle();
			if (m_lpMapiObjects)
			{
				// Don't care if these fail
				lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
				lpAddrBook = m_lpMapiObjects->GetAddrBook(false); // do not release
			}

			const auto bInLoadOp = m_lpContentsTableListCtrl && m_lpContentsTableListCtrl->IsLoading();

			pMenu->EnableMenuItem(ID_LOADMAPI, DIM(!hMAPI && !bInLoadOp));
			pMenu->EnableMenuItem(ID_UNLOADMAPI, DIM(hMAPI && !bInLoadOp));
			pMenu->EnableMenuItem(ID_DISPLAYMAPIPATH, DIM(hMAPI && !bInLoadOp));
			pMenu->CheckMenuItem(ID_LOADMAPI, CHECK(hMAPI && !bInLoadOp));
			pMenu->EnableMenuItem(ID_MAPIINITIALIZE, DIM(!bMAPIInitialized && !bInLoadOp));
			pMenu->EnableMenuItem(ID_MAPIUNINITIALIZE, DIM(bMAPIInitialized && !bInLoadOp));
			pMenu->CheckMenuItem(ID_MAPIINITIALIZE, CHECK(bMAPIInitialized && !bInLoadOp));
			pMenu->EnableMenuItem(ID_LOGON, DIM(!lpMAPISession && !bInLoadOp));
			pMenu->EnableMenuItem(ID_LOGONANDDISPLAYSTORES, DIM(!lpMAPISession && !bInLoadOp));
			pMenu->EnableMenuItem(ID_LOGONWITHFLAGS, DIM(!lpMAPISession && !bInLoadOp));
			pMenu->CheckMenuItem(ID_LOGON, CHECK(lpMAPISession && !bInLoadOp));
			pMenu->EnableMenuItem(ID_LOGOFF, DIM(lpMAPISession && !bInLoadOp));
			pMenu->EnableMenuItem(ID_LOGOFFWITHFLAGS, DIM(lpMAPISession && !bInLoadOp));
			pMenu->EnableMenuItem(ID_FASTSHUTDOWN, DIM(lpMAPISession));
			pMenu->EnableMenuItem(ID_ISATTACHMENTBLOCKED, DIM(lpMAPISession));
			pMenu->EnableMenuItem(ID_VIEWSTATUSTABLE, DIM(lpMAPISession));
			pMenu->EnableMenuItem(ID_QUERYDEFAULTMESSAGEOPT, DIM(lpMAPISession));
			pMenu->EnableMenuItem(ID_QUERYDEFAULTRECIPOPT, DIM(lpMAPISession));
			pMenu->EnableMenuItem(ID_QUERYIDENTITY, DIM(lpMAPISession));
			pMenu->EnableMenuItem(ID_OPENFORMCONTAINER, DIM(lpMAPISession));
			pMenu->EnableMenuItem(ID_RESOLVEMESSAGECLASS, DIM(lpMAPISession));
			pMenu->EnableMenuItem(ID_SELECTFORM, DIM(lpMAPISession));
			pMenu->EnableMenuItem(ID_SELECTFORMCONTAINER, DIM(lpMAPISession));

			pMenu->EnableMenuItem(ID_DISPLAYMAILBOXTABLE, DIM(lpMAPISession));
			pMenu->EnableMenuItem(ID_DISPLAYPUBLICFOLDERTABLE, DIM(lpMAPISession));
			pMenu->EnableMenuItem(ID_OPENMESSAGESTORETABLE, DIM(lpMAPISession && !bInLoadOp));
			pMenu->EnableMenuItem(ID_OPENDEFAULTMESSAGESTORE, DIM(lpMAPISession));

			pMenu->EnableMenuItem(ID_OPENPUBLICFOLDERS, DIM(lpMAPISession));

			pMenu->EnableMenuItem(ID_DUMPSERVERCONTENTSTOTEXT, DIM(lpMAPISession));

			pMenu->EnableMenuItem(ID_OPENOTHERUSERSMAILBOXFROMGAL, DIM(lpMAPISession));
			pMenu->EnableMenuItem(ID_OPENMAILBOXWITHDN, DIM(lpMAPISession));
			pMenu->EnableMenuItem(ID_OPENMESSAGESTOREEID, DIM(lpMAPISession));

			if (m_lpContentsTableListCtrl)
			{
				const int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
				pMenu->EnableMenuItem(ID_OPENSELECTEDSTOREDELETEDFOLDERS, DIM(lpMAPISession && iNumSel));
				pMenu->EnableMenuItem(ID_SETDEFAULTSTORE, DIM(lpMAPISession && 1 == iNumSel));
				pMenu->EnableMenuItem(ID_DUMPSTORECONTENTS, DIM(lpMAPISession && 1 == iNumSel));
				pMenu->EnableMenuItem(ID_COMPUTEGIVENSTOREHASH, DIM(lpMAPISession && 1 == iNumSel));
				pMenu->EnableMenuItem(ID_GETCAPABILITIES, DIM(1 == iNumSel));
			}

			pMenu->EnableMenuItem(ID_OPENADDRESSBOOK, DIM(lpMAPISession && !lpAddrBook));
			pMenu->CheckMenuItem(ID_OPENADDRESSBOOK, CHECK(lpAddrBook));
			pMenu->EnableMenuItem(ID_CLOSEADDRESSBOOK, DIM(lpAddrBook));

			pMenu->EnableMenuItem(ID_VIEWABHIERARCHY, DIM(lpMAPISession));
			pMenu->EnableMenuItem(ID_OPENDEFAULTDIR, DIM(lpMAPISession));
			pMenu->EnableMenuItem(ID_OPENPAB, DIM(lpMAPISession));

			UINT uidCurMenu = ID_LOADMAPIMENUMIN;
			for (; uidCurMenu <= ID_LOADMAPIMENUMAX; uidCurMenu++)
			{
				pMenu->EnableMenuItem(uidCurMenu, DIM(!hMAPI && !bInLoadOp));
			}
		}

		CContentsTableDlg::OnInitMenu(pMenu);
	}

	void CMainDlg::OnCloseAddressBook()
	{
		if (!m_lpMapiObjects) m_lpMapiObjects->SetAddrBook(nullptr);
	}

	void CMainDlg::OnOpenAddressBook()
	{
		LPADRBOOK lpAddrBook = nullptr;

		if (!m_lpMapiObjects) return;

		auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		const auto hRes = EC_MAPI(lpMAPISession->OpenAddressBook(NULL, nullptr, NULL, &lpAddrBook));

		m_lpMapiObjects->SetAddrBook(lpAddrBook);

		if (SUCCEEDED(hRes) && m_lpPropDisplay)
		{
			m_lpPropDisplay->SetDataSource(lpAddrBook, nullptr, true);
		}

		if (lpAddrBook) lpAddrBook->Release();
	}

	void CMainDlg::OnABHierarchy()
	{
		if (!m_lpMapiObjects) return;

		// ensure we have an AB
		static_cast<void>(m_lpMapiObjects->GetAddrBook(true)); // do not release

		// call the dialog
		new CAbContDlg(m_lpMapiObjects);
	}

	void CMainDlg::OnOpenDefaultDir()
	{
		if (!m_lpMapiObjects) return;

		// check if we have an AB - if we don't, get one
		auto lpAddrBook = m_lpMapiObjects->GetAddrBook(true); // do not release
		if (!lpAddrBook) return;

		ULONG cbEID = NULL;
		LPENTRYID lpEID = nullptr;
		ULONG ulObjType = NULL;

		EC_MAPI_S(lpAddrBook->GetDefaultDir(&cbEID, &lpEID));

		if (lpEID)
		{
			auto lpDefaultDir = mapi::CallOpenEntry<LPABCONT>(
				nullptr, lpAddrBook, nullptr, nullptr, cbEID, lpEID, nullptr, MAPI_MODIFY, &ulObjType);
			if (lpDefaultDir)
			{
				EC_H_S(DisplayObject(lpDefaultDir, ulObjType, objectType::otDefault, this));

				lpDefaultDir->Release();
			}

			MAPIFreeBuffer(lpEID);
		}
	}

	void CMainDlg::OnOpenPAB()
	{
		if (!m_lpMapiObjects) return;

		// check if we have an AB - if we don't, get one
		auto lpAddrBook = m_lpMapiObjects->GetAddrBook(true); // do not release
		if (!lpAddrBook) return;

		ULONG cbEID = NULL;
		LPENTRYID lpEID = nullptr;
		ULONG ulObjType = NULL;

		EC_MAPI_S(lpAddrBook->GetPAB(&cbEID, &lpEID));

		if (lpEID)
		{
			auto lpPAB = mapi::CallOpenEntry<LPABCONT>(
				nullptr, lpAddrBook, nullptr, nullptr, cbEID, lpEID, nullptr, MAPI_MODIFY, &ulObjType);
			if (lpPAB)
			{
				EC_H_S(DisplayObject(lpPAB, ulObjType, objectType::otDefault, this));

				lpPAB->Release();
			}

			MAPIFreeBuffer(lpEID);
		}
	}

	void CMainDlg::OnLogonAndDisplayStores()
	{
		if (m_lpContentsTableListCtrl && m_lpContentsTableListCtrl->IsLoading()) return;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		OnLogoff(); // make sure we're logged off first
		OnLogon();
		OnOpenMessageStoreTable();
	}

	_Check_return_ LPMAPIPROP CMainDlg::OpenItemProp(int iSelectedItem, modifyType bModify)
	{
		if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return nullptr;
		output::DebugPrintEx(
			output::dbgLevel::OpenItemProp, CLASS, L"OpenItemProp", L"iSelectedItem = 0x%X\n", iSelectedItem);

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return nullptr;

		LPMDB lpMDB = nullptr;
		const auto lpListData = m_lpContentsTableListCtrl->GetSortListData(iSelectedItem);

		if (lpListData)
		{
			const auto contents = lpListData->cast<sortlistdata::contentsData>();
			if (contents)
			{
				const auto lpEntryID = contents->getEntryID();
				if (lpEntryID)
				{
					ULONG ulFlags = NULL;
					if (modifyType::REQUEST_MODIFY == bModify) ulFlags |= MDB_WRITE;

					lpMDB = mapi::store::CallOpenMsgStore(
						lpMAPISession, reinterpret_cast<ULONG_PTR>(m_hWnd), lpEntryID, ulFlags);
				}
			}
		}

		return lpMDB;
	}

	void CMainDlg::OnOpenDefaultMessageStore()
	{
		if (!m_lpMapiObjects) return;

		auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		auto lpMDB = mapi::store::OpenDefaultMessageStore(lpMAPISession);
		if (!lpMDB) return;

		// Now that we have a message store, try to open the Admin version of it
		ULONG ulFlags = NULL;
		if (mapi::store::StoreSupportsManageStore(lpMDB))
		{
			editor::CEditor MyPrompt(
				this, IDS_OPENDEFMSGSTORE, IDS_OPENWITHFLAGSPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyPrompt.SetPromptPostFix(flags::AllFlagsToString(PROP_ID(PR_PROFILE_OPEN_FLAGS), true));
			MyPrompt.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_CREATESTORENTRYIDFLAGS, false));
			MyPrompt.SetHex(0, NULL);
			if (MyPrompt.DisplayDialog())
			{
				ulFlags = MyPrompt.GetHex(0);
			}
		}

		if (ulFlags)
		{
			ULONG cbEntryID = 0;
			LPENTRYID lpEntryID = nullptr;
			LPSPropValue lpMailboxName = nullptr;

			EC_MAPI_S(lpMAPISession->QueryIdentity(&cbEntryID, &lpEntryID));
			if (lpEntryID)
			{
				auto lpIdentity = mapi::CallOpenEntry<LPMAPIPROP>(
					nullptr, nullptr, nullptr, lpMAPISession, cbEntryID, lpEntryID, nullptr, NULL, nullptr);
				if (lpIdentity)
				{
					EC_MAPI_S(HrGetOneProp(lpIdentity, PR_EMAIL_ADDRESS_W, &lpMailboxName));

					if (strings::CheckStringProp(lpMailboxName, PT_UNICODE))
					{
						auto lpAdminMDB = mapi::store::OpenOtherUsersMailbox(
							lpMAPISession,
							lpMDB,
							L"",
							lpMailboxName->Value.lpszW,
							strings::emptystring,
							ulFlags,
							false);
						if (lpAdminMDB)
						{
							EC_H_S(DisplayObject(lpAdminMDB, NULL, objectType::store, this));
							lpAdminMDB->Release();
						}
					}

					lpIdentity->Release();
				}

				MAPIFreeBuffer(lpEntryID);
			}
		}

		if (lpMDB) lpMDB->Release();
	}

	void CMainDlg::OnOpenMessageStoreEID()
	{
		if (!m_lpMapiObjects) return;

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		editor::CEditor MyEID(
			this, IDS_OPENSTOREEID, IDS_OPENSTOREEIDPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		MyEID.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_EID, false));
		MyEID.AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_FLAGS, false));
		MyEID.SetHex(1, MDB_WRITE);
		MyEID.AddPane(viewpane::CheckPane::Create(2, IDS_EIDBASE64ENCODED, false, false));
		MyEID.AddPane(viewpane::CheckPane::Create(3, IDS_DISPLAYPROPS, false, false));
		MyEID.AddPane(viewpane::CheckPane::Create(4, IDS_UNWRAPSTORE, false, false));

		if (!MyEID.DisplayDialog()) return;

		// Get the entry ID as a binary
		const auto bin = MyEID.GetBinary(0, MyEID.GetCheck(2));

		if (!bin.empty())
		{
			auto sBin = SBinary{static_cast<ULONG>(bin.size()), const_cast<BYTE*>(bin.data())};
			auto lpMDB = mapi::store::CallOpenMsgStore(
				lpMAPISession, reinterpret_cast<ULONG_PTR>(m_hWnd), &sBin, MyEID.GetHex(1));

			if (lpMDB)
			{
				if (MyEID.GetCheck(4))
				{
					const auto lpUnwrappedMDB = mapi::store::HrUnWrapMDB(lpMDB);

					// Ditch the old MDB and substitute the unwrapped one.
					lpMDB->Release();
					lpMDB = lpUnwrappedMDB;
				}
			}

			if (lpMDB)
			{
				if (MyEID.GetCheck(3))
				{
					if (m_lpPropDisplay) m_lpPropDisplay->SetDataSource(lpMDB, nullptr, true);
				}
				else
				{
					EC_H_S(DisplayObject(lpMDB, NULL, objectType::store, this));
				}
			}
		}
	}

	void CMainDlg::OnOpenPublicFolders()
	{
		if (!m_lpMapiObjects) return;

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		editor::CEditor MyPrompt(
			this, IDS_OPENPUBSTORE, IDS_OPENWITHFLAGSPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyPrompt.SetPromptPostFix(flags::AllFlagsToString(PROP_ID(PR_PROFILE_OPEN_FLAGS), true));
		MyPrompt.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_CREATESTORENTRYIDFLAGS, false));
		MyPrompt.SetHex(0, NULL);
		if (!MyPrompt.DisplayDialog()) return;

		auto lpMDB = mapi::store::OpenPublicMessageStore(lpMAPISession, L"", MyPrompt.GetHex(0));
		if (lpMDB)
		{
			EC_H_S(DisplayObject(lpMDB, NULL, objectType::store, this));

			lpMDB->Release();
		}
	}

	void CMainDlg::OnOpenPublicFolderWithDN()
	{
		if (!m_lpMapiObjects) return;

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		editor::CEditor MyPrompt(
			this, IDS_OPENPUBSTORE, IDS_OPENWITHFLAGSPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyPrompt.SetPromptPostFix(flags::AllFlagsToString(PROP_ID(PR_PROFILE_OPEN_FLAGS), true));
		MyPrompt.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_SERVERNAME, false));
		MyPrompt.AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_CREATESTORENTRYIDFLAGS, false));
		MyPrompt.SetHex(1, OPENSTORE_PUBLIC);
		if (!MyPrompt.DisplayDialog()) return;

		auto lpMDB = mapi::store::OpenPublicMessageStore(lpMAPISession, MyPrompt.GetStringW(0), MyPrompt.GetHex(1));
		if (lpMDB)
		{
			EC_H_S(DisplayObject(lpMDB, NULL, objectType::store, this));

			lpMDB->Release();
		}
	}

	void CMainDlg::OnOpenMailboxWithDN()
	{
		if (!m_lpMapiObjects) return;

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		const auto szServerName = mapi::store::GetServerName(lpMAPISession);

		auto lpMDB = mapi::store::OpenDefaultMessageStore(lpMAPISession);
		if (!lpMDB) return;

		if (mapi::store::StoreSupportsManageStore(lpMDB))
		{
			auto lpOtherMDB = ui::mapiui::OpenMailboxWithPrompt(
				lpMAPISession,
				lpMDB,
				szServerName,
				strings::emptystring,
				OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_TAKE_OWNERSHIP);
			if (lpOtherMDB)
			{
				EC_H_S(DisplayObject(lpOtherMDB, NULL, objectType::store, this));

				lpOtherMDB->Release();
			}
		}

		lpMDB->Release();
	}

	void CMainDlg::OnOpenMessageStoreTable()
	{
		if (!m_lpContentsTableListCtrl) return;
		if (m_lpContentsTableListCtrl && m_lpContentsTableListCtrl->IsLoading()) return;
		LPMAPITABLE pStoresTbl = nullptr;

		if (!m_lpMapiObjects) return;

		auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		EC_MAPI_S(lpMAPISession->GetMsgStoresTable(0, &pStoresTbl));

		if (pStoresTbl)
		{
			m_lpContentsTableListCtrl->SetContentsTable(pStoresTbl, tableDisplayFlags::dfNormal, MAPI_STORE);

			pStoresTbl->Release();
		}
	}

	void CMainDlg::OnOpenOtherUsersMailboxFromGAL()
	{
		if (!m_lpMapiObjects) return;

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		const auto lpAddrBook = m_lpMapiObjects->GetAddrBook(true); // do not release
		if (lpAddrBook)
		{
			auto lpMailboxMDB = ui::mapiui::OpenOtherUsersMailboxFromGal(lpMAPISession, lpAddrBook);
			if (lpMailboxMDB)
			{
				EC_H_S(DisplayObject(lpMailboxMDB, NULL, objectType::store, this));

				lpMailboxMDB->Release();
			}
		}
	}

	void CMainDlg::OnOpenSelectedStoreDeletedFolders()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
		for (const auto& lpListData : items)
		{
			if (lpListData)
			{
				const auto contents = lpListData->cast<sortlistdata::contentsData>();
				if (contents)
				{
					const auto lpItemEID = contents->getEntryID();
					if (lpItemEID)
					{
						auto lpMDB = mapi::store::CallOpenMsgStore(
							lpMAPISession, reinterpret_cast<ULONG_PTR>(m_hWnd), lpItemEID, MDB_WRITE | MDB_ONLINE);

						if (lpMDB)
						{
							EC_H_S(DisplayObject(lpMDB, NULL, objectType::storeDeletedItems, this));

							lpMDB->Release();
						}
					}
				}
			}
		}
	}

	void CMainDlg::OnDumpStoreContents()
	{
		if (!m_lpContentsTableListCtrl || !m_lpMapiObjects) return;

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
		for (const auto& lpListData : items)
		{
			if (lpListData)
			{
				const auto contents = lpListData->cast<sortlistdata::contentsData>();
				if (contents)
				{
					const auto lpItemEID = contents->getEntryID();
					if (lpItemEID)
					{
						auto lpMDB = mapi::store::CallOpenMsgStore(
							lpMAPISession, reinterpret_cast<ULONG_PTR>(m_hWnd), lpItemEID, MDB_WRITE);
						if (lpMDB)
						{
							auto szDir = file::GetDirectoryPath(m_hWnd);
							if (!szDir.empty())
							{
								CWaitCursor Wait; // Change the mouse to an hourglass while we work.

								mapi::processor::dumpStore MyDumpStore;
								MyDumpStore.InitFolderPathRoot(szDir);
								MyDumpStore.InitMDB(lpMDB);
								MyDumpStore.ProcessStore();
							}

							lpMDB->Release();
						}
					}
				}
			}
		}
	}

	void CMainDlg::OnDumpServerContents()
	{
		if (!m_lpMapiObjects) return;

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		const auto szServerName = mapi::store::GetServerName(lpMAPISession);

		editor::CEditor MyData(
			this, IDS_DUMPSERVERPRIVATESTORE, IDS_SERVERNAMEPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_SERVERNAME, szServerName, false));

		if (!MyData.DisplayDialog()) return;

		auto szDir = file::GetDirectoryPath(m_hWnd);
		if (!szDir.empty())
		{
			CWaitCursor Wait; // Change the mouse to an hourglass while we work.

			mapi::processor::dumpStore MyDumpStore;
			MyDumpStore.InitMailboxTablePathRoot(szDir);
			MyDumpStore.InitSession(lpMAPISession);
			MyDumpStore.ProcessMailboxTable(MyData.GetStringW(0));
		}
	}

	void CMainDlg::OnLogoff()
	{
		if (m_lpContentsTableListCtrl && m_lpContentsTableListCtrl->IsLoading()) return;
		if (!m_lpMapiObjects) return;

		OnCloseAddressBook();

		// We do this first to free up any stray session pointers
		m_lpContentsTableListCtrl->SetContentsTable(nullptr, tableDisplayFlags::dfNormal, MAPI_STORE);

		m_lpMapiObjects->Logoff(m_hWnd, MAPI_LOGOFF_UI);
	}

	void CMainDlg::OnLogoffWithFlags()
	{
		if (m_lpContentsTableListCtrl && m_lpContentsTableListCtrl->IsLoading()) return;
		if (!m_lpMapiObjects) return;

		editor::CEditor MyData(this, IDS_LOGOFF, IDS_LOGOFFPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_FLAGSINHEX, false));
		MyData.SetHex(0, MAPI_LOGOFF_UI);

		if (!MyData.DisplayDialog()) return;

		OnCloseAddressBook();

		// We do this first to free up any stray session pointers
		m_lpContentsTableListCtrl->SetContentsTable(nullptr, tableDisplayFlags::dfNormal, MAPI_STORE);

		m_lpMapiObjects->Logoff(m_hWnd, MyData.GetHex(0));
	}

	void CMainDlg::OnLogon()
	{
		if (m_lpContentsTableListCtrl && m_lpContentsTableListCtrl->IsLoading()) return;
		if (!m_lpMapiObjects) return;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		constexpr ULONG ulFlags = MAPI_EXTENDED | MAPI_ALLOW_OTHERS | MAPI_NEW_SESSION | MAPI_LOGON_UI |
								  MAPI_EXPLICIT_PROFILE; // display a profile picker box
		m_lpMapiObjects->MAPILogonEx(m_hWnd, nullptr, ulFlags);
	}

	void CMainDlg::OnLogonWithFlags()
	{
		if (!m_lpMapiObjects) return;

		editor::CEditor MyData(
			this, IDS_PROFFORMAPILOGON, IDS_PROFFORMAPILOGONPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_PROFILE, false));
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(1, IDS_FLAGSINHEX, false));
		MyData.SetHex(1, MAPI_EXTENDED | MAPI_EXPLICIT_PROFILE | MAPI_ALLOW_OTHERS | MAPI_NEW_SESSION | MAPI_LOGON_UI);

		if (MyData.DisplayDialog())
		{
			auto szProfile = MyData.GetStringW(0);

			m_lpMapiObjects->MAPILogonEx(
				m_hWnd, szProfile.empty() ? nullptr : LPTSTR(szProfile.c_str()), MyData.GetHex(1) | MAPI_UNICODE);
		}
		else
		{
			output::DebugPrint(output::dbgLevel::Generic, L"MAPILogonEx call cancelled.\n");
		}
	}

	void CMainDlg::OnResolveMessageClass()
	{
		if (!m_lpMapiObjects || !m_lpPropDisplay) return;

		auto lpMAPIFormInfo = ResolveMessageClass(m_lpMapiObjects, nullptr);
		if (lpMAPIFormInfo)
		{
			m_lpPropDisplay->SetDataSource(lpMAPIFormInfo, nullptr, false);
			lpMAPIFormInfo->Release();
		}
	}

	void CMainDlg::OnSelectForm()
	{
		if (!m_lpMapiObjects || !m_lpPropDisplay) return;

		auto lpMAPIFormInfo = SelectForm(m_hWnd, m_lpMapiObjects, nullptr);
		if (lpMAPIFormInfo)
		{
			// TODO: Put some code in here which works with the returned Form Info pointer
			m_lpPropDisplay->SetDataSource(lpMAPIFormInfo, nullptr, false);
			lpMAPIFormInfo->Release();
		}
	}

	void CMainDlg::OnSelectFormContainer()
	{
		LPMAPIFORMMGR lpMAPIFormMgr = nullptr;
		LPMAPIFORMCONTAINER lpMAPIFormContainer = nullptr;

		if (!m_lpMapiObjects || !m_lpPropDisplay) return;

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		EC_MAPI_S(MAPIOpenFormMgr(lpMAPISession, &lpMAPIFormMgr));
		if (lpMAPIFormMgr)
		{
			editor::CEditor MyFlags(
				this, IDS_SELECTFORM, IDS_SELECTFORMPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyFlags.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_FLAGS, false));
			MyFlags.SetHex(0, MAPIFORM_SELECT_ALL_REGISTRIES);

			if (MyFlags.DisplayDialog())
			{
				const auto ulFlags = MyFlags.GetHex(0);
				EC_H_CANCEL_S(lpMAPIFormMgr->SelectFormContainer(
					reinterpret_cast<ULONG_PTR>(m_hWnd), ulFlags, &lpMAPIFormContainer));
				if (lpMAPIFormContainer)
				{
					new CFormContainerDlg(m_lpMapiObjects, lpMAPIFormContainer);

					lpMAPIFormContainer->Release();
				}
			}

			lpMAPIFormMgr->Release();
		}
	}

	void CMainDlg::OnOpenFormContainer()
	{
		LPMAPIFORMMGR lpMAPIFormMgr = nullptr;
		LPMAPIFORMCONTAINER lpMAPIFormContainer = nullptr;

		if (!m_lpMapiObjects || !m_lpPropDisplay) return;

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		EC_MAPI_S(MAPIOpenFormMgr(lpMAPISession, &lpMAPIFormMgr));
		if (lpMAPIFormMgr)
		{
			editor::CEditor MyFlags(
				this, IDS_OPENFORMCONTAINER, IDS_OPENFORMCONTAINERPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyFlags.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_HFRMREG, false));
			MyFlags.SetHex(0, MAPIFORM_SELECT_ALL_REGISTRIES);

			if (MyFlags.DisplayDialog())
			{
				const auto hFrmReg = MyFlags.GetHex(0);
				EC_H_CANCEL_S(lpMAPIFormMgr->OpenFormContainer(hFrmReg, nullptr, &lpMAPIFormContainer));
				if (lpMAPIFormContainer)
				{
					new CFormContainerDlg(m_lpMapiObjects, lpMAPIFormContainer);

					lpMAPIFormContainer->Release();
				}
			}

			lpMAPIFormMgr->Release();
		}
	}

	void CMainDlg::OnMAPIOpenLocalFormContainer()
	{
		cache::CGlobalCache::getInstance().MAPIInitialize(NULL);
		if (!cache::CGlobalCache::getInstance().bMAPIInitialized()) return;

		LPMAPIFORMCONTAINER lpMAPILocalFormContainer = nullptr;

		EC_MAPI_S(MAPIOpenLocalFormContainer(&lpMAPILocalFormContainer));

		if (lpMAPILocalFormContainer)
		{
			new CFormContainerDlg(m_lpMapiObjects, lpMAPILocalFormContainer);

			lpMAPILocalFormContainer->Release();
		}
	}

	void CMainDlg::OnLoadMAPI()
	{
		editor::CEditor MyData(this, IDS_LOADMAPI, IDS_LOADMAPIPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		const auto szDLLPath = file::GetSystemDirectory();
		if (!szDLLPath.empty())
		{
			const auto szFullPath = szDLLPath + L"\\mapi32.dll";
			MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_PATH, szFullPath, false));
		}

		if (MyData.DisplayDialog())
		{
			mapistub::UnloadPrivateMAPI();
			const auto hMAPI = EC_D(HMODULE, import::MyLoadLibraryW(MyData.GetStringW(0)));
			mapistub::SetMAPIHandle(hMAPI);
		}
	}

	void CMainDlg::OnUnloadMAPI()
	{
		editor::CEditor MyData(this, IDS_UNLOADMAPI, IDS_MAPICRASHWARNING, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		if (MyData.DisplayDialog())
		{
			mapistub::UnloadPrivateMAPI();
		}
	}

	void CMainDlg::OnDisplayMAPIPath()
	{
		output::DebugPrint(output::dbgLevel::Generic, L"OnDisplayMAPIPath()\n");
		const auto hMAPI = mapistub::GetMAPIHandle();

		editor::CEditor MyData(this, IDS_MAPIPATHTITLE, NULL, CEDITOR_BUTTON_OK);
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_FILEPATH, true));
		if (hMAPI)
		{
			const auto szMAPIPath = file::GetModuleFileName(hMAPI);
			if (!szMAPIPath.empty())
			{
				MyData.SetStringW(0, szMAPIPath);
			}
		}

		MyData.AddPane(viewpane::CheckPane::Create(1, IDS_REGKEY_FORCEOUTLOOKMAPI, registry::forceOutlookMAPI, true));
		MyData.AddPane(viewpane::CheckPane::Create(2, IDS_REGKEY_FORCESYSTEMMAPI, registry::forceSystemMAPI, true));

		static_cast<void>(MyData.DisplayDialog());
	}

	void CMainDlg::OnMAPIInitialize()
	{
		if (!m_lpMapiObjects) return;

		editor::CEditor MyData(this, IDS_MAPIINIT, IDS_MAPIINITPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_FLAGSINHEX, false));
		MyData.SetHex(0, NULL);

		if (MyData.DisplayDialog())
		{
			cache::CGlobalCache::getInstance().MAPIInitialize(MyData.GetHex(0));
		}
	}

	void CMainDlg::OnMAPIUninitialize()
	{
		editor::CEditor MyData(this, IDS_MAPIUNINIT, IDS_MAPICRASHWARNING, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

		if (MyData.DisplayDialog())
		{
			cache::CGlobalCache::getInstance().MAPIUninitialize();
		}
	}

	void CMainDlg::OnFastShutdown()
	{
		if (!m_lpMapiObjects) return;
		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		auto lpClientShutdown = mapi::safe_cast<LPMAPICLIENTSHUTDOWN>(lpMAPISession);
		if (lpClientShutdown)
		{
			auto hRes = EC_H_MSG(IDS_EDQUERYFASTSHUTDOWNFAILED, lpClientShutdown->QueryFastShutdown());
			WC_H_MSG_S(IDS_EDNOTIFYPROCESSSHUTDOWNFAILED, lpClientShutdown->NotifyProcessShutdown());

			if (SUCCEEDED(hRes))
			{
				hRes = EC_H_MSG(IDS_EDDOFASTSHUTDOWNFAILED, lpClientShutdown->DoFastShutdown());

				if (SUCCEEDED(hRes))
				{
					error::ErrDialog(__FILE__, __LINE__, IDS_EDDOFASTSHUTDOWNSUCCEEDED);
				}
			}
		}

		if (lpClientShutdown) lpClientShutdown->Release();
	}

	void CMainDlg::OnGetCapabilities()
	{
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		auto items = m_lpContentsTableListCtrl->GetSelectedItemData();
		for (const auto& lpListData : items)
		{
			if (lpListData)
			{
				const auto contents = lpListData->cast<sortlistdata::contentsData>();
				if (contents)
				{
					const auto lpItemEID = contents->getEntryID();
					if (lpItemEID)
					{
						auto lpMDB = mapi::store::CallOpenMsgStore(
							lpMAPISession, reinterpret_cast<ULONG_PTR>(m_hWnd), lpItemEID, NULL);

						if (lpMDB)
						{
							auto caps = mapi::safe_cast<LPMSCAPABILITIES>(lpMDB);
							if (caps)
							{
								editor::CEditor MyData(
									this, IDS_CAPABILITIES_DIALOG, IDS_CAPABILITIES_PROMPT, CEDITOR_BUTTON_OK);

								const auto res1 = caps->GetCapabilities(MSCAP_SELECTOR::MSCAP_SEL_RESERVED1);
								MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(
									0,
									strings::formatmessage(IDS_RESERVED, 0),
									strings::format(L"0x%08X", res1),
									true));

								const auto res2 = caps->GetCapabilities(MSCAP_SELECTOR::MSCAP_SEL_RESERVED2);
								MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(
									1,
									strings::formatmessage(IDS_RESERVED, 1),
									strings::format(L"0x%08X", res2),
									true));

								const auto folder = caps->GetCapabilities(MSCAP_SELECTOR::MSCAP_SEL_FOLDER);
								MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(
									2,
									IDS_CAPABILITIES_FOLDER,
									flags::InterpretFlags(flagCapabilitiesFolder, folder),
									true));

								const auto res3 = caps->GetCapabilities(MSCAP_SELECTOR::MSCAP_SEL_RESERVED3);
								MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(
									3,
									strings::formatmessage(IDS_RESERVED, 3),
									strings::format(L"0x%08X", res3),
									true));

								const auto rest = caps->GetCapabilities(MSCAP_SELECTOR::MSCAP_SEL_RESTRICTION);
								MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(
									4,
									IDS_CAPABILITIES_RESTRICTION,
									flags::InterpretFlags(flagCapabilitiesRestriction, rest),
									true));

								for (auto iCap = 5; iCap <= 11; iCap++)
								{
									const auto res = caps->GetCapabilities(static_cast<MSCAP_SELECTOR>(iCap));
									MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(
										iCap,
										strings::formatmessage(IDS_RESERVED, iCap + 1),
										strings::format(L"0x%08X", res),
										true));
								}

								caps->Release();
								static_cast<void>(MyData.DisplayDialog());
							}

							lpMDB->Release();
						}
					}
				}
			}
		}
	}

	void CMainDlg::OnQueryDefaultMessageOpt()
	{
		if (!m_lpMapiObjects) return;
		auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		editor::CEditor MyData(
			this, IDS_QUERYDEFMSGOPT, IDS_ADDRESSTYPEPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.AddPane(
			viewpane::TextPane::CreateSingleLinePane(0, IDS_ADDRESSTYPE, std::wstring(L"EX"), false)); // STRING_OK

		if (!MyData.DisplayDialog()) return;

		ULONG cValues = NULL;
		LPSPropValue lpOptions = nullptr;
		auto adrType = strings::wstringTostring(MyData.GetStringW(0));

		EC_MAPI_S(lpMAPISession->QueryDefaultMessageOpt(
			strings::LPCSTRToLPTSTR(adrType.c_str()),
			NULL, // API doesn't like Unicode
			&cValues,
			&lpOptions));
		if (lpOptions)
		{
			output::outputProperties(output::dbgLevel::Generic, nullptr, cValues, lpOptions, nullptr, false);

			editor::CEditor MyResult(this, IDS_QUERYDEFMSGOPT, IDS_RESULTOFCALLPROMPT, CEDITOR_BUTTON_OK);
			MyResult.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_COUNTOPTIONS, true));
			MyResult.SetHex(0, cValues);

			std::wstring szPropString;
			std::wstring szProp;
			std::wstring szAltProp;
			for (ULONG i = 0; i < cValues; i++)
			{
				property::parseProperty(&lpOptions[i], &szProp, &szAltProp);
				szPropString += strings::formatmessage(
					IDS_OPTIONSSTRUCTURE,
					i,
					proptags::TagToString(lpOptions[i].ulPropTag, nullptr, false, true).c_str(),
					szProp.c_str(),
					szAltProp.c_str());
			}

			MyResult.AddPane(viewpane::TextPane::CreateMultiLinePane(1, IDS_OPTIONS, szPropString, true));

			static_cast<void>(MyResult.DisplayDialog());

			MAPIFreeBuffer(lpOptions);
		}
	}

	void CMainDlg::OnQueryDefaultRecipOpt()
	{
		if (!m_lpMapiObjects) return;
		auto lpAddrBook = m_lpMapiObjects->GetAddrBook(true); // do not release
		if (!lpAddrBook) return;

		editor::CEditor MyData(
			this, IDS_QUERYDEFRECIPOPT, IDS_ADDRESSTYPEPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.AddPane(
			viewpane::TextPane::CreateSingleLinePane(0, IDS_ADDRESSTYPE, std::wstring(L"EX"), false)); // STRING_OK

		if (!MyData.DisplayDialog()) return;

		ULONG cValues = NULL;
		LPSPropValue lpOptions = nullptr;

		auto adrType = strings::wstringTostring(MyData.GetStringW(0));

		EC_MAPI_S(lpAddrBook->QueryDefaultRecipOpt(
			strings::LPCSTRToLPTSTR(adrType.c_str()),
			NULL, // API doesn't like Unicode
			&cValues,
			&lpOptions));

		if (lpOptions)
		{
			output::outputProperties(output::dbgLevel::Generic, nullptr, cValues, lpOptions, nullptr, false);

			editor::CEditor MyResult(this, IDS_QUERYDEFRECIPOPT, IDS_RESULTOFCALLPROMPT, CEDITOR_BUTTON_OK);
			MyResult.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_COUNTOPTIONS, true));
			MyResult.SetHex(0, cValues);

			std::wstring szPropString;
			std::wstring szProp;
			std::wstring szAltProp;
			for (ULONG i = 0; i < cValues; i++)
			{
				property::parseProperty(&lpOptions[i], &szProp, &szAltProp);
				szPropString += strings::formatmessage(
					IDS_OPTIONSSTRUCTURE,
					i,
					proptags::TagToString(lpOptions[i].ulPropTag, nullptr, false, true).c_str(),
					szProp.c_str(),
					szAltProp.c_str());
			}

			MyResult.AddPane(viewpane::TextPane::CreateMultiLinePane(1, IDS_OPTIONS, szPropString, true));

			static_cast<void>(MyResult.DisplayDialog());

			MAPIFreeBuffer(lpOptions);
		}
	}

	void CMainDlg::OnQueryIdentity()
	{
		ULONG cbEntryID = 0;
		LPENTRYID lpEntryID = nullptr;

		if (!m_lpMapiObjects || !m_lpPropDisplay) return;

		auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		EC_MAPI_S(lpMAPISession->QueryIdentity(&cbEntryID, &lpEntryID));

		if (cbEntryID && lpEntryID)
		{
			editor::CEditor MyPrompt(this, IDS_QUERYID, IDS_QUERYIDPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyPrompt.AddPane(viewpane::CheckPane::Create(0, IDS_DISPLAYDETAILSDLG, false, false));
			if (MyPrompt.DisplayDialog())
			{
				if (MyPrompt.GetCheck(0))
				{
					auto lpAB = m_lpMapiObjects->GetAddrBook(true); // do not release
					auto ulUIParam = reinterpret_cast<ULONG_PTR>(m_hWnd);

					EC_H_CANCEL_S(lpAB->Details(
						&ulUIParam,
						nullptr,
						nullptr,
						cbEntryID,
						lpEntryID,
						nullptr,
						nullptr,
						nullptr,
						DIALOG_MODAL)); // API doesn't like Unicode
				}
				else
				{
					auto lpIdentity = mapi::CallOpenEntry<LPMAPIPROP>(
						nullptr, nullptr, nullptr, lpMAPISession, cbEntryID, lpEntryID, nullptr, NULL, nullptr);
					if (lpIdentity)
					{
						m_lpPropDisplay->SetDataSource(lpIdentity, nullptr, true);

						lpIdentity->Release();
					}
				}
			}

			MAPIFreeBuffer(lpEntryID);
		}
	}

	void CMainDlg::OnSetDefaultStore()
	{
		if (!m_lpMapiObjects) return;

		auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		const auto lpListData = m_lpContentsTableListCtrl->GetFirstSelectedItemData();
		if (lpListData)
		{
			const auto contents = lpListData->cast<sortlistdata::contentsData>();
			if (contents)
			{
				const auto lpItemEID = contents->getEntryID();
				if (lpItemEID)
				{
					editor::CEditor MyData(
						this, IDS_SETDEFSTORE, IDS_SETDEFSTOREPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
					MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_FLAGS, false));
					MyData.SetHex(0, MAPI_DEFAULT_STORE);

					if (MyData.DisplayDialog())
					{
						EC_MAPI_S(lpMAPISession->SetDefaultStore(
							MyData.GetHex(0), lpItemEID->cb, reinterpret_cast<LPENTRYID>(lpItemEID->lpb)));
					}
				}
			}
		}
	}

	void CMainDlg::OnIsAttachmentBlocked()
	{
		const auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		editor::CEditor MyData(this, IDS_ISATTBLOCKED, IDS_ENTERFILENAME, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_FILENAME, false));

		if (MyData.DisplayDialog())
		{
			auto bBlocked = false;
			const auto hRes = EC_H(mapi::IsAttachmentBlocked(lpMAPISession, MyData.GetStringW(0).c_str(), &bBlocked));
			if (SUCCEEDED(hRes))
			{
				editor::CEditor MyResult(this, IDS_ISATTBLOCKED, IDS_RESULTOFCALLPROMPT, CEDITOR_BUTTON_OK);
				const auto szResult = strings::loadstring(bBlocked ? IDS_TRUE : IDS_FALSE);
				MyResult.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_RESULT, szResult, true));

				static_cast<void>(MyResult.DisplayDialog());
			}
		}
	}

	void CMainDlg::OnShowProfiles()
	{
		LPMAPITABLE lpProfTable = nullptr;

		cache::CGlobalCache::getInstance().MAPIInitialize(NULL);

		LPPROFADMIN lpProfAdmin = nullptr;
		EC_MAPI_S(MAPIAdminProfiles(0, &lpProfAdmin));
		if (!lpProfAdmin) return;

		EC_MAPI_S(lpProfAdmin->GetProfileTable(
			0, // fMapiUnicode is not supported
			&lpProfTable));
		if (lpProfTable)
		{
			new CProfileListDlg(m_lpMapiObjects, lpProfTable);

			lpProfTable->Release();
		}

		lpProfAdmin->Release();
	}

	void CMainDlg::OnLaunchProfileWizard()
	{
		cache::CGlobalCache::getInstance().MAPIInitialize(NULL);
		if (!cache::CGlobalCache::getInstance().bMAPIInitialized()) return;

		editor::CEditor MyData(
			this, IDS_LAUNCHPROFWIZ, IDS_LAUNCHPROFWIZPROMPT, CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyData.AddPane(viewpane::TextPane::CreateSingleLinePane(0, IDS_FLAGS, false));
		MyData.SetHex(0, MAPI_PW_LAUNCHED_BY_CONFIG);
		MyData.AddPane(
			viewpane::TextPane::CreateSingleLinePane(1, IDS_SERVICE, std::wstring(L"MSEMS"), false)); // STRING_OK

		if (MyData.DisplayDialog())
		{
			auto szProfName = ui::profile::LaunchProfileWizard(m_hWnd, MyData.GetHex(0), MyData.GetStringW(1));
		}
	}

	void CMainDlg::OnGetMAPISVC() { ui::profile::DisplayMAPISVCPath(this); }

	void CMainDlg::OnAddServicesToMAPISVC() { ui::profile::AddServicesToMapiSvcInf(); }

	void CMainDlg::OnRemoveServicesFromMAPISVC() { ui::profile::RemoveServicesFromMapiSvcInf(); }

	void CMainDlg::OnStatusTable()
	{
		LPMAPITABLE lpMAPITable = nullptr;

		if (!m_lpMapiObjects) return;

		auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		EC_MAPI_S(lpMAPISession->GetStatusTable(
			NULL, // This table does not support MAPI_UNICODE!
			&lpMAPITable));
		if (lpMAPITable)
		{

			EC_H_S(DisplayTable(lpMAPITable, objectType::status, this));
			lpMAPITable->Release();
		}
	}

	void CMainDlg::OnDisplayItem()
	{
		// This is an example of how to override the default OnDisplayItem
		CContentsTableDlg::OnDisplayItem();
	}

	void CMainDlg::OnDisplayAccountsDialog()
	{
		if (!m_lpParent || !m_lpMapiObjects) return;

		const auto lpMAPISession = m_lpMapiObjects->LogonGetSession(m_lpParent->GetSafeHwnd()); // do not release
		if (lpMAPISession)
		{
			// Since we've opened a session, populate the store table in the UI
			OnOpenMessageStoreTable();
			new AccountsDialog(m_lpMapiObjects, lpMAPISession);
		}
	}

	void CMainDlg::OnDisplayRegistryDialog()
	{
		if (!m_lpMapiObjects) return;

		new RegistryDialog(m_lpMapiObjects);
	}

	void CMainDlg::OnDisplayMailboxTable()
	{
		if (!m_lpMapiObjects) return;

		DisplayMailboxTable(m_lpMapiObjects);
	}

	void CMainDlg::OnDisplayPublicFolderTable()
	{
		if (!m_lpMapiObjects) return;

		DisplayPublicFolderTable(m_lpMapiObjects);
	}

	void CMainDlg::OnViewMSGProperties()
	{
		if (!m_lpPropDisplay) return;
		cache::CGlobalCache::getInstance().MAPIInitialize(NULL);
		if (!cache::CGlobalCache::getInstance().bMAPIInitialized()) return;

		auto file = file::CFileDialogExW::OpenFile(
			L"msg", // STRING_OK
			strings::emptystring,
			OFN_FILEMUSTEXIST,
			strings::loadstring(IDS_MSGFILES),
			this);
		if (!file.empty())
		{
			auto lpNewMessage = file::LoadMSGToMessage(file);
			if (lpNewMessage)
			{
				WC_H_S(DisplayObject(lpNewMessage, MAPI_MESSAGE, objectType::otDefault, this));
				lpNewMessage->Release();
			}
		}
	}

	void CMainDlg::OnConvertMSGToEML()
	{
		cache::CGlobalCache::getInstance().MAPIInitialize(NULL);
		if (!cache::CGlobalCache::getInstance().bMAPIInitialized()) return;

		auto ulConvertFlags = CCSF_SMTP;
		auto et = IET_UNKNOWN;
		auto mst = USE_DEFAULT_SAVETYPE;
		ULONG ulWrapLines = USE_DEFAULT_WRAPPING;
		auto bDoAdrBook = false;

		const auto hRes =
			WC_H(ui::mapiui::GetConversionToEMLOptions(this, &ulConvertFlags, &et, &mst, &ulWrapLines, &bDoAdrBook));
		if (hRes == S_OK)
		{
			LPADRBOOK lpAdrBook = nullptr;
			if (bDoAdrBook) lpAdrBook = m_lpMapiObjects->GetAddrBook(true); // do not release

			auto msgfile = file::CFileDialogExW::OpenFile(
				L"msg", // STRING_OK
				strings::emptystring,
				OFN_FILEMUSTEXIST,
				strings::loadstring(IDS_MSGFILES),
				this);
			if (!msgfile.empty())
			{
				file::CFileDialogExW dlgFilePickerEML;

				auto emlfile = file::CFileDialogExW::SaveAs(
					L"eml", // STRING_OK
					strings::emptystring,
					OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
					strings::loadstring(IDS_EMLFILES),
					this);
				if (!emlfile.empty())
				{
					EC_H_S(mapi::mapimime::ConvertMSGToEML(
						msgfile.c_str(), emlfile.c_str(), ulConvertFlags, et, mst, ulWrapLines, lpAdrBook));
				}
			}
		}
	}

	void CMainDlg::OnConvertEMLToMSG()
	{
		cache::CGlobalCache::getInstance().MAPIInitialize(NULL);
		if (!cache::CGlobalCache::getInstance().bMAPIInitialized()) return;

		auto ulConvertFlags = CCSF_SMTP;
		auto bDoAdrBook = false;
		auto bDoApply = false;
		auto bUnicode = false;
		HCHARSET hCharSet = nullptr;
		auto cSetApplyType = CSET_APPLY_UNTAGGED;
		const auto hRes = WC_H(ui::mapiui::GetConversionFromEMLOptions(
			this, &ulConvertFlags, &bDoAdrBook, &bDoApply, &hCharSet, &cSetApplyType, &bUnicode));
		if (hRes == S_OK)
		{
			LPADRBOOK lpAdrBook = nullptr;
			if (bDoAdrBook) lpAdrBook = m_lpMapiObjects->GetAddrBook(true); // do not release

			auto emlfile = file::CFileDialogExW::OpenFile(
				L"eml", // STRING_OK
				strings::emptystring,
				OFN_FILEMUSTEXIST,
				strings::loadstring(IDS_EMLFILES),
				this);
			if (!emlfile.empty())
			{
				auto msgfile = file::CFileDialogExW::SaveAs(
					L"msg", // STRING_OK
					strings::emptystring,
					OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
					strings::loadstring(IDS_MSGFILES),
					this);
				if (!msgfile.empty())
				{
					EC_H_S(mapi::mapimime::ConvertEMLToMSG(
						emlfile.c_str(),
						msgfile.c_str(),
						ulConvertFlags,
						bDoApply,
						hCharSet,
						cSetApplyType,
						lpAdrBook,
						bUnicode));
				}
			}
		}
	}

	void CMainDlg::OnConvertMSGToXML()
	{
		cache::CGlobalCache::getInstance().MAPIInitialize(NULL);
		if (!cache::CGlobalCache::getInstance().bMAPIInitialized()) return;

		auto szFileSpec = strings::loadstring(IDS_MSGFILES);

		auto msgfile = file::CFileDialogExW::OpenFile(
			L"msg", // STRING_OK
			strings::emptystring,
			OFN_FILEMUSTEXIST,
			strings::loadstring(IDS_MSGFILES),
			this);
		if (!msgfile.empty())
		{
			file::CFileDialogExW dlgFilePickerXML;

			auto xmlfile = file::CFileDialogExW::SaveAs(
				L"xml", // STRING_OK
				strings::emptystring,
				OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
				strings::loadstring(IDS_XMLFILES),
				this);
			if (!xmlfile.empty())
			{
				auto lpMessage = file::LoadMSGToMessage(msgfile);
				if (lpMessage)
				{
					mapi::processor::dumpStore MyDumpStore;
					MyDumpStore.InitMessagePath(xmlfile);
					// Just assume this message might have attachments
					MyDumpStore.ProcessMessage(lpMessage, true, nullptr);

					lpMessage->Release();
				}
			}
		}
	}

	void CMainDlg::OnComputeGivenStoreHash()
	{
		auto hRes = S_OK;
		CWaitCursor Wait; // Change the mouse to an hourglass while we work.

		if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

		auto lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
		if (!lpMAPISession) return;

		const auto lpListData = m_lpContentsTableListCtrl->GetFirstSelectedItemData();
		if (lpListData)
		{
			const auto contents = lpListData->cast<sortlistdata::contentsData>();
			if (contents)
			{
				const auto lpItemEID = contents->getEntryID();
				if (lpItemEID)
				{
					SBinary serviceUID = {};
					SBinary providerUID = {};
					auto lpProp = lpListData->GetOneProp(PR_SERVICE_UID);
					if (lpProp && PT_BINARY == PROP_TYPE(lpProp->ulPropTag)) serviceUID = mapi::getBin(lpProp);
					lpProp = lpListData->GetOneProp(PR_MDB_PROVIDER);
					if (lpProp && PT_BINARY == PROP_TYPE(lpProp->ulPropTag)) providerUID = mapi::getBin(lpProp);

					MAPIUID emsmdbUID = {0};
					LPPROFSECT lpProfSect = nullptr;
					auto fPublicExchangeStore = false;
					auto fPrivateExchangeStore = false;
					if (providerUID.cb && providerUID.lpb)
					{
						fPublicExchangeStore = mapi::FExchangePublicStore(reinterpret_cast<LPMAPIUID>(providerUID.lpb));
						fPrivateExchangeStore =
							mapi::FExchangePrivateStore(reinterpret_cast<LPMAPIUID>(providerUID.lpb));
					}

					auto fCached = false;
					LPSPropValue lpConfigProp = nullptr;
					LPSPropValue lpPathPropA = nullptr;
					LPSPropValue lpPathPropW = nullptr;
					LPSPropValue lpMappingSig = nullptr;
					LPSTR szPath = nullptr; // do not free
					LPWSTR wzPath = nullptr; // do not free

					// Get profile section
					if (serviceUID.cb && serviceUID.lpb)
					{
						hRes = WC_H(mapi::HrEmsmdbUIDFromStore(
							lpMAPISession, reinterpret_cast<LPMAPIUID>(serviceUID.lpb), &emsmdbUID));
						if (SUCCEEDED(hRes))
						{
							if (fIsSet(output::dbgLevel::Generic))
							{
								auto szGUID = guid::GUIDToString(reinterpret_cast<LPCGUID>(&emsmdbUID));
								output::DebugPrint(
									output::dbgLevel::Generic,
									L"CMainDlg::OnComputeGivenStoreHash, emsmdbUID from PR_EMSMDB_SECTION_UID = %ws\n",
									szGUID.c_str());
							}

							hRes = WC_MAPI(lpMAPISession->OpenProfileSection(&emsmdbUID, nullptr, 0, &lpProfSect));
						}
					}

					if (!(serviceUID.cb && serviceUID.lpb) || FAILED(hRes))
					{
						// For Outlook 2003/2007, HrEmsmdbUIDFromStore may not succeed,
						// so use pbGlobalProfileSectionGuid instead
						WC_MAPI_S(lpMAPISession->OpenProfileSection(
							LPMAPIUID(pbGlobalProfileSectionGuid), nullptr, 0, &lpProfSect));
					}

					if (lpProfSect)
					{
						WC_MAPI_S(HrGetOneProp(lpProfSect, PR_PROFILE_CONFIG_FLAGS, &lpConfigProp));
						if (lpConfigProp && PROP_TYPE(lpConfigProp->ulPropTag) != PT_ERROR)
						{
							if (fPrivateExchangeStore)
							{
								fCached = (lpConfigProp->Value.l & CONFIG_OST_CACHE_PRIVATE) != 0;
							}

							if (fPublicExchangeStore)
							{
								fCached = (lpConfigProp->Value.l & CONFIG_OST_CACHE_PUBLIC) == CONFIG_OST_CACHE_PUBLIC;
							}
						}

						output::DebugPrint(
							output::dbgLevel::Generic,
							L"CMainDlg::OnComputeGivenStoreHash, fPrivateExchangeStore = %d\n",
							fPrivateExchangeStore);
						output::DebugPrint(
							output::dbgLevel::Generic,
							L"CMainDlg::OnComputeGivenStoreHash, fPublicExchangeStore = %d\n",
							fPublicExchangeStore);
						output::DebugPrint(
							output::dbgLevel::Generic, L"CMainDlg::OnComputeGivenStoreHash, fCached = %d\n", fCached);

						if (fCached)
						{
							hRes = WC_MAPI(HrGetOneProp(lpProfSect, PR_PROFILE_OFFLINE_STORE_PATH_W, &lpPathPropW));
							if (FAILED(hRes))
							{
								hRes = WC_MAPI(HrGetOneProp(lpProfSect, PR_PROFILE_OFFLINE_STORE_PATH_A, &lpPathPropA));
							}

							if (SUCCEEDED(hRes))
							{
								if (lpPathPropW && lpPathPropW->Value.lpszW)
								{
									wzPath = lpPathPropW->Value.lpszW;
									output::DebugPrint(
										output::dbgLevel::Generic,
										L"CMainDlg::OnComputeGivenStoreHash, PR_PROFILE_OFFLINE_STORE_PATH_W = %ws\n",
										wzPath);
								}
								else if (lpPathPropA && lpPathPropA->Value.lpszA)
								{
									szPath = lpPathPropA->Value.lpszA;
									output::DebugPrint(
										output::dbgLevel::Generic,
										L"CMainDlg::OnComputeGivenStoreHash, PR_PROFILE_OFFLINE_STORE_PATH_A = %hs\n",
										szPath);
								}
							}
							// If this is an Exchange store with an OST path, it's an OST, so we get the mapping signature
							if ((fPrivateExchangeStore || fPublicExchangeStore) && (wzPath || szPath))
							{
								WC_MAPI_S(HrGetOneProp(lpProfSect, PR_MAPPING_SIGNATURE, &lpMappingSig));
							}
						}
					}

					DWORD dwSigHash = NULL;
					if (lpMappingSig && PT_BINARY == PROP_TYPE(lpMappingSig->ulPropTag))
					{
						dwSigHash = mapi::ComputeStoreHash(
							mapi::getBin(lpMappingSig).cb,
							mapi::getBin(lpMappingSig).lpb,
							nullptr,
							nullptr,
							fPublicExchangeStore);
					}

					const auto dwEIDHash =
						mapi::ComputeStoreHash(lpItemEID->cb, lpItemEID->lpb, szPath, wzPath, fPublicExchangeStore);

					std::wstring szHash;
					if (dwSigHash)
					{
						szHash = strings::formatmessage(IDS_STOREHASHDOUBLEVAL, dwEIDHash, dwSigHash);
					}
					else
					{
						szHash = strings::formatmessage(IDS_STOREHASHVAL, dwEIDHash);
					}

					output::DebugPrint(
						output::dbgLevel::Generic,
						L"CMainDlg::OnComputeGivenStoreHash, Entry ID hash = 0x%08X\n",
						dwEIDHash);
					if (dwSigHash)
						output::DebugPrint(
							output::dbgLevel::Generic,
							L"CMainDlg::OnComputeGivenStoreHash, Mapping Signature hash = 0x%08X\n",
							dwSigHash);

					editor::CEditor Result(this, IDS_STOREHASH, NULL, CEDITOR_BUTTON_OK);
					Result.SetPromptPostFix(szHash);
					static_cast<void>(Result.DisplayDialog());

					MAPIFreeBuffer(lpMappingSig);
					MAPIFreeBuffer(lpPathPropA);
					MAPIFreeBuffer(lpPathPropW);
					MAPIFreeBuffer(lpConfigProp);
					if (lpProfSect) lpProfSect->Release();
				}
			}
		}
	}

	void CMainDlg::HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ LPMAPICONTAINER /*lpContainer*/)
	{
		if (lpParams)
		{
			lpParams->lpMDB = mapi::safe_cast<LPMDB>(lpMAPIProp);
		}

		ui::addinui::InvokeAddInMenu(lpParams);

		if (lpParams && lpParams->lpMDB)
		{
			lpParams->lpMDB->Release();
			lpParams->lpMDB = nullptr;
		}
	}
} // namespace dialog
