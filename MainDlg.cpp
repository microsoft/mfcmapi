#include "stdafx.h"
#include "MainDlg.h"
#include "ContentsTableListCtrl.h"
#include "MapiObjects.h"
#include "SingleMAPIPropListCtrl.h"
#include "MAPIFunctions.h"
#include "MAPIStoreFunctions.h"
#include "MAPIProfileFunctions.h"
#include "ColumnTags.h"
#include "MFCUtilityFunctions.h"
#include "AbContDlg.h"
#include "Editor.h"
#include "Dumpstore.h"
#include "File.h"
#include "ProfileListDlg.h"
#include "ImportProcs.h"
#include "AboutDlg.h"
#include "FormContainerDlg.h"
#include "FileDialogEx.h"
#include "MAPIMime.h"
#include "InterpretProp2.h"
#include "QuickStart.h"
#include "UIFunctions.h"
#include "SortList/ContentsData.h"

static wstring CLASS = L"CMainDlg";

CMainDlg::CMainDlg(
	_In_ CParentWnd* pParentWnd,
	_In_ CMapiObjects* lpMapiObjects
) :
	CContentsTableDlg(
		pParentWnd,
		lpMapiObjects,
		ID_PRODUCTNAME,
		mfcmapiDO_NOT_CALL_CREATE_DIALOG,
		NULL,
		(LPSPropTagArray)&sptSTORECols,
		NUMSTORECOLUMNS,
		STOREColumns,
		IDR_MENU_MAIN_POPUP,
		MENU_CONTEXT_MAIN)
{
	TRACE_CONSTRUCTOR(CLASS);

	CreateDialogAndMenu(IDR_MENU_MAIN);
	AddLoadMAPIMenus();

	if (RegKeys[regkeyDISPLAY_ABOUT_DIALOG].ulCurDWORD)
	{
		DisplayAboutDlg(this);
	}
}

CMainDlg::~CMainDlg()
{
	TRACE_DESTRUCTOR(CLASS);
}

BEGIN_MESSAGE_MAP(CMainDlg, CContentsTableDlg)
	ON_COMMAND(ID_CLOSEADDRESSBOOK, OnCloseAddressBook)
	ON_COMMAND(ID_COMPUTEGIVENSTOREHASH, OnComputeGivenStoreHash)
	ON_COMMAND(ID_DISPLAYMAILBOXTABLE, OnDisplayMailboxTable)
	ON_COMMAND(ID_DISPLAYPUBLICFOLDERTABLE, OnDisplayPublicFolderTable)
	ON_COMMAND(ID_DUMPSTORECONTENTS, OnDumpStoreContents)
	ON_COMMAND(ID_DUMPSERVERCONTENTSTOTEXT, OnDumpServerContents)
	ON_COMMAND(ID_FASTSHUTDOWN, OnFastShutdown)
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
END_MESSAGE_MAP()

void CMainDlg::AddLoadMAPIMenus()//HMENU hMenu, ULONG ulAddInContext)
{
	DebugPrint(DBGLoadMAPI, L"AddLoadMAPIMenus - Extending menus\n");
	HRESULT hRes = S_OK;

	// Find the submenu with ID_LOADMAPI on it
	HMENU hAddInMenu = LocateSubmenu(::GetMenu(m_hWnd), ID_LOADMAPI);

	UINT uidCurMenu = ID_LOADMAPIMENUMIN;

	// Now add each of the menu entries
	if (SUCCEEDED(hRes))
	{
		MAPIPathIterator* mpi = new MAPIPathIterator(true);

		LPWSTR szPath = NULL;

		if (mpi)
		{
			while (uidCurMenu <= ID_LOADMAPIMENUMAX)
			{
				szPath = mpi->GetNextMAPIPath();
				if (!szPath) break;

				DebugPrint(DBGLoadMAPI, L"Found MAPI path %ws\n", szPath);
				LPMENUENTRY lpMenu = CreateMenuEntry(szPath);

				EC_B(AppendMenu(
					hAddInMenu,
					MF_ENABLED | MF_OWNERDRAW,
					uidCurMenu++,
					(LPCTSTR)lpMenu));
				delete[] szPath;
			}
		}

		delete mpi;
	}

	DebugPrint(DBGLoadMAPI, L"Done extending menus\n");
}

bool CMainDlg::InvokeLoadMAPIMenu(WORD wMenuSelect)
{
	if (wMenuSelect < ID_LOADMAPIMENUMIN || wMenuSelect > ID_LOADMAPIMENUMAX) return false;
	DebugPrint(DBGLoadMAPI, L"InvokeLoadMAPIMenu - got menu item %u\n", wMenuSelect);

	HRESULT hRes = S_OK;
	MENUITEMINFOW subMenu = { 0 };
	subMenu.cbSize = sizeof(MENUITEMINFO);
	subMenu.fMask = MIIM_DATA;

	WC_B(GetMenuItemInfoW(
		::GetMenu(m_hWnd),
		wMenuSelect,
		false,
		&subMenu));
	if (subMenu.dwItemData)
	{
		LPMENUENTRY lme = (LPMENUENTRY)subMenu.dwItemData;
		DebugPrint(DBGLoadMAPI, L"Loading MAPI from %ws\n", lme->m_pName);
		HMODULE hMAPI = NULL;
		EC_D(hMAPI, MyLoadLibraryW(lme->m_pName));
		SetMAPIHandle(hMAPI);
	}

	return false;
}

_Check_return_ bool CMainDlg::HandleMenu(WORD wMenuSelect)
{
	DebugPrint(DBGMenu, L"CMainDlg::HandleMenu wMenuSelect = 0x%X = %u\n", wMenuSelect, wMenuSelect);
	if (HandleQuickStart(wMenuSelect, this, m_hWnd)) return true;
	if (InvokeLoadMAPIMenu(wMenuSelect)) return true;

	return CContentsTableDlg::HandleMenu(wMenuSelect);
}

void CMainDlg::OnInitMenu(_In_ CMenu* pMenu)
{
	if (pMenu)
	{
		LPMAPISESSION lpMAPISession = NULL;
		LPADRBOOK lpAddrBook = NULL;
		bool bMAPIInitialized = false;
		HMODULE hMAPI = GetMAPIHandle();
		if (m_lpMapiObjects)
		{
			// Don't care if these fail
			lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
			lpAddrBook = m_lpMapiObjects->GetAddrBook(false); // do not release
			bMAPIInitialized = m_lpMapiObjects->bMAPIInitialized();
		}

		bool bInLoadOp = m_lpContentsTableListCtrl && m_lpContentsTableListCtrl->IsLoading();

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
			int iNumSel = m_lpContentsTableListCtrl->GetSelectedCount();
			pMenu->EnableMenuItem(ID_OPENSELECTEDSTOREDELETEDFOLDERS, DIM(lpMAPISession && iNumSel));
			pMenu->EnableMenuItem(ID_SETDEFAULTSTORE, DIM(lpMAPISession && 1 == iNumSel));
			pMenu->EnableMenuItem(ID_DUMPSTORECONTENTS, DIM(lpMAPISession && 1 == iNumSel));
			pMenu->EnableMenuItem(ID_COMPUTEGIVENSTOREHASH, DIM(lpMAPISession && 1 == iNumSel));
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
	m_lpMapiObjects->SetAddrBook(NULL);
}

void CMainDlg::OnOpenAddressBook()
{
	LPADRBOOK lpAddrBook = NULL;
	HRESULT hRes = S_OK;

	if (!m_lpMapiObjects) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	EC_MAPI(lpMAPISession->OpenAddressBook(
		NULL,
		NULL,
		NULL,
		&lpAddrBook));

	m_lpMapiObjects->SetAddrBook(lpAddrBook);

	if (m_lpPropDisplay)
		EC_H(m_lpPropDisplay->SetDataSource(lpAddrBook, NULL, true));

	if (lpAddrBook) lpAddrBook->Release();
}

void CMainDlg::OnABHierarchy()
{
	if (!m_lpMapiObjects) return;

	// ensure we have an AB
	(void)m_lpMapiObjects->GetAddrBook(true); // do not release

	// call the dialog
	new CAbContDlg(
		m_lpParent,
		m_lpMapiObjects);

}

void CMainDlg::OnOpenDefaultDir()
{
	if (!m_lpMapiObjects) return;

	// check if we have an AB - if we don't, get one
	LPADRBOOK lpAddrBook = m_lpMapiObjects->GetAddrBook(true); // do not release
	if (!lpAddrBook) return;

	HRESULT hRes = S_OK;
	ULONG cbEID = NULL;
	LPENTRYID lpEID = NULL;
	ULONG ulObjType = NULL;
	LPABCONT lpDefaultDir = NULL;

	EC_MAPI(lpAddrBook->GetDefaultDir(
		&cbEID,
		&lpEID));

	EC_H(CallOpenEntry(
		NULL,
		lpAddrBook,
		NULL,
		NULL,
		cbEID,
		lpEID,
		NULL,
		MAPI_MODIFY,
		&ulObjType,
		(LPUNKNOWN*)&lpDefaultDir));

	if (lpDefaultDir)
	{
		EC_H(DisplayObject(lpDefaultDir, ulObjType, otDefault, this));

		lpDefaultDir->Release();
	}

	MAPIFreeBuffer(lpEID);
}

void CMainDlg::OnOpenPAB()
{
	if (!m_lpMapiObjects) return;

	// check if we have an AB - if we don't, get one
	LPADRBOOK lpAddrBook = m_lpMapiObjects->GetAddrBook(true); // do not release
	if (!lpAddrBook) return;

	HRESULT hRes = S_OK;
	ULONG cbEID = NULL;
	LPENTRYID lpEID = NULL;
	ULONG ulObjType = NULL;
	LPABCONT lpPAB = NULL;

	EC_MAPI(lpAddrBook->GetPAB(
		&cbEID,
		&lpEID));

	EC_H(CallOpenEntry(
		NULL,
		lpAddrBook,
		NULL,
		NULL,
		cbEID,
		lpEID,
		NULL,
		MAPI_MODIFY,
		&ulObjType,
		(LPUNKNOWN*)&lpPAB));

	if (lpPAB)
	{
		EC_H(DisplayObject(lpPAB, ulObjType, otDefault, this));

		lpPAB->Release();
	}

	MAPIFreeBuffer(lpEID);
}

void CMainDlg::OnLogonAndDisplayStores()
{
	if (m_lpContentsTableListCtrl && m_lpContentsTableListCtrl->IsLoading()) return;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	OnLogoff(); // make sure we're logged off first
	OnLogon();
	OnOpenMessageStoreTable();
}

_Check_return_ HRESULT CMainDlg::OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp)
{
	HRESULT hRes = S_OK;

	*lppMAPIProp = NULL;
	DebugPrintEx(DBGOpenItemProp, CLASS, L"OpenItemProp", L"iSelectedItem = 0x%X\n", iSelectedItem);

	if (!m_lpContentsTableListCtrl || !lppMAPIProp) return MAPI_E_INVALID_PARAMETER;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return MAPI_E_INVALID_PARAMETER;

	SortListData* lpListData = NULL;
	lpListData = (SortListData*)m_lpContentsTableListCtrl->GetItemData(iSelectedItem);
	if (lpListData && lpListData->Contents())
	{
		LPSBinary lpEntryID = NULL;
		lpEntryID = lpListData->Contents()->lpEntryID;
		if (lpEntryID)
		{
			ULONG ulFlags = NULL;
			if (mfcmapiREQUEST_MODIFY == bModify) ulFlags |= MDB_WRITE;

			EC_H(CallOpenMsgStore(
				lpMAPISession,
				(ULONG_PTR)m_hWnd,
				lpEntryID,
				ulFlags,
				(LPMDB*)lppMAPIProp));
		}
	}

	return hRes;
}

void CMainDlg::OnOpenDefaultMessageStore()
{
	LPMDB lpMDB = NULL;
	HRESULT hRes = S_OK;

	if (!m_lpMapiObjects) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	EC_H(OpenDefaultMessageStore(lpMAPISession, &lpMDB));
	if (!lpMDB) return;

	// Now that we have a message store, try to open the Admin version of it
	ULONG ulFlags = NULL;
	if (StoreSupportsManageStore(lpMDB))
	{
		CEditor MyPrompt(
			this,
			IDS_OPENDEFMSGSTORE,
			IDS_OPENWITHFLAGSPROMPT,
			1,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyPrompt.SetPromptPostFix(AllFlagsToString(PROP_ID(PR_PROFILE_OPEN_FLAGS), true));
		MyPrompt.InitPane(0, CreateSingleLinePane(IDS_CREATESTORENTRYIDFLAGS, false));
		MyPrompt.SetHex(0, NULL);
		WC_H(MyPrompt.DisplayDialog());
		if (S_OK == hRes)
		{
			ulFlags = MyPrompt.GetHex(0);
		}
	}

	if (ulFlags)
	{
		ULONG cbEntryID = 0;
		LPENTRYID lpEntryID = NULL;
		LPMAPIPROP lpIdentity = NULL;
		LPSPropValue lpMailboxName = NULL;

		EC_MAPI(lpMAPISession->QueryIdentity(
			&cbEntryID,
			&lpEntryID));

		if (lpEntryID)
		{
			EC_H(CallOpenEntry(
				NULL,
				NULL,
				NULL,
				lpMAPISession,
				cbEntryID,
				lpEntryID,
				NULL,
				NULL,
				NULL,
				(LPUNKNOWN*)&lpIdentity));
			if (lpIdentity)
			{
				EC_MAPI(HrGetOneProp(
					lpIdentity,
					PR_EMAIL_ADDRESS,
					&lpMailboxName));

				if (CheckStringProp(lpMailboxName, PT_TSTRING))
				{
					LPMDB lpAdminMDB = NULL;
					EC_H(OpenOtherUsersMailbox(
						lpMAPISession,
						lpMDB,
						emptystring,
						LPCTSTRToWstring(lpMailboxName->Value.LPSZ),
						ulFlags,
						false,
						&lpAdminMDB));
					if (lpAdminMDB)
					{
						EC_H(DisplayObject(
							lpAdminMDB,
							NULL,
							otStore,
							this));
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
	HRESULT hRes = S_OK;
	if (!m_lpMapiObjects) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	CEditor MyEID(
		this,
		IDS_OPENSTOREEID,
		IDS_OPENSTOREEIDPROMPT,
		5,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	MyEID.InitPane(0, CreateSingleLinePane(IDS_EID, false));
	MyEID.InitPane(1, CreateSingleLinePane(IDS_FLAGS, false));
	MyEID.SetHex(1, MDB_WRITE);
	MyEID.InitPane(2, CreateCheckPane(IDS_EIDBASE64ENCODED, false, false));
	MyEID.InitPane(3, CreateCheckPane(IDS_DISPLAYPROPS, false, false));
	MyEID.InitPane(4, CreateCheckPane(IDS_UNWRAPSTORE, false, false));

	WC_H(MyEID.DisplayDialog());
	if (S_OK != hRes) return;

	// Get the entry ID as a binary
	SBinary sBin = { 0 };
	EC_H(MyEID.GetEntryID(0, MyEID.GetCheck(2), (size_t*)&sBin.cb, (LPENTRYID*)&sBin.lpb));

	if (SUCCEEDED(hRes))
	{
		LPMDB lpMDB = NULL;
		EC_H(CallOpenMsgStore(
			lpMAPISession,
			(ULONG_PTR)m_hWnd,
			&sBin,
			MyEID.GetHex(1),
			&lpMDB));

		if (lpMDB)
		{
			if (MyEID.GetCheck(4))
			{
				LPMDB lpUnwrappedMDB = NULL;
				EC_H(HrUnWrapMDB(lpMDB, &lpUnwrappedMDB));

				// Ditch the old MDB and substitute the unwrapped one.
				lpMDB->Release();
				lpMDB = lpUnwrappedMDB;
			}

			if (MyEID.GetCheck(3))
			{
				if (m_lpPropDisplay)
					EC_H(m_lpPropDisplay->SetDataSource(lpMDB, NULL, true));
			}
			else
			{
				EC_H(DisplayObject(
					lpMDB,
					NULL,
					otStore,
					this));
			}
		}
	}

	delete[] sBin.lpb;
}

void CMainDlg::OnOpenPublicFolders()
{
	LPMDB lpMDB = NULL;
	HRESULT hRes = S_OK;

	if (!m_lpMapiObjects) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	CEditor MyPrompt(
		this,
		IDS_OPENPUBSTORE,
		IDS_OPENWITHFLAGSPROMPT,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyPrompt.SetPromptPostFix(AllFlagsToString(PROP_ID(PR_PROFILE_OPEN_FLAGS), true));
	MyPrompt.InitPane(0, CreateSingleLinePane(IDS_CREATESTORENTRYIDFLAGS, false));
	MyPrompt.SetHex(0, NULL);
	WC_H(MyPrompt.DisplayDialog());
	if (S_OK == hRes)
	{
		EC_H(OpenPublicMessageStore(
			lpMAPISession,
			MyPrompt.GetHex(0),
			&lpMDB));

		if (lpMDB)
		{
			EC_H(DisplayObject(
				lpMDB,
				NULL,
				otStore,
				this));

			lpMDB->Release();
		}
	}
}

void CMainDlg::OnOpenMailboxWithDN()
{
	HRESULT hRes = S_OK;
	LPMDB lpMDB = NULL;
	LPMDB lpOtherMDB = NULL;

	if (!m_lpMapiObjects) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	wstring szServerName = GetServerName(lpMAPISession);

	EC_H(OpenDefaultMessageStore(lpMAPISession, &lpMDB));
	if (!lpMDB) return;

	if (StoreSupportsManageStore(lpMDB))
	{
		WC_H(OpenMailboxWithPrompt(
			lpMAPISession,
			lpMDB,
			szServerName,
			emptystring,
			OPENSTORE_USE_ADMIN_PRIVILEGE | OPENSTORE_TAKE_OWNERSHIP,
			&lpOtherMDB));
		if (SUCCEEDED(hRes) && lpOtherMDB)
		{
			EC_H(DisplayObject(
				lpOtherMDB,
				NULL,
				otStore,
				this));

			lpOtherMDB->Release();
		}
	}

	lpMDB->Release();
}

void CMainDlg::OnOpenMessageStoreTable()
{
	if (!m_lpContentsTableListCtrl) return;
	if (m_lpContentsTableListCtrl && m_lpContentsTableListCtrl->IsLoading()) return;
	LPMAPITABLE pStoresTbl = NULL;
	HRESULT hRes = S_OK;

	if (!m_lpMapiObjects) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	EC_MAPI(lpMAPISession->GetMsgStoresTable(0, &pStoresTbl));

	if (pStoresTbl)
	{
		EC_H(m_lpContentsTableListCtrl->SetContentsTable(
			pStoresTbl,
			dfNormal,
			MAPI_STORE));

		pStoresTbl->Release();
	}
}

void CMainDlg::OnOpenOtherUsersMailboxFromGAL()
{
	HRESULT hRes = S_OK;
	LPMDB lpMailboxMDB = NULL; // Ptr to any another's msg store interface.

	if (!m_lpMapiObjects) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	LPADRBOOK lpAddrBook = m_lpMapiObjects->GetAddrBook(true); // do not release
	if (lpAddrBook)
	{
		EC_H_CANCEL(OpenOtherUsersMailboxFromGal(
			lpMAPISession,
			lpAddrBook,
			&lpMailboxMDB));

		if (lpMailboxMDB)
		{
			EC_H(DisplayObject(
				lpMailboxMDB,
				NULL,
				otStore,
				this));

			lpMailboxMDB->Release();
		}
	}
}

void CMainDlg::OnOpenSelectedStoreDeletedFolders()
{
	HRESULT hRes = S_OK;
	LPMDB lpMDB = NULL;
	LPSBinary lpItemEID = NULL;
	int iItem = -1;
	SortListData* lpListData = NULL;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	do
	{
		hRes = S_OK;
		lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
		if (lpListData && lpListData->Contents())
		{
			lpItemEID = lpListData->Contents()->lpEntryID;
			if (lpItemEID)
			{
				EC_H(CallOpenMsgStore(
					lpMAPISession,
					(ULONG_PTR)m_hWnd,
					lpItemEID,
					MDB_WRITE | MDB_ONLINE,
					&lpMDB));

				if (lpMDB)
				{
					EC_H(DisplayObject(
						lpMDB,
						NULL,
						otStoreDeletedItems,
						this));

					lpMDB->Release();
					lpMDB = NULL;
				}
			}
		}
	} while (iItem != -1);

	if (lpMDB) lpMDB->Release();
}

void CMainDlg::OnDumpStoreContents()
{
	HRESULT hRes = S_OK;
	LPMDB lpMDB = NULL;
	LPSBinary lpItemEID = NULL;
	int iItem = -1;
	SortListData* lpListData = NULL;
	WCHAR szDir[MAX_PATH];

	if (!m_lpContentsTableListCtrl || !m_lpMapiObjects) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	do
	{
		hRes = S_OK;
		lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(&iItem);
		if (lpListData && lpListData->Contents())
		{
			lpItemEID = lpListData->Contents()->lpEntryID;
			if (lpItemEID)
			{
				EC_H(CallOpenMsgStore(
					lpMAPISession,
					(ULONG_PTR)m_hWnd,
					lpItemEID,
					MDB_WRITE,
					&lpMDB));
				if (lpMDB)
				{
					WC_H(GetDirectoryPath(m_hWnd, szDir));
					if (S_OK == hRes && szDir[0])
					{
						CWaitCursor Wait; // Change the mouse to an hourglass while we work.

						CDumpStore MyDumpStore;
						MyDumpStore.InitFolderPathRoot(szDir);
						MyDumpStore.InitMDB(lpMDB);
						MyDumpStore.ProcessStore();
					}

					lpMDB->Release();
					lpMDB = NULL;
				}
			}
		}
	} while (iItem != -1);
}

void CMainDlg::OnDumpServerContents()
{
	HRESULT hRes = S_OK;
	WCHAR szDir[MAX_PATH] = { 0 };

	if (!m_lpMapiObjects) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	wstring szServerName = GetServerName(lpMAPISession);

	CEditor MyData(
		this,
		IDS_DUMPSERVERPRIVATESTORE,
		IDS_SERVERNAMEPROMPT,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, CreateSingleLinePane(IDS_SERVERNAME, szServerName, false));

	WC_H(MyData.DisplayDialog());

	if (S_OK == hRes)
	{
		WC_H(GetDirectoryPath(m_hWnd, szDir));

		if (S_OK == hRes && szDir[0])
		{
			CWaitCursor Wait; // Change the mouse to an hourglass while we work.

			CDumpStore MyDumpStore;
			MyDumpStore.InitMailboxTablePathRoot(szDir);
			MyDumpStore.InitSession(lpMAPISession);
			MyDumpStore.ProcessMailboxTable(MyData.GetStringW(0));
		}
	}
}

void CMainDlg::OnLogoff()
{
	if (m_lpContentsTableListCtrl && m_lpContentsTableListCtrl->IsLoading()) return;
	HRESULT hRes = S_OK;

	if (!m_lpMapiObjects) return;

	OnCloseAddressBook();

	// We do this first to free up any stray session pointers
	EC_H(m_lpContentsTableListCtrl->SetContentsTable(
		NULL,
		dfNormal,
		MAPI_STORE));

	m_lpMapiObjects->Logoff(m_hWnd, MAPI_LOGOFF_UI);
}

void CMainDlg::OnLogoffWithFlags()
{
	if (m_lpContentsTableListCtrl && m_lpContentsTableListCtrl->IsLoading()) return;
	HRESULT hRes = S_OK;

	if (!m_lpMapiObjects) return;

	CEditor MyData(
		this,
		IDS_LOGOFF,
		IDS_LOGOFFPROMPT,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, CreateSingleLinePane(IDS_FLAGSINHEX, false));
	MyData.SetHex(0, MAPI_LOGOFF_UI);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		OnCloseAddressBook();

		// We do this first to free up any stray session pointers
		EC_H(m_lpContentsTableListCtrl->SetContentsTable(
			NULL,
			dfNormal,
			MAPI_STORE));

		m_lpMapiObjects->Logoff(m_hWnd, MyData.GetHex(0));
	}
}

void CMainDlg::OnLogon()
{
	if (m_lpContentsTableListCtrl && m_lpContentsTableListCtrl->IsLoading()) return;
	ULONG ulFlags = MAPI_EXTENDED | MAPI_ALLOW_OTHERS | MAPI_NEW_SESSION | MAPI_LOGON_UI | MAPI_EXPLICIT_PROFILE; // display a profile picker box

	if (!m_lpMapiObjects) return;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	m_lpMapiObjects->MAPILogonEx(
		m_hWnd, // handle of current window
		NULL, // profile name
		ulFlags);
}

void CMainDlg::OnLogonWithFlags()
{
	HRESULT hRes = S_OK;

	if (!m_lpMapiObjects) return;

	CEditor MyData(
		this,
		IDS_PROFFORMAPILOGON,
		IDS_PROFFORMAPILOGONPROMPT,
		2,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, CreateSingleLinePane(IDS_PROFILE, false));
	MyData.InitPane(1, CreateSingleLinePane(IDS_FLAGSINHEX, false));
	MyData.SetHex(1, MAPI_EXTENDED | MAPI_EXPLICIT_PROFILE | MAPI_ALLOW_OTHERS | MAPI_NEW_SESSION | MAPI_LOGON_UI);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		wstring szProfile = MyData.GetStringW(0);

		m_lpMapiObjects->MAPILogonEx(
			m_hWnd, // handle of current window (from def'n of CWnd)
			szProfile.empty() ? nullptr : (LPTSTR)szProfile.c_str(), // profile name
			MyData.GetHex(1) | MAPI_UNICODE);
	}
	else
	{
		DebugPrint(DBGGeneric, L"MAPILogonEx call cancelled.\n");
	}
}

void CMainDlg::OnResolveMessageClass()
{
	HRESULT hRes = S_OK;
	if (!m_lpMapiObjects || !m_lpPropDisplay) return;

	LPMAPIFORMINFO lpMAPIFormInfo = NULL;
	ResolveMessageClass(m_lpMapiObjects, NULL, &lpMAPIFormInfo);
	if (lpMAPIFormInfo)
	{
		EC_H(m_lpPropDisplay->SetDataSource(lpMAPIFormInfo, NULL, false));
		lpMAPIFormInfo->Release();
	}
}

void CMainDlg::OnSelectForm()
{
	HRESULT hRes = S_OK;
	LPMAPIFORMINFO lpMAPIFormInfo = NULL;

	if (!m_lpMapiObjects || !m_lpPropDisplay) return;
	SelectForm(m_hWnd, m_lpMapiObjects, NULL, &lpMAPIFormInfo);
	if (lpMAPIFormInfo)
	{
		// TODO: Put some code in here which works with the returned Form Info pointer
		EC_H(m_lpPropDisplay->SetDataSource(lpMAPIFormInfo, NULL, false));
		lpMAPIFormInfo->Release();
	}
}

void CMainDlg::OnSelectFormContainer()
{
	HRESULT hRes = S_OK;
	LPMAPIFORMMGR lpMAPIFormMgr = NULL;
	LPMAPIFORMCONTAINER lpMAPIFormContainer = NULL;

	if (!m_lpMapiObjects || !m_lpPropDisplay) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	EC_MAPI(MAPIOpenFormMgr(lpMAPISession, &lpMAPIFormMgr));
	if (lpMAPIFormMgr)
	{
		CEditor MyFlags(
			this,
			IDS_SELECTFORM,
			IDS_SELECTFORMPROMPT,
			(ULONG)1,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyFlags.InitPane(0, CreateSingleLinePane(IDS_FLAGS, false));
		MyFlags.SetHex(0, MAPIFORM_SELECT_ALL_REGISTRIES);

		WC_H(MyFlags.DisplayDialog());
		if (S_OK == hRes)
		{
			ULONG ulFlags = MyFlags.GetHex(0);
			EC_H_CANCEL(lpMAPIFormMgr->SelectFormContainer(
				(ULONG_PTR)m_hWnd,
				ulFlags,
				&lpMAPIFormContainer));
			if (lpMAPIFormContainer)
			{
				new CFormContainerDlg(
					m_lpParent,
					m_lpMapiObjects,
					lpMAPIFormContainer);

				lpMAPIFormContainer->Release();
			}
		}

		lpMAPIFormMgr->Release();
	}
}

void CMainDlg::OnOpenFormContainer()
{
	HRESULT hRes = S_OK;
	LPMAPIFORMMGR lpMAPIFormMgr = NULL;
	LPMAPIFORMCONTAINER lpMAPIFormContainer = NULL;

	if (!m_lpMapiObjects || !m_lpPropDisplay) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	EC_MAPI(MAPIOpenFormMgr(lpMAPISession, &lpMAPIFormMgr));
	if (lpMAPIFormMgr)
	{
		CEditor MyFlags(
			this,
			IDS_OPENFORMCONTAINER,
			IDS_OPENFORMCONTAINERPROMPT,
			(ULONG)1,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyFlags.InitPane(0, CreateSingleLinePane(IDS_HFRMREG, false));
		MyFlags.SetHex(0, MAPIFORM_SELECT_ALL_REGISTRIES);

		WC_H(MyFlags.DisplayDialog());
		if (S_OK == hRes)
		{
			HFRMREG hFrmReg = MyFlags.GetHex(0);
			EC_H_CANCEL(lpMAPIFormMgr->OpenFormContainer(
				hFrmReg,
				NULL,
				&lpMAPIFormContainer));
			if (lpMAPIFormContainer)
			{
				new CFormContainerDlg(
					m_lpParent,
					m_lpMapiObjects,
					lpMAPIFormContainer);

				lpMAPIFormContainer->Release();
			}
		}

		lpMAPIFormMgr->Release();
	}
}

void CMainDlg::OnMAPIOpenLocalFormContainer()
{
	if (!m_lpMapiObjects) return;
	m_lpMapiObjects->MAPIInitialize(NULL);
	if (!m_lpMapiObjects->bMAPIInitialized()) return;

	HRESULT hRes = S_OK;
	LPMAPIFORMCONTAINER lpMAPILocalFormContainer = NULL;

	EC_MAPI(MAPIOpenLocalFormContainer(&lpMAPILocalFormContainer));

	if (lpMAPILocalFormContainer)
	{
		new CFormContainerDlg(
			m_lpParent,
			m_lpMapiObjects,
			lpMAPILocalFormContainer);

		lpMAPILocalFormContainer->Release();
	}
}

void CMainDlg::OnLoadMAPI()
{
	HRESULT hRes = S_OK;
	TCHAR szDLLPath[MAX_PATH] = { 0 };
	UINT cchDllPath = NULL;
	CEditor MyData(
		this,
		IDS_LOADMAPI,
		IDS_LOADMAPIPROMPT,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	WC_D(cchDllPath, GetSystemDirectory(szDLLPath, _countof(szDLLPath)));
	if (cchDllPath < _countof(szDLLPath))
	{
		WC_H(StringCchCat(szDLLPath, _countof(szDLLPath), _T("\\mapi32.dll"))); // STRING_OK
		MyData.InitPane(0, CreateSingleLinePane(IDS_PATH, LPCTSTRToWstring(szDLLPath), false));
	}

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		HMODULE hMAPI = NULL;
		UnLoadPrivateMAPI();
		EC_D(hMAPI, MyLoadLibraryW(MyData.GetStringW(0)));
		SetMAPIHandle(hMAPI);
	}
}

void CMainDlg::OnUnloadMAPI()
{
	HRESULT hRes = S_OK;

	CEditor MyData(
		this,
		IDS_UNLOADMAPI,
		IDS_MAPICRASHWARNING,
		0,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		UnLoadPrivateMAPI();
	}
}

void CMainDlg::OnDisplayMAPIPath()
{
	HRESULT hRes = S_OK;
	TCHAR szMAPIPath[MAX_PATH] = { 0 };

	DebugPrint(DBGGeneric, L"OnDisplayMAPIPath()\n");
	HMODULE hMAPI = GetMAPIHandle();

	CEditor MyData(
		this,
		IDS_MAPIPATHTITLE,
		NULL,
		3,
		CEDITOR_BUTTON_OK);
	MyData.InitPane(0, CreateSingleLinePane(IDS_FILEPATH, true));
	if (hMAPI)
	{
		DWORD dw = GetModuleFileName(hMAPI, szMAPIPath, _countof(szMAPIPath));
		if (dw)
		{
			MyData.SetString(0, szMAPIPath);
		}
	}

	MyData.InitPane(1, CreateCheckPane(IDS_REGKEY_FORCEOUTLOOKMAPI, 0 != RegKeys[regkeyFORCEOUTLOOKMAPI].ulCurDWORD, true));
	MyData.InitPane(2, CreateCheckPane(IDS_REGKEY_FORCESYSTEMMAPI, 0 != RegKeys[regkeyFORCESYSTEMMAPI].ulCurDWORD, true));

	WC_H(MyData.DisplayDialog());
}

void CMainDlg::OnMAPIInitialize()
{
	HRESULT hRes = S_OK;
	if (!m_lpMapiObjects) return;

	CEditor MyData(
		this,
		IDS_MAPIINIT,
		IDS_MAPIINITPROMPT,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, CreateSingleLinePane(IDS_FLAGSINHEX, false));
	MyData.SetHex(0, NULL);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		m_lpMapiObjects->MAPIInitialize(MyData.GetHex(0));
	}
}

void CMainDlg::OnMAPIUninitialize()
{
	HRESULT hRes = S_OK;
	if (!m_lpMapiObjects) return;

	CEditor MyData(
		this,
		IDS_MAPIUNINIT,
		IDS_MAPICRASHWARNING,
		0,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		m_lpMapiObjects->MAPIUninitialize();
	}
}

void CMainDlg::OnFastShutdown()
{
	HRESULT hRes = S_OK;
	if (!m_lpMapiObjects) return;
	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	LPMAPICLIENTSHUTDOWN lpClientShutdown = NULL;

	EC_H_MSG(lpMAPISession->QueryInterface(IID_IMAPIClientShutdown, (LPVOID*)&lpClientShutdown), IDS_EDNOMAPICLIENTSHUTDOWN);

	if (SUCCEEDED(hRes) && lpClientShutdown)
	{
		EC_H_MSG(lpClientShutdown->QueryFastShutdown(), IDS_EDQUERYFASTSHUTDOWNFAILED);

		HRESULT hResNotify = NULL;

		hResNotify = lpClientShutdown->NotifyProcessShutdown();
		CHECKHRESMSG(hResNotify, IDS_EDNOTIFYPROCESSSHUTDOWNFAILED);

		if (SUCCEEDED(hRes))
		{
			EC_H_MSG(lpClientShutdown->DoFastShutdown(), IDS_EDDOFASTSHUTDOWNFAILED);

			if (SUCCEEDED(hRes))
			{
				ErrDialog(__FILE__, __LINE__, IDS_EDDOFASTSHUTDOWNSUCCEEDED);
			}
		}
	}

	if (lpClientShutdown) lpClientShutdown->Release();
}

void CMainDlg::OnQueryDefaultMessageOpt()
{
	HRESULT hRes = S_OK;
	if (!m_lpMapiObjects) return;
	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	CEditor MyData(
		this,
		IDS_QUERYDEFMSGOPT,
		IDS_ADDRESSTYPEPROMPT,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, CreateSingleLinePane(IDS_ADDRESSTYPE, wstring(L"EX"), false)); // STRING_OK

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		ULONG cValues = NULL;
		LPSPropValue lpOptions = NULL;
		string adrType = wstringTostring(MyData.GetStringW(0));

		EC_MAPI(lpMAPISession->QueryDefaultMessageOpt(
			reinterpret_cast<LPTSTR>(const_cast<LPSTR>(adrType.c_str())),
			NULL, // API doesn't like Unicode
			&cValues,
			&lpOptions));

		DebugPrintProperties(DBGGeneric, cValues, lpOptions, NULL);

		if (SUCCEEDED(hRes))
		{
			CEditor MyResult(
				this,
				IDS_QUERYDEFMSGOPT,
				IDS_RESULTOFCALLPROMPT,
				lpOptions ? 2 : 1,
				CEDITOR_BUTTON_OK);
			MyResult.InitPane(0, CreateSingleLinePane(IDS_COUNTOPTIONS, true));
			MyResult.SetHex(0, cValues);

			if (lpOptions)
			{
				ULONG i = 0;
				wstring szPropString;
				wstring szProp;
				wstring szAltProp;
				for (i = 0; i < cValues; i++)
				{
					InterpretProp(&lpOptions[i], &szProp, &szAltProp);
					szPropString += formatmessage(IDS_OPTIONSSTRUCTURE,
						i,
						TagToString(lpOptions[i].ulPropTag, NULL, false, true).c_str(),
						szProp.c_str(),
						szAltProp.c_str());
				}

				MyResult.InitPane(1, CreateMultiLinePane(IDS_OPTIONS, szPropString, true));
			}

			WC_H(MyResult.DisplayDialog());
		}

		MAPIFreeBuffer(lpOptions);
	}
}

void CMainDlg::OnQueryDefaultRecipOpt()
{
	HRESULT hRes = S_OK;
	if (!m_lpMapiObjects) return;
	LPADRBOOK lpAddrBook = m_lpMapiObjects->GetAddrBook(true); // do not release
	if (!lpAddrBook) return;

	CEditor MyData(
		this,
		IDS_QUERYDEFRECIPOPT,
		IDS_ADDRESSTYPEPROMPT,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, CreateSingleLinePane(IDS_ADDRESSTYPE, wstring(L"EX"), false)); // STRING_OK

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		ULONG cValues = NULL;
		LPSPropValue lpOptions = NULL;

		string adrType = wstringTostring(MyData.GetStringW(0));

		EC_MAPI(lpAddrBook->QueryDefaultRecipOpt(
			reinterpret_cast<LPTSTR>(const_cast<LPSTR>(adrType.c_str())),
			NULL, // API doesn't like Unicode
			&cValues,
			&lpOptions));

		DebugPrintProperties(DBGGeneric, cValues, lpOptions, NULL);

		if (SUCCEEDED(hRes))
		{
			CEditor MyResult(
				this,
				IDS_QUERYDEFRECIPOPT,
				IDS_RESULTOFCALLPROMPT,
				lpOptions ? 2 : 1,
				CEDITOR_BUTTON_OK);
			MyResult.InitPane(0, CreateSingleLinePane(IDS_COUNTOPTIONS, true));
			MyResult.SetHex(0, cValues);

			if (lpOptions)
			{
				ULONG i = 0;
				wstring szPropString;
				wstring szProp;
				wstring szAltProp;
				for (i = 0; i < cValues; i++)
				{
					InterpretProp(&lpOptions[i], &szProp, &szAltProp);
					szPropString += formatmessage(IDS_OPTIONSSTRUCTURE,
						i,
						TagToString(lpOptions[i].ulPropTag, NULL, false, true).c_str(),
						szProp.c_str(),
						szAltProp.c_str());
				}

				MyResult.InitPane(1, CreateMultiLinePane(IDS_OPTIONS, szPropString, true));
			}

			WC_H(MyResult.DisplayDialog());
		}

		MAPIFreeBuffer(lpOptions);
	}
}

void CMainDlg::OnQueryIdentity()
{
	HRESULT hRes = S_OK;
	ULONG cbEntryID = 0;
	LPENTRYID lpEntryID = NULL;
	LPMAPIPROP lpIdentity = NULL;

	if (!m_lpMapiObjects || !m_lpPropDisplay) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	EC_MAPI(lpMAPISession->QueryIdentity(
		&cbEntryID,
		&lpEntryID));

	if (cbEntryID && lpEntryID)
	{
		CEditor MyPrompt(
			this,
			IDS_QUERYID,
			IDS_QUERYIDPROMPT,
			1,
			CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
		MyPrompt.InitPane(0, CreateCheckPane(IDS_DISPLAYDETAILSDLG, false, false));
		WC_H(MyPrompt.DisplayDialog());
		if (S_OK == hRes)
		{
			if (MyPrompt.GetCheck(0))
			{
				LPADRBOOK lpAB = m_lpMapiObjects->GetAddrBook(true); // do not release
				ULONG_PTR ulUIParam = (ULONG_PTR)(void*)m_hWnd;

				EC_H_CANCEL(lpAB->Details(
					&ulUIParam,
					NULL,
					NULL,
					cbEntryID,
					lpEntryID,
					NULL,
					NULL,
					NULL,
					DIALOG_MODAL)); // API doesn't like Unicode
			}
			else
			{
				EC_H(CallOpenEntry(
					NULL,
					NULL,
					NULL,
					lpMAPISession,
					cbEntryID,
					lpEntryID,
					NULL,
					NULL,
					NULL,
					(LPUNKNOWN*)&lpIdentity));
				if (lpIdentity)
				{
					EC_H(m_lpPropDisplay->SetDataSource(lpIdentity, NULL, true));

					lpIdentity->Release();
				}
			}
		}

		MAPIFreeBuffer(lpEntryID);
	}
}

void CMainDlg::OnSetDefaultStore()
{
	HRESULT hRes = S_OK;
	LPSBinary lpItemEID = NULL;
	SortListData* lpListData = NULL;

	if (!m_lpMapiObjects) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(NULL);
	if (lpListData && lpListData->Contents())
	{
		lpItemEID = lpListData->Contents()->lpEntryID;
		if (lpItemEID)
		{
			CEditor MyData(
				this,
				IDS_SETDEFSTORE,
				IDS_SETDEFSTOREPROMPT,
				1,
				CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
			MyData.InitPane(0, CreateSingleLinePane(IDS_FLAGS, false));
			MyData.SetHex(0, MAPI_DEFAULT_STORE);

			WC_H(MyData.DisplayDialog());
			if (S_OK == hRes)
			{
				EC_MAPI(lpMAPISession->SetDefaultStore(
					MyData.GetHex(0),
					lpItemEID->cb,
					(LPENTRYID)lpItemEID->lpb));
			}
		}
	}
}

void CMainDlg::OnIsAttachmentBlocked()
{
	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	HRESULT hRes = S_OK;

	CEditor MyData(
		this,
		IDS_ISATTBLOCKED,
		IDS_ENTERFILENAME,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, CreateSingleLinePane(IDS_FILENAME, false));

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		bool bBlocked = false;
		EC_H(IsAttachmentBlocked(lpMAPISession, MyData.GetStringW(0), &bBlocked));
		if (SUCCEEDED(hRes))
		{
			CEditor MyResult(
				this,
				IDS_ISATTBLOCKED,
				IDS_RESULTOFCALLPROMPT,
				1,
				CEDITOR_BUTTON_OK);
			wstring szResult = loadstring(bBlocked ? IDS_TRUE : IDS_FALSE);
			MyResult.InitPane(0, CreateSingleLinePane(IDS_RESULT, szResult, true));

			WC_H(MyResult.DisplayDialog());
		}
	}
}

void CMainDlg::OnShowProfiles()
{
	HRESULT hRes = S_OK;
	LPMAPITABLE lpProfTable = NULL;

	if (!m_lpMapiObjects) return;
	m_lpMapiObjects->MAPIInitialize(NULL);

	LPPROFADMIN lpProfAdmin = NULL;
	EC_MAPI(MAPIAdminProfiles(0, &lpProfAdmin));
	if (!lpProfAdmin) return;

	EC_MAPI(lpProfAdmin->GetProfileTable(
		0, // fMapiUnicode is not supported
		&lpProfTable));

	if (lpProfTable)
	{
		new CProfileListDlg(
			m_lpParent,
			m_lpMapiObjects,
			lpProfTable);

		lpProfTable->Release();
	}

	lpProfAdmin->Release();
}

void CMainDlg::OnLaunchProfileWizard()
{
	if (!m_lpMapiObjects) return;
	m_lpMapiObjects->MAPIInitialize(NULL);
	if (!m_lpMapiObjects->bMAPIInitialized()) return;

	HRESULT hRes = S_OK;
	CEditor MyData(
		this,
		IDS_LAUNCHPROFWIZ,
		IDS_LAUNCHPROFWIZPROMPT,
		2,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);
	MyData.InitPane(0, CreateSingleLinePane(IDS_FLAGS, false));
	MyData.SetHex(0, MAPI_PW_LAUNCHED_BY_CONFIG);
	MyData.InitPane(1, CreateSingleLinePane(IDS_SERVICE, wstring(L"MSEMS"), false)); // STRING_OK

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		wstring szProfName = LaunchProfileWizard(
			m_hWnd,
			MyData.GetHex(0),
			wstringTostring(MyData.GetStringW(1)));
	}
}

void CMainDlg::OnGetMAPISVC()
{
	DisplayMAPISVCPath(this);
}

void CMainDlg::OnAddServicesToMAPISVC()
{
	AddServicesToMapiSvcInf();
}

void CMainDlg::OnRemoveServicesFromMAPISVC()
{
	RemoveServicesFromMapiSvcInf();
}

void CMainDlg::OnStatusTable()
{
	HRESULT hRes = S_OK;
	LPMAPITABLE lpMAPITable = NULL;

	if (!m_lpMapiObjects) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	EC_MAPI(lpMAPISession->GetStatusTable(
		NULL, // This table does not support MAPI_UNICODE!
		&lpMAPITable));
	if (lpMAPITable)
	{

		EC_H(DisplayTable(
			lpMAPITable,
			otStatus,
			this));
		lpMAPITable->Release();
	}
}

void CMainDlg::OnDisplayItem()
{
	// This is an example of how to override the default OnDisplayItem
	CContentsTableDlg::OnDisplayItem();
}

void CMainDlg::OnDisplayMailboxTable()
{
	if (!m_lpParent || !m_lpMapiObjects) return;

	DisplayMailboxTable(m_lpParent, m_lpMapiObjects);
}

void CMainDlg::OnDisplayPublicFolderTable()
{
	if (!m_lpParent || !m_lpMapiObjects) return;

	DisplayPublicFolderTable(m_lpParent, m_lpMapiObjects);
}

void CMainDlg::OnViewMSGProperties()
{
	if (!m_lpMapiObjects || !m_lpPropDisplay) return;
	m_lpMapiObjects->MAPIInitialize(NULL);
	if (!m_lpMapiObjects->bMAPIInitialized()) return;

	HRESULT hRes = S_OK;
	LPMESSAGE lpNewMessage = NULL;
	INT_PTR iDlgRet = IDOK;

	CStringW szFileSpec;
	CFileDialogExW dlgFilePicker;

	EC_B(szFileSpec.LoadString(IDS_MSGFILES));

	EC_D_DIALOG(dlgFilePicker.DisplayDialog(
		true,
		L"msg", // STRING_OK
		NULL,
		OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT | OFN_FILEMUSTEXIST,
		szFileSpec,
		this));
	if (iDlgRet == IDOK)
	{
		EC_H(LoadMSGToMessage(
			dlgFilePicker.GetFileName(),
			&lpNewMessage));
		if (lpNewMessage)
		{
			WC_H(DisplayObject(lpNewMessage, MAPI_MESSAGE, otDefault, this));
			lpNewMessage->Release();
		}
	}
}

void CMainDlg::OnConvertMSGToEML()
{
	if (!m_lpMapiObjects) return;
	m_lpMapiObjects->MAPIInitialize(NULL);
	if (!m_lpMapiObjects->bMAPIInitialized()) return;

	HRESULT hRes = S_OK;
	ULONG ulConvertFlags = CCSF_SMTP;
	ENCODINGTYPE et = IET_UNKNOWN;
	MIMESAVETYPE mst = USE_DEFAULT_SAVETYPE;
	ULONG ulWrapLines = USE_DEFAULT_WRAPPING;
	bool bDoAdrBook = false;

	WC_H(GetConversionToEMLOptions(this, &ulConvertFlags, &et, &mst, &ulWrapLines, &bDoAdrBook));
	if (S_OK == hRes)
	{
		LPADRBOOK lpAdrBook = NULL;
		if (bDoAdrBook) lpAdrBook = m_lpMapiObjects->GetAddrBook(true); // do not release

		INT_PTR iDlgRet = IDOK;

		CStringW szFileSpec;
		CFileDialogExW dlgFilePickerMSG;
		EC_B(szFileSpec.LoadString(IDS_MSGFILES));

		EC_D_DIALOG(dlgFilePickerMSG.DisplayDialog(
			true,
			L"msg", // STRING_OK
			NULL,
			OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT | OFN_FILEMUSTEXIST,
			szFileSpec,
			this));
		if (iDlgRet == IDOK)
		{
			EC_B(szFileSpec.LoadString(IDS_EMLFILES));

			CFileDialogExW dlgFilePickerEML;

			EC_D_DIALOG(dlgFilePickerEML.DisplayDialog(
				true,
				L"eml", // STRING_OK
				NULL,
				OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
				szFileSpec,
				this));
			if (iDlgRet == IDOK)
			{
				EC_H(ConvertMSGToEML(
					dlgFilePickerMSG.GetFileName(),
					dlgFilePickerEML.GetFileName(),
					ulConvertFlags,
					et,
					mst,
					ulWrapLines,
					lpAdrBook));
			}
		}
	}
}

void CMainDlg::OnConvertEMLToMSG()
{
	if (!m_lpMapiObjects) return;
	m_lpMapiObjects->MAPIInitialize(NULL);
	if (!m_lpMapiObjects->bMAPIInitialized()) return;

	HRESULT hRes = S_OK;
	ULONG ulConvertFlags = CCSF_SMTP;
	bool bDoAdrBook = false;
	bool bDoApply = false;
	bool bUnicode = false;
	HCHARSET hCharSet = NULL;
	CSETAPPLYTYPE cSetApplyType = CSET_APPLY_UNTAGGED;
	WC_H(GetConversionFromEMLOptions(this, &ulConvertFlags, &bDoAdrBook, &bDoApply, &hCharSet, &cSetApplyType, &bUnicode));
	if (S_OK == hRes)
	{
		LPADRBOOK lpAdrBook = NULL;
		if (bDoAdrBook) lpAdrBook = m_lpMapiObjects->GetAddrBook(true); // do not release

		INT_PTR iDlgRet = IDOK;

		CStringW szFileSpec;
		EC_B(szFileSpec.LoadString(IDS_EMLFILES));

		CFileDialogExW dlgFilePickerEML;

		EC_D_DIALOG(dlgFilePickerEML.DisplayDialog(
			true,
			L"eml", // STRING_OK
			NULL,
			OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT | OFN_FILEMUSTEXIST,
			szFileSpec,
			this));
		if (iDlgRet == IDOK)
		{
			EC_B(szFileSpec.LoadString(IDS_MSGFILES));

			CFileDialogExW dlgFilePickerMSG;
			EC_D_DIALOG(dlgFilePickerMSG.DisplayDialog(
				true,
				L"msg", // STRING_OK
				NULL,
				OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
				szFileSpec,
				this));
			if (iDlgRet == IDOK)
			{
				EC_H(ConvertEMLToMSG(
					dlgFilePickerEML.GetFileName(),
					dlgFilePickerMSG.GetFileName(),
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
	if (!m_lpMapiObjects) return;
	m_lpMapiObjects->MAPIInitialize(NULL);
	if (!m_lpMapiObjects->bMAPIInitialized()) return;

	HRESULT hRes = S_OK;

	INT_PTR iDlgRet = IDOK;

	CStringW szFileSpec;
	CFileDialogExW dlgFilePickerMSG;
	EC_B(szFileSpec.LoadString(IDS_MSGFILES));

	EC_D_DIALOG(dlgFilePickerMSG.DisplayDialog(
		true,
		L"msg", // STRING_OK
		NULL,
		OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT | OFN_FILEMUSTEXIST,
		szFileSpec,
		this));
	if (iDlgRet == IDOK)
	{
		EC_B(szFileSpec.LoadString(IDS_XMLFILES));

		CFileDialogExW dlgFilePickerXML;

		EC_D_DIALOG(dlgFilePickerXML.DisplayDialog(
			true,
			L"xml", // STRING_OK
			NULL,
			OFN_HIDEREADONLY | OFN_OVERWRITEPROMPT,
			szFileSpec,
			this));
		if (iDlgRet == IDOK)
		{
			LPMESSAGE lpMessage = NULL;

			EC_H(LoadMSGToMessage(dlgFilePickerMSG.GetFileName(), &lpMessage));

			if (lpMessage)
			{
				CDumpStore MyDumpStore;
				MyDumpStore.InitMessagePath(dlgFilePickerXML.GetFileName());
				// Just assume this message might have attachments
				MyDumpStore.ProcessMessage(lpMessage, true, NULL);

				lpMessage->Release();
			}
		}
	}
}

void CMainDlg::OnComputeGivenStoreHash()
{
	HRESULT hRes = S_OK;
	SortListData* lpListData = NULL;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (!m_lpMapiObjects || !m_lpContentsTableListCtrl) return;

	LPMAPISESSION lpMAPISession = m_lpMapiObjects->GetSession(); // do not release
	if (!lpMAPISession) return;

	lpListData = m_lpContentsTableListCtrl->GetNextSelectedItemData(NULL);
	if (lpListData && lpListData->Contents())
	{
		LPSBinary lpItemEID = lpListData->Contents()->lpEntryID;

		if (lpItemEID)
		{
			LPSPropValue lpProp = NULL;
			LPSBinary lpServiceUID = NULL;
			LPSBinary lpProviderUID = NULL;
			lpProp = PpropFindProp(
				lpListData->lpSourceProps,
				lpListData->cSourceProps,
				PR_SERVICE_UID);
			if (lpProp && PT_BINARY == PROP_TYPE(lpProp->ulPropTag)) lpServiceUID = &lpProp->Value.bin;
			lpProp = PpropFindProp(
				lpListData->lpSourceProps,
				lpListData->cSourceProps,
				PR_MDB_PROVIDER);
			if (lpProp && PT_BINARY == PROP_TYPE(lpProp->ulPropTag)) lpProviderUID = &lpProp->Value.bin;

			MAPIUID emsmdbUID = { 0 };
			LPPROFSECT lpProfSect = NULL;
			bool fPublicExchangeStore = false;
			bool fPrivateExchangeStore = false;
			if (lpProviderUID)
			{
				fPublicExchangeStore = FExchangePublicStore((LPMAPIUID)lpProviderUID->lpb);
				fPrivateExchangeStore = FExchangePrivateStore((LPMAPIUID)lpProviderUID->lpb);
			}

			bool fCached = false;
			LPSPropValue lpConfigProp = NULL;
			LPSPropValue lpPathPropA = NULL;
			LPSPropValue lpPathPropW = NULL;
			LPSPropValue lpMappingSig = NULL;
			LPSTR szPath = NULL; // do not free
			LPWSTR wzPath = NULL; // do not free

			// Get profile section
			if (lpServiceUID)
			{
				WC_H(HrEmsmdbUIDFromStore(lpMAPISession,
					reinterpret_cast<LPMAPIUID>(lpServiceUID->lpb),
					&emsmdbUID));
				if (SUCCEEDED(hRes))
				{
					if (fIsSet(DBGGeneric))
					{
						wstring szGUID = GUIDToString((LPCGUID)&emsmdbUID);
						DebugPrint(DBGGeneric, L"CMainDlg::OnComputeGivenStoreHash, emsmdbUID from PR_EMSMDB_SECTION_UID = %ws\n", szGUID.c_str());
					}

					WC_MAPI(lpMAPISession->OpenProfileSection(&emsmdbUID, NULL, 0, &lpProfSect));
				}
			}

			if (!lpServiceUID || FAILED(hRes))
			{
				hRes = S_OK;
				// For Outlook 2003/2007, HrEmsmdbUIDFromStore may not succeed,
				// so use pbGlobalProfileSectionGuid instead
				WC_MAPI(lpMAPISession->OpenProfileSection((LPMAPIUID)pbGlobalProfileSectionGuid, NULL, 0, &lpProfSect));
			}

			if (lpProfSect)
			{
				hRes = S_OK;
				WC_MAPI(HrGetOneProp(lpProfSect, PR_PROFILE_CONFIG_FLAGS, &lpConfigProp));
				if (SUCCEEDED(hRes) && PROP_TYPE(lpConfigProp->ulPropTag) != PT_ERROR)
				{
					if (fPrivateExchangeStore)
					{
						fCached = ((lpConfigProp->Value.l & CONFIG_OST_CACHE_PRIVATE) != 0);
					}

					if (fPublicExchangeStore)
					{
						fCached = ((lpConfigProp->Value.l & CONFIG_OST_CACHE_PUBLIC) == CONFIG_OST_CACHE_PUBLIC);
					}
				}

				DebugPrint(DBGGeneric, L"CMainDlg::OnComputeGivenStoreHash, fPrivateExchangeStore = %d\n", fPrivateExchangeStore);
				DebugPrint(DBGGeneric, L"CMainDlg::OnComputeGivenStoreHash, fPublicExchangeStore = %d\n", fPublicExchangeStore);
				DebugPrint(DBGGeneric, L"CMainDlg::OnComputeGivenStoreHash, fCached = %d\n", fCached);

				if (fCached)
				{
					hRes = S_OK;
					WC_MAPI(HrGetOneProp(lpProfSect, PR_PROFILE_OFFLINE_STORE_PATH_W, &lpPathPropW));
					if (FAILED(hRes))
					{
						hRes = S_OK;
						WC_MAPI(HrGetOneProp(lpProfSect, PR_PROFILE_OFFLINE_STORE_PATH_A, &lpPathPropA));
					}

					if (SUCCEEDED(hRes))
					{
						if (lpPathPropW && lpPathPropW->Value.lpszW)
						{
							wzPath = lpPathPropW->Value.lpszW;
							DebugPrint(DBGGeneric, L"CMainDlg::OnComputeGivenStoreHash, PR_PROFILE_OFFLINE_STORE_PATH_W = %ws\n", wzPath);
						}
						else if (lpPathPropA && lpPathPropA->Value.lpszA)
						{
							szPath = lpPathPropA->Value.lpszA;
							DebugPrint(DBGGeneric, L"CMainDlg::OnComputeGivenStoreHash, PR_PROFILE_OFFLINE_STORE_PATH_A = %hs\n", szPath);
						}
					}
					// If this is an Exchange store with an OST path, it's an OST, so we get the mapping signature
					if ((fPrivateExchangeStore || fPublicExchangeStore) && (wzPath || szPath))
					{
						hRes = S_OK;
						WC_MAPI(HrGetOneProp(lpProfSect, PR_MAPPING_SIGNATURE, &lpMappingSig));
					}

					hRes = S_OK;
				}
			}

			DWORD dwSigHash = NULL;
			if (lpMappingSig && PT_BINARY == PROP_TYPE(lpMappingSig->ulPropTag))
			{
				dwSigHash = ComputeStoreHash(lpMappingSig->Value.bin.cb, lpMappingSig->Value.bin.lpb, NULL, NULL, fPublicExchangeStore);
			}

			DWORD dwEIDHash = ComputeStoreHash(lpItemEID->cb, lpItemEID->lpb, szPath, wzPath, fPublicExchangeStore);

			wstring szHash;
			if (dwSigHash)
			{
				szHash = formatmessage(IDS_STOREHASHDOUBLEVAL, dwEIDHash, dwSigHash);
			}
			else
			{
				szHash = formatmessage(IDS_STOREHASHVAL, dwEIDHash);
			}

			DebugPrint(DBGGeneric, L"CMainDlg::OnComputeGivenStoreHash, Entry ID hash = 0x%08X\n", dwEIDHash);
			if (dwSigHash)
				DebugPrint(DBGGeneric, L"CMainDlg::OnComputeGivenStoreHash, Mapping Signature hash = 0x%08X\n", dwSigHash);

			CEditor Result(
				this,
				IDS_STOREHASH,
				NULL,
				(ULONG)0,
				CEDITOR_BUTTON_OK);
			Result.SetPromptPostFix(szHash);
			(void)Result.DisplayDialog();

			MAPIFreeBuffer(lpMappingSig);
			MAPIFreeBuffer(lpPathPropA);
			MAPIFreeBuffer(lpPathPropW);
			MAPIFreeBuffer(lpConfigProp);
			if (lpProfSect) lpProfSect->Release();
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
		lpParams->lpMDB = (LPMDB)lpMAPIProp;
	}

	InvokeAddInMenu(lpParams);
}