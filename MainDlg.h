#pragma once
// MainDlg.h : header file

class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "ContentsTableDlg.h"

class CMainDlg : public CContentsTableDlg
{
public:
	CMainDlg(
		CParentWnd* pParentWnd,
		CMapiObjects* lpMapiObjects);
	virtual ~CMainDlg();

	// public so CBaseDialog can call it
	void OnOpenMessageStoreTable();

private:
	// Overrides from base class
	void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);
	BOOL HandleMenu(WORD wMenuSelect);
	HRESULT OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, LPMAPIPROP* lppMAPIProp);
	void OnDisplayItem();
	void OnInitMenu(CMenu* pMenu);

	// Menu items
	void OnABHierarchy();
	void OnAddServicesToMAPISVC();
	void OnCloseAddressBook();
	void OnConvertMSGToEML();
	void OnConvertEMLToMSG();
	void OnDisplayMailboxTable();
	void OnDisplayPublicFolderTable();
	void OnDumpStoreContents();
	void OnDumpServerContents();
	void OnFastShutdown();
	void OnGetMAPISVC();
	void OnIsAttachmentBlocked();
	void OnLoadMAPI();
	void OnLaunchProfileWizard();
	void OnLogoff();
	void OnLogoffWithFlags();
	void OnLogon();
	void OnLogonAndDisplayStores();
	void OnLogonWithFlags();
	void OnMAPIInitialize();
	void OnMAPIOpenLocalFormContainer();
	void OnMAPIUninitialize();
	void OnOpenAddressBook();
	void OnOpenDefaultDir();
	void OnOpenDefaultMessageStore();
	void OnOpenFormContainer();
	void OnOpenMailboxWithDN();
	void OnOpenMessageStoreEID();
	void OnOpenOtherUsersMailboxFromGAL();
	void OnOpenPAB();
	void OnOpenPublicFolders();
	void OnOpenSelectedStoreDeletedFolders();
	void OnRemoveServicesFromMAPISVC();
	void OnQueryDefaultMessageOpt();
	void OnQueryDefaultRecipOpt();
	void OnQueryIdentity();
	void OnSelectForm();
	void OnSelectFormContainer();
	void OnSetDefaultStore();
	void OnStatusTable();
	void OnShowProfiles();
	void OnUnloadMAPI();
	void OnViewMSGProperties();

	DECLARE_MESSAGE_MAP()
};