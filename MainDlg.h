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
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects);
	virtual ~CMainDlg();

	// public so CBaseDialog can call it
	void OnOpenMessageStoreTable();

private:
	// Overrides from base class
	void HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ LPMAPICONTAINER lpContainer);
	_Check_return_ bool HandleMenu(WORD wMenuSelect);
	_Check_return_ HRESULT OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp);
	void OnDisplayItem();
	void OnInitMenu(_In_ CMenu* pMenu);

	void AddLoadMAPIMenus();
	bool InvokeLoadMAPIMenu(WORD wMenuSelect);

// Menu items
	void OnABHierarchy();
	void OnAddServicesToMAPISVC();
	void OnCloseAddressBook();
	void OnComputeGivenStoreHash();
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
	void OnResolveMessageClass();
	void OnSelectForm();
	void OnSelectFormContainer();
	void OnSetDefaultStore();
	void OnStatusTable();
	void OnShowProfiles();
	void OnUnloadMAPI();
	void OnViewMSGProperties();
	void OnDisplayMAPIPath();

	DECLARE_MESSAGE_MAP()
};