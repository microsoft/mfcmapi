#pragma once
#include <UI/Dialogs/ContentsTable/ContentsTableDlg.h>

namespace cache
{
	class CMapiObjects;
}

namespace dialog
{
	class CMainDlg : public CContentsTableDlg
	{
	public:
		CMainDlg(_In_ ui::CParentWnd* pParentWnd, _In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects);
		~CMainDlg();

		// public so CBaseDialog can call it
		void OnOpenMessageStoreTable();

	private:
		// Overrides from base class
		void HandleAddInMenuSingle(
			_In_ LPADDINMENUPARAMS lpParams,
			_In_ LPMAPIPROP lpMAPIProp,
			_In_ LPMAPICONTAINER lpContainer) override;
		_Check_return_ bool HandleMenu(WORD wMenuSelect) override;
		_Check_return_ LPMAPIPROP OpenItemProp(int iSelectedItem, modifyType bModify) override;
		void OnDisplayItem() override;
		void OnInitMenu(_In_ CMenu* pMenu) override;

		void AddLoadMAPIMenus() const;
		bool InvokeLoadMAPIMenu(WORD wMenuSelect) const;

		// Menu items
		void OnABHierarchy();
		void OnAddServicesToMAPISVC();
		void OnCloseAddressBook();
		void OnComputeGivenStoreHash();
		void OnConvertMSGToEML();
		void OnConvertEMLToMSG();
		void OnConvertMSGToXML();
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
		void OnOpenPublicFolderWithDN();
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
} // namespace dialog