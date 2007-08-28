#pragma once
// MainDlg.h : header file
//

//forward definitions
class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "ContentsTableDlg.h"

/////////////////////////////////////////////////////////////////////////////
// CMainDlg dialog

class CMainDlg : public CContentsTableDlg
{
	// Construction
public:
	CMainDlg(
		CParentWnd* pParentWnd,
		CMapiObjects* lpMapiObjects);
	virtual ~CMainDlg();

	virtual HRESULT	OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, LPMAPIPROP* lppMAPIProp);

	afx_msg void OnOpenMessageStoreTable();
	// Implementation
protected:
	// Generated message map functions
	//{{AFX_MSG(CMainDlg)
	afx_msg void OnInitMenu(CMenu* pMenu);
	afx_msg void OnOpenPAB();
	afx_msg void OnABHierarchy();
	afx_msg void OnCloseAddressBook();
	afx_msg void OnConvertMSGToEML();
	afx_msg void OnConvertEMLToMSG();
	afx_msg void OnDisplayMailboxTable();
	afx_msg void OnDisplayPublicFolderTable();
	afx_msg void OnDumpStoreContents();
	afx_msg void OnDumpServerContents();
	afx_msg void OnLoadMAPI();
	afx_msg void OnLogoff();
	afx_msg void OnLogon();
	afx_msg void OnLogonAndDisplayStores();
	afx_msg void OnLogonWithFlags();
	afx_msg void OnIsAttachmentBlocked();
	afx_msg void OnMAPIInitialize();
	afx_msg void OnMAPIOpenLocalFormContainer();
	afx_msg void OnMAPIUninitialize();
	afx_msg void OnOpenAddressBook();
	afx_msg void OnOpenDefaultMessageStore();
	afx_msg void OnOpenFormContainer();
	afx_msg void OnOpenMailboxWithDN();
	afx_msg void OnOpenMessageStoreEID();
	afx_msg void OnOpenOtherUsersMailboxFromGAL();
	afx_msg void OnOpenPublicFolders();
	afx_msg void OnOpenSelectedStoreDeletedFolders();
	afx_msg void OnQueryDefaultMessageOpt();
	afx_msg void OnQueryDefaultRecipOpt();
	afx_msg void OnQueryIdentity();
	afx_msg void OnSelectForm();
	afx_msg void OnSelectFormContainer();
	afx_msg void OnSetDefaultStore();
	afx_msg void OnStatusTable();
	afx_msg void OnShowProfiles();
	afx_msg void OnUnloadMAPI();
	afx_msg void OnLaunchProfileWizard();
	afx_msg void OnGetMAPISVC();
	afx_msg void OnAddServicesToMAPISVC();
	afx_msg void OnRemoveServicesFromMAPISVC();
	afx_msg void OnViewMSGProperties();
	afx_msg void OnViewTNEFProperties();
	//}}AFX_MSG

	void OnDisplayItem();

	DECLARE_MESSAGE_MAP()

	virtual BOOL HandleMenu(WORD wMenuSelect);
	virtual void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
