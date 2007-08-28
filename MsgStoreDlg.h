#pragma once
// MsgStoreDlg.h : header file
//

//forward definitions
class CHierarchyTableTreeCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "HierarchyTableDlg.h"

/////////////////////////////////////////////////////////////////////////////
// CMsgStoreDlg dialog

class CMsgStoreDlg : public CHierarchyTableDlg
{
	// Construction
public:
	CMsgStoreDlg(
		CParentWnd* pParentWnd,
		CMapiObjects *lpMapiObjects,
		LPMAPIFOLDER lpRootFolder,
		__mfcmapiDeletedItemsEnum bShowingDeletedFolders);
	virtual ~CMsgStoreDlg();

	// Implementation
protected:

	// Generated message map functions
	//{{AFX_MSG(CMsgStoreDlg)
	afx_msg void OnInitMenu(CMenu* pMenu);
	afx_msg void OnCreateSubFolder();
	afx_msg void OnDisplayACLTable();
	afx_msg void OnDisplayAssociatedContents();
	afx_msg void OnDisplayCalendarFolder();
	afx_msg void OnDisplayContactsFolder();
	afx_msg void OnDisplayDeletedContents();
	afx_msg void OnDisplayDeletedSubFolders();
	afx_msg void OnDisplayInbox();
	afx_msg void OnDisplayMailboxTable();
	afx_msg void OnDisplayOutgoingQueueTable();
	afx_msg void OnDisplayReceiveFolderTable();
	afx_msg void OnDisplayRulesTable();
	afx_msg void OnDisplayTasksFolder();
	afx_msg void OnEmptyFolder();
	afx_msg void OnOpenFormContainer();
	afx_msg void OnSelectForm();
	afx_msg void OnSaveFolderContentsAsMSG();
	afx_msg void OnSaveFolderContentsAsTextFiles();
	afx_msg void OnSetReceiveFolder();
	afx_msg void OnResendAllMessages();
	afx_msg void OnResetPermissionsOnItems();
	afx_msg void OnRestoreDeletedFolder();
	afx_msg void OnValidateIPMSubtree();
	afx_msg void OnDeleteSelectedItem();
	afx_msg void OnPasteRules();
	//}}AFX_MSG

	DECLARE_MESSAGE_MAP()

private:
	virtual BOOL	HandleCopy();
	virtual BOOL	HandlePaste();

	//used by OnPaste
	void OnPasteFolder();
	void OnPasteFolderContents();
	void OnPasteMessages();
	void OnDisplaySpecialFolder(ULONG ulPropTag);

	virtual void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
