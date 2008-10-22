#pragma once
// MsgStoreDlg.h : header file

class CHierarchyTableTreeCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "HierarchyTableDlg.h"

class CMsgStoreDlg : public CHierarchyTableDlg
{
public:
	CMsgStoreDlg(
		CParentWnd* pParentWnd,
		CMapiObjects* lpMapiObjects,
		LPMAPIFOLDER lpRootFolder,
		ULONG ulDisplayFlags);
	virtual ~CMsgStoreDlg();

private:
	// Overrides from base class
	void OnDeleteSelectedItem();
	void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);
	BOOL HandleCopy();
	BOOL HandlePaste();
	void OnInitMenu(CMenu* pMenu);

	// Menu items
	void OnCreateSubFolder();
	void OnDisplayACLTable();
	void OnDisplayAssociatedContents();
	void OnDisplayCalendarFolder();
	void OnDisplayContactsFolder();
	void OnDisplayDeletedContents();
	void OnDisplayDeletedSubFolders();
	void OnDisplayInbox();
	void OnDisplayMailboxTable();
	void OnDisplayOutgoingQueueTable();
	void OnDisplayReceiveFolderTable();
	void OnDisplayRulesTable();
	void OnDisplaySpecialFolder(ULONG ulPropTag);
	void OnDisplayTasksFolder();
	void OnEmptyFolder();
	void OnOpenFormContainer();
	void OnSaveFolderContentsAsMSG();
	void OnSaveFolderContentsAsTextFiles();
	void OnSelectForm();
	void OnSetReceiveFolder();
	void OnResendAllMessages();
	void OnResetPermissionsOnItems();
	void OnRestoreDeletedFolder();
	void OnValidateIPMSubtree();
	void OnPasteFolder();
	void OnPasteFolderContents();
	void OnPasteRules();
	void OnPasteMessages();

	DECLARE_MESSAGE_MAP()
};