#pragma once
// FolderDlg.h : header file

class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "ContentsTableDlg.h"

class CFolderDlg : public CContentsTableDlg
{
public:
	CFolderDlg(
		CParentWnd* pParentWnd,
		CMapiObjects* lpMapiObjects,
		LPMAPIFOLDER lpMAPIFolder,
		ULONG ulDisplayFlags);
	virtual ~CFolderDlg();

private:
	// Overrides from base class
	void EnableAddInMenus(CMenu* pMenu, ULONG ulMenu, LPMENUITEM lpAddInMenu, UINT uiEnable);
	void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);
	BOOL HandleCopy();
	BOOL HandleMenu(WORD wMenuSelect);
	BOOL HandlePaste();
	void OnDeleteSelectedItem();
	void OnDisplayItem();
	void OnInitMenu(CMenu* pMenu);

	// Menu items
	void OnAddOneOffAddress();
	void OnDeleteAttachments();
	void OnExecuteVerbOnForm();
	void OnGetPropsUsingLongTermEID();
	void OnLoadFromEML();
	void OnLoadFromMSG();
	void OnLoadFromTNEF();
	void OnManualResolve();
	void OnNewCustomForm();
	void OnNewMessage();
	void OnRemoveOneOff();
	void OnRTFSync();
	void OnSaveFolderContentsAsTextFiles();
	void OnSelectForm();
	void OnSendBulkMail();
	void OnSaveMessageToFile();
	void OnSetReadFlag();
	void OnSetMessageStatus();
	void OnGetMessageOptions();
	HRESULT OnAbortSubmit(int iItem, SortListData* lpData);
	HRESULT OnAttachmentProperties(int iItem, SortListData*	lpData);
	HRESULT OnGetMessageStatus(int iItem, SortListData* lpData);
	HRESULT OnOpenModal(int iItem, SortListData* lpData);
	HRESULT OnOpenNonModal(int iItem, SortListData* lpData);
	HRESULT OnRecipientProperties(int iItem, SortListData* lpData);
	HRESULT OnResendSelectedItem(int iItem, SortListData* lpData);
	HRESULT OnSaveAttachments(int iItem, SortListData* lpData);
	HRESULT OnSubmitMessage(int iItem, SortListData* lpData);

	BOOL MultiSelectComplex(WORD wMenuSelect);
	BOOL MultiSelectSimple(WORD wMenuSelect);
	void NewSpecialItem(WORD wMenuSelect);
};