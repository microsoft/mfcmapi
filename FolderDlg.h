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
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		_In_ LPMAPIFOLDER lpMAPIFolder,
		ULONG ulDisplayFlags);
	virtual ~CFolderDlg();

private:
	// Overrides from base class
	void EnableAddInMenus(_In_ CMenu* pMenu, ULONG ulMenu, _In_ LPMENUITEM lpAddInMenu, UINT uiEnable);
	void HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ LPMAPICONTAINER lpContainer);
	void HandleCopy();
	_Check_return_ BOOL HandleMenu(WORD wMenuSelect);
	_Check_return_ BOOL HandlePaste();
	void OnDeleteSelectedItem();
	void OnDisplayItem();
	void OnInitMenu(_In_ CMenu* pMenu);

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
	void OnResolveMessageClass();
	void OnSelectForm();
	void OnSendBulkMail();
	void OnSaveMessageToFile();
	void OnSetReadFlag();
	void OnSetMessageStatus();
	void OnGetMessageOptions();
	_Check_return_ HRESULT OnAbortSubmit(int iItem, _In_ SortListData* lpData);
	_Check_return_ HRESULT OnAttachmentProperties(int iItem, _In_ SortListData*	lpData);
	_Check_return_ HRESULT OnGetMessageStatus(int iItem, _In_ SortListData* lpData);
	_Check_return_ HRESULT OnOpenModal(int iItem, _In_ SortListData* lpData);
	_Check_return_ HRESULT OnOpenNonModal(int iItem, _In_ SortListData* lpData);
	_Check_return_ HRESULT OnRecipientProperties(int iItem, _In_ SortListData* lpData);
	_Check_return_ HRESULT OnResendSelectedItem(int iItem, _In_ SortListData* lpData);
	_Check_return_ HRESULT OnSaveAttachments(int iItem, _In_ SortListData* lpData);
	_Check_return_ HRESULT OnSubmitMessage(int iItem, _In_ SortListData* lpData);

	_Check_return_ BOOL MultiSelectComplex(WORD wMenuSelect);
	_Check_return_ BOOL MultiSelectSimple(WORD wMenuSelect);
	void NewSpecialItem(WORD wMenuSelect);
};