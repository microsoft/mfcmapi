#pragma once
// FolderDlg.h : header file
//

class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "ContentsTableDlg.h"

/////////////////////////////////////////////////////////////////////////////
// CFolderDlg dialog

class CFolderDlg : public CContentsTableDlg
{
public:
	CFolderDlg(
		CParentWnd* pParentWnd,
		CMapiObjects *lpMapiObjects,
		LPMAPIFOLDER lpMAPIFolder,
		__mfcmapiAssociatedContentsEnum bShowingAssociatedContents,
		__mfcmapiDeletedItemsEnum bShowingDeletedItems);
	virtual ~CFolderDlg();

protected:
	afx_msg void OnInitMenu(CMenu* pMenu);

	void OnGetPropsUsingLongTermEID();
	void OnLoadFromMSG();
	void OnLoadFromTNEF();
	void OnLoadFromEML();
	void OnSelectForm();
	void OnManualResolve();
	void OnNewCustomForm();
	void OnNewMessage();
	void OnSaveFolderContentsAsTextFiles();
	void OnSendBulkMail();

	void OnAddOneOffAddress();
	void OnExecuteVerbOnForm();
	void OnRemoveOneOff();
	void OnRTFSync();
	void OnSaveMessageToFile();
	void OnSetReadFlag();
	void OnSetMessageStatus();
	void OnGetMessageOptions();

	void OnDeleteSelectedItem();
	virtual BOOL	HandleCopy();
	virtual BOOL	HandlePaste();

	virtual void	OnDisplayItem();

	HRESULT OnAttachmentProperties(int iItem, SortListData*	lpData);
	HRESULT OnDeleteAttachments();
	HRESULT OnGetMessageStatus(int iItem, SortListData* lpData);
	HRESULT OnOpenModal(int iItem, SortListData* lpData);
	HRESULT OnOpenNonModal(int iItem, SortListData* lpData);
	HRESULT OnRecipientProperties(int iItem, SortListData* lpData);
	HRESULT OnResendSelectedItem(int iItem, SortListData* lpData);
	HRESULT OnSaveAttachments(int iItem, SortListData* lpData);
	HRESULT OnSubmitMessage(int iItem, SortListData* lpData);
	HRESULT OnAbortSubmit(int iItem, SortListData* lpData);

	virtual BOOL HandleMenu(WORD wMenuSelect);

	virtual void EnableAddInMenus(CMenu* pMenu, ULONG ulMenu, LPMENUITEM lpAddInMenu, UINT uiEnable);
	virtual void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);
private:
	BOOL MultiSelectSimple(WORD wMenuSelect);
	BOOL MultiSelectComplex(WORD wMenuSelect);
	void NewSpecialItem(WORD wMenuSelect);
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
