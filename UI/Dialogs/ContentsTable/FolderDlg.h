#pragma once
#include <UI/Dialogs/ContentsTable/ContentsTableDlg.h>

class CParentWnd;

namespace
{
	class CMapiObjects;
}

namespace dialog
{
	class CFolderDlg : public CContentsTableDlg
	{
	public:
		CFolderDlg(
			_In_ ui::CParentWnd* pParentWnd,
			_In_ cache::CMapiObjects* lpMapiObjects,
			_In_ LPMAPIPROP lpMAPIFolder,
			ULONG ulDisplayFlags);
		virtual ~CFolderDlg();

	private:
		// Overrides from base class
		void EnableAddInMenus(_In_ HMENU hMenu, ULONG ulMenu, _In_ LPMENUITEM lpAddInMenu, UINT uiEnable) override;
		void HandleAddInMenuSingle(
			_In_ LPADDINMENUPARAMS lpParams,
			_In_ LPMAPIPROP lpMAPIProp,
			_In_ LPMAPICONTAINER lpContainer) override;
		void HandleCopy() override;
		_Check_return_ bool HandleMenu(WORD wMenuSelect) override;
		_Check_return_ bool HandlePaste() override;
		void OnDeleteSelectedItem() override;
		void OnDisplayItem() override;
		void OnInitMenu(_In_ CMenu* pMenu) override;

		// Menu items
		void OnAddOneOffAddress();
		void OnDeleteAttachments();
		void OnExecuteVerbOnForm();
		void OnGetPropsUsingLongTermEID() const;
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
		void OnExportMessages();
		void OnSetReadFlag();
		void OnSetMessageStatus();
		void OnGetMessageOptions();
		void OnCreateMessageRestriction();
		void OnDisplayFolder(WORD wMenuSelect);
		_Check_return_ HRESULT OnAbortSubmit(int iItem, _In_ controls::sortlistdata::SortListData* lpData);
		_Check_return_ HRESULT OnAttachmentProperties(int iItem, _In_ controls::sortlistdata::SortListData* lpData);
		_Check_return_ HRESULT OnGetMessageStatus(int iItem, _In_ controls::sortlistdata::SortListData* lpData);
		_Check_return_ HRESULT OnOpenModal(int iItem, _In_ controls::sortlistdata::SortListData* lpData);
		_Check_return_ HRESULT OnOpenNonModal(int iItem, _In_ controls::sortlistdata::SortListData* lpData);
		_Check_return_ HRESULT OnRecipientProperties(int iItem, _In_ controls::sortlistdata::SortListData* lpData);
		_Check_return_ HRESULT OnResendSelectedItem(int iItem, _In_ controls::sortlistdata::SortListData* lpData);
		_Check_return_ HRESULT OnSaveAttachments(int iItem, _In_ controls::sortlistdata::SortListData* lpData);
		_Check_return_ HRESULT OnSubmitMessage(int iItem, _In_ controls::sortlistdata::SortListData* lpData);
		LPMESSAGE OpenMessage(int iSelectedItem, __mfcmapiModifyEnum bModify);

		_Check_return_ bool MultiSelectComplex(WORD wMenuSelect);
		_Check_return_ bool MultiSelectSimple(WORD wMenuSelect);
		void NewSpecialItem(WORD wMenuSelect) const;

		LPMAPIFOLDER m_lpFolder;
	};
} // namespace dialog