#pragma once
class CParentWnd;

namespace cache
{
	class CMapiObjects;
}

#include <UI/Dialogs/HierarchyTable/HierarchyTableDlg.h>

namespace dialog
{
	class CMsgStoreDlg : public CHierarchyTableDlg
	{
	public:
		CMsgStoreDlg(
			_In_ ui::CParentWnd* pParentWnd,
			_In_ cache::CMapiObjects* lpMapiObjects,
			_In_opt_ LPMAPIPROP lpRootFolder,
			ULONG ulDisplayFlags);
		virtual ~CMsgStoreDlg();

	private:
		// Overrides from base class
		void OnDeleteSelectedItem() override;
		void HandleAddInMenuSingle(
			_In_ LPADDINMENUPARAMS lpParams,
			_In_ LPMAPIPROP lpMAPIProp,
			_In_ LPMAPICONTAINER lpContainer) override;
		void HandleCopy() override;
		_Check_return_ bool HandlePaste() override;
		void OnInitMenu(_In_ CMenu* pMenu) override;

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
		void OnDisplaySpecialFolder(ULONG ulFolder);
		void OnDisplayTasksFolder();
		void OnEmptyFolder();
		void OnOpenFormContainer();
		void OnSaveFolderContentsAsMSG();
		void OnSaveFolderContentsAsTextFiles();
		void OnResolveMessageClass();
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
		void OnExportMessages();

		DECLARE_MESSAGE_MAP()
	};
}