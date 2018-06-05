#pragma once
#include <UI/Dialogs/ContentsTable/ContentsTableDlg.h>

namespace dialog
{
	class CAbDlg : public CContentsTableDlg
	{
	public:
		CAbDlg(
			_In_ ui::CParentWnd* pParentWnd,
			_In_ cache::CMapiObjects* lpMapiObjects,
			_In_ LPMAPIPROP lpUnk);
		virtual ~CAbDlg();

	private:
		// Overrides from base class
		void CreateDialogAndMenu(UINT nIDMenuResource) override;
		void HandleAddInMenuSingle(
			_In_ LPADDINMENUPARAMS lpParams,
			_In_ LPMAPIPROP lpMAPIProp,
			_In_ LPMAPICONTAINER lpContainer) override;
		void HandleCopy() override;
		_Check_return_ bool HandlePaste() override;
		void OnCreatePropertyStringRestriction() override;
		void OnDeleteSelectedItem() override;
		void OnDisplayDetails();
		void OnInitMenu(_In_ CMenu* pMenu) override;

		// Menu items
		void OnOpenContact();
		void OnOpenManager();
		void OnOpenOwner();

		LPABCONT m_lpAbCont;

		DECLARE_MESSAGE_MAP()
	};
}