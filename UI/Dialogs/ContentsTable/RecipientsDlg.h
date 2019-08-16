#pragma once
#include <UI/Dialogs/ContentsTable/ContentsTableDlg.h>

namespace
{
	class CMapiObjects;
}

namespace dialog
{
	class CRecipientsDlg : public CContentsTableDlg
	{
	public:
		CRecipientsDlg(
			_In_ ui::CParentWnd* pParentWnd,
			_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
			_In_ LPMAPITABLE lpMAPITable,
			_In_ LPMAPIPROP lpMessage);
		virtual ~CRecipientsDlg();

	private:
		// Overrides from base class
		void OnDeleteSelectedItem() override;
		void OnInitMenu(_In_ CMenu* pMenu) override;
		_Check_return_ LPMAPIPROP OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify) override;

		// Menu items
		void OnModifyRecipients();
		void OnRecipOptions();
		void OnSaveChanges();
		void OnViewRecipientABEntry();

		LPMESSAGE m_lpMessage;
		bool m_bViewRecipientABEntry;

		DECLARE_MESSAGE_MAP()
	};
} // namespace dialog