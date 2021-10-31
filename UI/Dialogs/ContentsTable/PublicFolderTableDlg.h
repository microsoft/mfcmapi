#pragma once
#include <UI/Dialogs/ContentsTable/ContentsTableDlg.h>

namespace cache
{
	class CMapiObjects;
}

namespace dialog
{
	class CPublicFolderTableDlg : public CContentsTableDlg
	{
	public:
		CPublicFolderTableDlg(
			_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
			_In_ const std::wstring& lpszServerName,
			_In_ LPMAPITABLE lpMAPITable);
		~CPublicFolderTableDlg();

	private:
		// Overrides from base class
		void CreateDialogAndMenu(UINT nIDMenuResource) override;
		void OnCreatePropertyStringRestriction() override;
		void OnDisplayItem() override;
		_Check_return_ LPMAPIPROP OpenItemProp(int iSelectedItem, modifyType bModify) override;

		std::wstring m_lpszServerName;
	};
} // namespace dialog