#pragma once
#include <UI/Dialogs/ContentsTable/ContentsTableDlg.h>

namespace cache
{
	class CMapiObjects;
}

namespace dialog
{
	class CProviderTableDlg : public CContentsTableDlg
	{
	public:
		CProviderTableDlg(
			_In_ ui::CParentWnd* pParentWnd,
			_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
			_In_ LPMAPITABLE lpMAPITable,
			_In_ LPPROVIDERADMIN lpProviderAdmin);
		virtual ~CProviderTableDlg();

	private:
		// Overrides from base class
		void HandleAddInMenuSingle(
			_In_ LPADDINMENUPARAMS lpParams,
			_In_ LPMAPIPROP lpMAPIProp,
			_In_ LPMAPICONTAINER lpContainer) override;
		_Check_return_ LPMAPIPROP OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify) override;

		// Menu items
		void OnOpenProfileSection();

		LPPROVIDERADMIN m_lpProviderAdmin;

		DECLARE_MESSAGE_MAP()
	};
} // namespace dialog