#pragma once
#include <UI/Dialogs/ContentsTable/ContentsTableDlg.h>

namespace cache
{
	class CMapiObjects;
}

namespace dialog
{
	class CFormContainerDlg : public CContentsTableDlg
	{
	public:
		CFormContainerDlg(
			_In_ ui::CParentWnd* pParentWnd,
			_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
			_In_ LPMAPIFORMCONTAINER lpFormContainer);
		virtual ~CFormContainerDlg();

	private:
		// Overrides from base class
		void HandleAddInMenuSingle(
			_In_ LPADDINMENUPARAMS lpParams,
			_In_ LPMAPIPROP lpMAPIProp,
			_In_ LPMAPICONTAINER lpContainer) override;
		void OnDeleteSelectedItem() override;
		BOOL OnInitDialog() override;
		void OnInitMenu(_In_ CMenu* pMenu) override;
		void OnRefreshView() override;
		_Check_return_ LPMAPIPROP OpenItemProp(int iSelectedItem, modifyType bModify) override;

		// Menu items
		void OnCalcFormPropSet();
		void OnGetDisplay();
		void OnInstallForm();
		void OnRemoveForm();
		void OnResolveMessageClass();
		void OnResolveMultipleMessageClasses();

		LPMAPIFORMCONTAINER m_lpFormContainer;

		DECLARE_MESSAGE_MAP()
	};
} // namespace dialog