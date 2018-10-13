#pragma once
#include <UI/Dialogs/BaseDialog.h>
#include <UI/Controls/HierarchyTableTreeCtrl.h>

class CParentWnd;

namespace cache
{
	class CMapiObjects;
}

namespace dialog
{
	class CHierarchyTableDlg : public CBaseDialog
	{
	public:
		CHierarchyTableDlg(
			_In_ ui::CParentWnd* pParentWnd,
			_In_ cache::CMapiObjects* lpMapiObjects,
			UINT uidTitle,
			_In_opt_ LPMAPIPROP lpRootContainer,
			ULONG nIDContextMenu,
			ULONG ulAddInContext);
		virtual ~CHierarchyTableDlg();

	protected:
		// Overrides from base class
		void CreateDialogAndMenu(UINT nIDMenuResource);
		void OnInitMenu(_In_ CMenu* pMenu) override;
		// Get the current root container - does not addref
		LPMAPICONTAINER GetRootContainer() { return m_lpContainer; }
		void SetRootContainer(LPUNKNOWN container);

		controls::CHierarchyTableTreeCtrl m_lpHierarchyTableTreeCtrl;
		ULONG m_ulDisplayFlags;

	private:
		virtual void HandleAddInMenuSingle(
			_In_ LPADDINMENUPARAMS lpParams,
			_In_opt_ LPMAPIPROP lpMAPIProp,
			_In_opt_ LPMAPICONTAINER lpContainer);

		// Overrides from base class
		_Check_return_ bool HandleAddInMenu(WORD wMenuSelect) override;
		void OnCancel() override;
		BOOL OnInitDialog() override;
		void OnRefreshView() override;

		// Menu items
		void OnDisplayHierarchyTable();
		void OnDisplayItem();
		void OnEditSearchCriteria();

		UINT m_nIDContextMenu;

		LPMAPICONTAINER m_lpContainer;

		DECLARE_MESSAGE_MAP()
	};
}