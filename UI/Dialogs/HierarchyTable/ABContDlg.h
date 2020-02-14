#pragma once
#include <UI/Dialogs/HierarchyTable/HierarchyTableDlg.h>

namespace dialog
{
	class CAbContDlg : public CHierarchyTableDlg
	{
	public:
		CAbContDlg(_In_ ui::CParentWnd* pParentWnd, _In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects);
		~CAbContDlg();

	private:
		// Overrides from base class
		void HandleAddInMenuSingle(
			_In_ LPADDINMENUPARAMS lpParams,
			_In_ LPMAPIPROP lpMAPIProp,
			_In_ LPMAPICONTAINER lpContainer) override;

		// Menu items
		void OnSetDefaultDir();
		void OnSetPAB();

		DECLARE_MESSAGE_MAP()
	};
} // namespace dialog