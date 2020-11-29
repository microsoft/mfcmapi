#pragma once
#include <UI/Dialogs/BaseDialog.h>
#include <UI/Controls/StyleTree/StyleTreeCtrl.h>

class CContentsTableListCtrl;

namespace dialog
{
	class AccountsDialog : public CBaseDialog
	{
	public:
		AccountsDialog(
			_In_ ui::CParentWnd* pParentWnd,
			_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
			_In_opt_ LPMAPISESSION lpMAPISession);
		~AccountsDialog();

	protected:
		// Overrides from base class
		void CreateDialogAndMenu(UINT nIDMenuResource);
		BOOL OnInitDialog() override;

	private:
		controls::StyleTreeCtrl m_lpAccountsList{};
		LPMAPISESSION m_lpMAPISession;

		// Menu items
		void OnRefreshView() override;

		DECLARE_MESSAGE_MAP()
	};
} // namespace dialog