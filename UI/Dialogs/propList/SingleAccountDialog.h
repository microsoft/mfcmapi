#pragma once
#include <UI/Dialogs/BaseDialog.h>

class CContentsTableListCtrl;

namespace dialog
{
	class SingleAccountDialog : public CBaseDialog
	{
	public:
		SingleAccountDialog(
			_In_ ui::CParentWnd* pParentWnd,
			_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
			_In_opt_ LPMAPISESSION lpMAPISession);
		~SingleAccountDialog();

	protected:
		// Overrides from base class
		void CreateDialogAndMenu(UINT nIDMenuResource);
		BOOL OnInitDialog() override;

	private:
		LPMAPISESSION m_lpMAPISession;

		// Menu items
		void OnRefreshView() override;

		DECLARE_MESSAGE_MAP()
	};
} // namespace dialog