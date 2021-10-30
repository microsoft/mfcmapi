#pragma once
#include <UI/Dialogs/BaseDialog.h>

class CContentsTableListCtrl;

namespace dialog
{
	class SingleRecipientDialog : public CBaseDialog
	{
	public:
		SingleRecipientDialog(_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects, _In_opt_ LPMAPIPROP lpMailUser);
		~SingleRecipientDialog();

	protected:
		// Overrides from base class
		void CreateDialogAndMenu(UINT nIDMenuResource);
		BOOL OnInitDialog() override;

	private:
		LPMAILUSER m_lpMailUser;

		// Menu items
		void OnRefreshView() override;

		DECLARE_MESSAGE_MAP()
	};
} // namespace dialog