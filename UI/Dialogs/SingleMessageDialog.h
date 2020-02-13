#pragma once
class CContentsTableListCtrl;

#include <UI/Dialogs/BaseDialog.h>

namespace dialog
{
	class SingleMessageDialog : public CBaseDialog
	{
	public:
		SingleMessageDialog(
			_In_ ui::CParentWnd* pParentWnd,
			_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
			_In_opt_ LPMAPIPROP lpMAPIProp);
		~SingleMessageDialog();

	protected:
		// Overrides from base class
		BOOL OnInitDialog() override;

	private:
		LPMESSAGE m_lpMessage;

		// Menu items
		void OnRefreshView() override;
		void OnAttachmentProperties();
		void OnRecipientProperties();
		void OnRTFSync();

		void OnTestEditBody();
		void OnTestEditHTML();
		void OnTestEditRTF();
		void OnSaveChanges();

		DECLARE_MESSAGE_MAP()
	};
} // namespace dialog