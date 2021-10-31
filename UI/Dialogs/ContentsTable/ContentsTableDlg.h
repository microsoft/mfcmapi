#pragma once
#include <UI/Dialogs/BaseDialog.h>
#include <core/mapi/columnTags.h>
#include <core/addin/mfcmapi.h>

namespace controls::sortlistctrl
{
	class CContentsTableListCtrl;
} // namespace controls::sortlistctrl

namespace dialog
{
	class CContentsTableDlg : public CBaseDialog
	{
	public:
		CContentsTableDlg(
			_In_ std::shared_ptr<cache::CMapiObjects> lpMapiObjects,
			UINT uidTitle,
			createDialogType bCreateDialog,
			_In_opt_ LPMAPIPROP lpContainer,
			_In_opt_ LPMAPITABLE lpContentsTable,
			_In_ LPSPropTagArray sptDefaultDisplayColumnTags,
			_In_ const std::vector<columns::TagNames>& lpDefaultDisplayColumns,
			ULONG nIDContextMenu,
			ULONG ulAddInContext);
		~CContentsTableDlg();

		_Check_return_ virtual LPMAPIPROP OpenItemProp(int iSelectedItem, modifyType bModify);

	protected:
		// Overrides from base class
		virtual void CreateDialogAndMenu(UINT nIDMenuResource);
		_Check_return_ bool HandleMenu(WORD wMenuSelect) override;
		BOOL OnInitDialog() override;
		void OnInitMenu(_In_opt_ CMenu* pMenu) override;
		void OnRefreshView() override;

		virtual void OnDisplayItem();

		void SetRestrictionType(restrictionType RestrictionType);
		_Check_return_ HRESULT OpenAttachmentsFromMessage(_In_ LPMESSAGE lpMessage);
		_Check_return_ HRESULT OpenRecipientsFromMessage(_In_ LPMESSAGE lpMessage);

		controls::sortlistctrl::CContentsTableListCtrl* m_lpContentsTableListCtrl;
		LPMAPITABLE m_lpContentsTable;
		tableDisplayFlags m_displayFlags;

	private:
		virtual void HandleAddInMenuSingle(
			_In_ LPADDINMENUPARAMS lpParams,
			_In_opt_ LPMAPIPROP lpMAPIProp,
			_In_opt_ LPMAPICONTAINER lpContainer);

		// Overrides from base class
		_Check_return_ bool HandleAddInMenu(WORD wMenuSelect) override;
		void OnCancel() override;
		void OnEscHit() override;

		virtual void OnCreatePropertyStringRestriction();

		// Menu items
		void OnCreateRangeRestriction();
		void OnEditRestriction();
		void OnGetStatus();
		void OnNotificationOff();
		void OnNotificationOn();
		void OnOutputTable();
		void OnSetColumns();
		void OnSortTable();

		// Custom messages
		_Check_return_ LRESULT msgOnResetColumns(WPARAM wParam, LPARAM lParam);

		// Values held only for use in InitDialog to create our CContentsTableListCtrl
		std::vector<columns::TagNames> m_lpDefaultDisplayColumns;
		LPSPropTagArray m_sptDefaultDisplayColumnTags;
		UINT m_nIDContextMenu;

		LPMAPICONTAINER m_lpContainer;

		DECLARE_MESSAGE_MAP()
	};
} // namespace dialog