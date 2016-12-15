#pragma once
#include <Dialogs/BaseDialog.h>

class CContentsTableListCtrl;

class CContentsTableDlg : public CBaseDialog
{
public:
	CContentsTableDlg(
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		UINT uidTitle,
		__mfcmapiCreateDialogEnum bCreateDialog,
		_In_opt_ LPMAPITABLE lpContentsTable,
		_In_ LPSPropTagArray sptExtraColumnTags,
		_In_ const vector<TagNames>& lpExtraDisplayColumns,
		ULONG nIDContextMenu,
		ULONG ulAddInContext
	);
	virtual ~CContentsTableDlg();

	_Check_return_ virtual HRESULT OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp);

protected:
	// Overrides from base class
	virtual void CreateDialogAndMenu(UINT nIDMenuResource);
	_Check_return_ bool HandleMenu(WORD wMenuSelect) override;
	BOOL OnInitDialog() override;
	void OnInitMenu(_In_opt_ CMenu* pMenu) override;
	void OnRefreshView() override;

	virtual void OnDisplayItem();

	void SetRestrictionType(__mfcmapiRestrictionTypeEnum RestrictionType);
	_Check_return_ HRESULT OpenAttachmentsFromMessage(_In_ LPMESSAGE lpMessage);
	_Check_return_ HRESULT OpenRecipientsFromMessage(_In_ LPMESSAGE lpMessage);

	CContentsTableListCtrl* m_lpContentsTableListCtrl;
	LPMAPITABLE m_lpContentsTable;
	ULONG m_ulDisplayFlags;

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
	void OnCreateMessageRestriction();
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
	vector<TagNames> m_lpExtraDisplayColumns;
	LPSPropTagArray m_sptExtraColumnTags;
	UINT m_nIDContextMenu;

	DECLARE_MESSAGE_MAP()
};
