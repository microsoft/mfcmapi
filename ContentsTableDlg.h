#pragma once
// ContentsTableDlg.h : header file

class CContentsTableListCtrl;

#include "BaseDialog.h"

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
		ULONG iNumExtraDisplayColumns,
		_In_count_(iNumExtraDisplayColumns) TagNames* lpExtraDisplayColumns,
		ULONG nIDContextMenu,
		ULONG ulAddInContext
		);
	virtual ~CContentsTableDlg();

	_Check_return_ virtual HRESULT OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp);

protected:
	// Overrides from base class
	void CreateDialogAndMenu(UINT nIDMenuResource);
	_Check_return_ BOOL HandleMenu(WORD wMenuSelect);
	_Check_return_ BOOL OnInitDialog();
	void OnInitMenu(_In_opt_ CMenu* pMenu);
	void OnRefreshView();

	virtual void OnDisplayItem();

	void SetRestrictionType(__mfcmapiRestrictionTypeEnum RestrictionType);
	_Check_return_ HRESULT OpenAttachmentsFromMessage(_In_ LPMESSAGE lpMessage, BOOL fSaveMessageAtClose);
	_Check_return_ HRESULT OpenRecipientsFromMessage(_In_ LPMESSAGE lpMessage);

	CContentsTableListCtrl*	m_lpContentsTableListCtrl;
	LPMAPITABLE				m_lpContentsTable;
	ULONG					m_ulDisplayFlags;

private:
	virtual void HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_opt_ LPMAPIPROP lpMAPIProp,
		_In_opt_ LPMAPICONTAINER lpContainer);

	// Overrides from base class
	_Check_return_ BOOL HandleAddInMenu(WORD wMenuSelect);
	void OnCancel();
	void OnEscHit();

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

	// Values held only for use in InitDialog to create our CContentsTableListCtrl
	ULONG			m_iNumExtraDisplayColumns;
	TagNames*		m_lpExtraDisplayColumns;
	LPSPropTagArray	m_sptExtraColumnTags;
	UINT			m_nIDContextMenu;

	DECLARE_MESSAGE_MAP()
};
