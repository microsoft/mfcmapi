#pragma once
// ContentsTableDlg.h : header file

class CContentsTableListCtrl;

#include "BaseDialog.h"

class CContentsTableDlg : public CBaseDialog
{
public:
	CContentsTableDlg(
		CParentWnd* pParentWnd,
		CMapiObjects* lpMapiObjects,
		UINT uidTitle,
		__mfcmapiCreateDialogEnum bCreateDialog,
		LPMAPITABLE lpContentsTable,
		LPSPropTagArray	sptExtraColumnTags,
		ULONG iNumExtraDisplayColumns,
		TagNames *lpExtraDisplayColumns,
		ULONG nIDContextMenu,
		ULONG ulAddInContext
		);
	virtual ~CContentsTableDlg();

	virtual HRESULT OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, LPMAPIPROP* lppMAPIProp);

protected:
	// Overrides from base class
	void CreateDialogAndMenu(UINT nIDMenuResource);
	BOOL HandleMenu(WORD wMenuSelect);
	BOOL OnInitDialog();
	void OnInitMenu(CMenu* pMenu);
	void OnRefreshView();

	virtual void OnDisplayItem();

	void SetRestrictionType(__mfcmapiRestrictionTypeEnum RestrictionType);
	HRESULT OpenAttachmentsFromMessage(LPMESSAGE lpMessage, BOOL fSaveMessageAtClose);
	HRESULT OpenRecipientsFromMessage(LPMESSAGE lpMessage);

	CContentsTableListCtrl*	m_lpContentsTableListCtrl;
	LPMAPITABLE				m_lpContentsTable;
	ULONG					m_ulDisplayFlags;

private:
	virtual void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);

	// Overrides from base class
	BOOL HandleAddInMenu(WORD wMenuSelect);
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
