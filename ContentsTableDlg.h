#pragma once
// ContentsTableDlg.h : header file
//

class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "BaseDialog.h"

/////////////////////////////////////////////////////////////////////////////
// CContentsTableDlg dialog

class CContentsTableDlg : public CBaseDialog
{
friend class CContentsTableListCtrl;
public:
	CContentsTableDlg(
		CParentWnd* pParentWnd,
		CMapiObjects *lpMapiObjects,
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

protected:
	DECLARE_MESSAGE_MAP()
	BOOL CreateDialogAndMenu(UINT nIDMenuResource);
	void SetRestrictionType(__mfcmapiRestrictionTypeEnum RestrictionType);
	virtual void	OnInitMenu(CMenu* pMenu);
	virtual void	OnRefreshView();
	virtual BOOL	HandleMenu(WORD wMenuSelect);
	virtual HRESULT	OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, LPMAPIPROP* lppMAPIProp);
	virtual void	OnDisplayItem();
	virtual BOOL	OnInitDialog();

	CContentsTableListCtrl*			m_lpContentsTableListCtrl;
	ULONG                           m_ulDisplayFlags;
	LPMAPITABLE						m_lpContentsTable;
private:

	void	OnEscHit();

	void	OnCancel();

	virtual void	OnCreateMessageRestriction();
	virtual void	OnCreatePropertyStringRestriction();
	virtual void	OnCreateRangeRestriction();
	virtual void	OnEditRestriction();
	virtual void	OnNotificationOn();
	virtual void	OnNotificationOff();
	virtual void	OnOutputTable();
	virtual void	OnSetColumns();
	virtual void	OnSortTable();
	virtual void	OnGetStatus();

	virtual BOOL HandleAddInMenu(WORD wMenuSelect);
	virtual void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);

	//Values held only for use in InitDialog to create our CContentsTableListCtrl
	ULONG					m_iNumExtraDisplayColumns;
	TagNames*				m_lpExtraDisplayColumns;
	LPSPropTagArray			m_sptExtraColumnTags;
	UINT					m_nIDContextMenu;
};
