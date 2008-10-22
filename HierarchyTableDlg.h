#pragma once
// HierarchyTableDlg.h : header file

class CHierarchyTableTreeCtrl;
class CParentWnd;
class CMapiObjects;

#include "BaseDialog.h"

class CHierarchyTableDlg : public CBaseDialog
{
public:
	CHierarchyTableDlg(
		CParentWnd* pParentWnd,
		CMapiObjects* lpMapiObjects,
		UINT uidTitle,
		LPUNKNOWN lpRootContainer,
		ULONG nIDContextMenu,
		ULONG ulAddInContext
		);
	virtual ~CHierarchyTableDlg();

protected:
	// Overrides from base class
	void CreateDialogAndMenu(UINT nIDMenuResource);
	void OnInitMenu(CMenu* pMenu);

	CHierarchyTableTreeCtrl*	m_lpHierarchyTableTreeCtrl;
	ULONG						m_ulDisplayFlags;

private:
	virtual void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);

	// Overrides from base class
	BOOL HandleAddInMenu(WORD wMenuSelect);
	void OnCancel();
	BOOL OnInitDialog();
	void OnRefreshView();

	BOOL PreTranslateMessage(MSG* pMsg);

	// Menu items
	void OnDisplayHierarchyTable();
	void OnDisplayItem();
	void OnEditSearchCriteria();

	UINT m_nIDContextMenu;

	DECLARE_MESSAGE_MAP()
};
