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
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		UINT uidTitle,
		_In_opt_ LPUNKNOWN lpRootContainer,
		ULONG nIDContextMenu,
		ULONG ulAddInContext
		);
	virtual ~CHierarchyTableDlg();

protected:
	// Overrides from base class
	void CreateDialogAndMenu(UINT nIDMenuResource);
	void OnInitMenu(_In_ CMenu* pMenu);

	CHierarchyTableTreeCtrl*	m_lpHierarchyTableTreeCtrl;
	ULONG						m_ulDisplayFlags;

private:
	virtual void HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_opt_ LPMAPIPROP lpMAPIProp,
		_In_opt_ LPMAPICONTAINER lpContainer);

	// Overrides from base class
	_Check_return_ bool HandleAddInMenu(WORD wMenuSelect);
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
