#pragma once
// RulesDlg.h : header file

class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "ContentsTableDlg.h"

class CRulesDlg : public CContentsTableDlg
{
public:
	CRulesDlg(
		CParentWnd* pParentWnd,
		CMapiObjects* lpMapiObjects,
		LPEXCHANGEMODIFYTABLE lpExchTbl);
	virtual ~CRulesDlg();

private:
	// Overrides from base class
	void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);
	void OnDeleteSelectedItem();
	void OnInitMenu(CMenu* pMenu);
	void OnRefreshView();

	// Menu items
	void OnModifySelectedItem();

	HRESULT GetSelectedItems(ULONG ulFlags, ULONG ulRowFlags, LPROWLIST* lppRowList);

	LPEXCHANGEMODIFYTABLE m_lpExchTbl;

	DECLARE_MESSAGE_MAP()
};