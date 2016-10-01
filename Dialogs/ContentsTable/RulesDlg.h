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
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		_In_ LPEXCHANGEMODIFYTABLE lpExchTbl);
	virtual ~CRulesDlg();

private:
	// Overrides from base class
	void HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ LPMAPICONTAINER lpContainer);
	void OnDeleteSelectedItem();
	void OnInitMenu(_In_opt_ CMenu* pMenu);
	void OnRefreshView();

	// Menu items
	void OnModifySelectedItem();

	_Check_return_ HRESULT GetSelectedItems(ULONG ulFlags, ULONG ulRowFlags, _In_ LPROWLIST* lppRowList) const;

	LPEXCHANGEMODIFYTABLE m_lpExchTbl;

	DECLARE_MESSAGE_MAP()
};