#pragma once
// AclDlg.h : header file

#include "ContentsTableDlg.h"

class CAclDlg : public CContentsTableDlg
{
public:
	CAclDlg(
		CParentWnd* pParentWnd,
		CMapiObjects* lpMapiObjects,
		LPEXCHANGEMODIFYTABLE lpExchTbl,
		BOOL bFreeBusyVisible);
	virtual ~CAclDlg();

private:
	// Overrides from base class
	void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);
	void OnDeleteSelectedItem();
	void OnInitMenu(CMenu* pMenu);
	void OnRefreshView();
	HRESULT OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, LPMAPIPROP* lppMAPIProp);

	// Menu items
	void OnAddItem();
	void OnModifySelectedItem();

	HRESULT GetSelectedItems(ULONG ulFlags, ULONG ulRowFlags, LPROWLIST* lppRowList);

	LPEXCHANGEMODIFYTABLE	m_lpExchTbl;
	ULONG					m_ulTableFlags;

	DECLARE_MESSAGE_MAP()
};