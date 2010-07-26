#pragma once
// AclDlg.h : header file

#include "ContentsTableDlg.h"

class CAclDlg : public CContentsTableDlg
{
public:
	CAclDlg(
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		_In_ LPEXCHANGEMODIFYTABLE lpExchTbl,
		BOOL bFreeBusyVisible);
	virtual ~CAclDlg();

private:
	// Overrides from base class
	void HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ LPMAPICONTAINER lpContainer);
	void OnDeleteSelectedItem();
	void OnInitMenu(_In_ CMenu* pMenu);
	void OnRefreshView();
	_Check_return_ HRESULT OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp);

	// Menu items
	void OnAddItem();
	void OnModifySelectedItem();

	_Check_return_ HRESULT GetSelectedItems(ULONG ulFlags, ULONG ulRowFlags, _In_ LPROWLIST* lppRowList);

	LPEXCHANGEMODIFYTABLE	m_lpExchTbl;
	ULONG					m_ulTableFlags;

	DECLARE_MESSAGE_MAP()
};