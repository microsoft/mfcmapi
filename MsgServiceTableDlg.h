#pragma once
// MsgServiceTableDlg.h : header file

class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "ContentsTableDlg.h"

class CMsgServiceTableDlg : public CContentsTableDlg
{
public:

	CMsgServiceTableDlg(
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		_In_ wstring szProfileName);
	virtual ~CMsgServiceTableDlg();

private:
	// Overrides from base class
	void HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ LPMAPICONTAINER lpContainer);
	void OnDeleteSelectedItem();
	void OnDisplayItem();
	void OnInitMenu(_In_ CMenu* pMenu);
	void OnRefreshView();
	_Check_return_ HRESULT OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp);

	// Menu items
	void OnConfigureMsgService();
	void OnOpenProfileSection();

	LPSERVICEADMIN	m_lpServiceAdmin;
	wstring m_szProfileName;

	DECLARE_MESSAGE_MAP()
};