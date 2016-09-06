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
		_In_ string szProfileName);
	virtual ~CMsgServiceTableDlg();

private:
	// Overrides from base class
	void HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ LPMAPICONTAINER lpContainer) override;
	void OnDeleteSelectedItem() override;
	void OnDisplayItem() override;
	void OnRefreshView() override;
	_Check_return_ HRESULT OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp) override;
	void OnInitMenu(_In_ CMenu* pMenu);

	// Menu items
	void OnConfigureMsgService();
	void OnOpenProfileSection();

	LPSERVICEADMIN	m_lpServiceAdmin;
	string m_szProfileName;

	DECLARE_MESSAGE_MAP()
};