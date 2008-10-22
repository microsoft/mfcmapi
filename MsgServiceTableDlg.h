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
		CParentWnd* pParentWnd,
		CMapiObjects* lpMapiObjects,
		LPCSTR szProfileName);
	virtual ~CMsgServiceTableDlg();

private:
	// Overrides from base class
	void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);
	void OnDeleteSelectedItem();
	void OnDisplayItem();
	void OnInitMenu(CMenu* pMenu);
	void OnRefreshView();
	HRESULT	OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, LPMAPIPROP* lppMAPIProp);

	// Menu items
	void OnConfigureMsgService();
	void OnOpenProfileSection();

	LPSERVICEADMIN	m_lpServiceAdmin;
	LPSTR			m_szProfileName;

	DECLARE_MESSAGE_MAP()
};