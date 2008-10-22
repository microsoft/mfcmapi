#pragma once
// ProviderTableDlg.h : header file

class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "ContentsTableDlg.h"

class CProviderTableDlg : public CContentsTableDlg
{
public:
	CProviderTableDlg(
		CParentWnd* pParentWnd,
		CMapiObjects* lpMapiObjects,
		LPMAPITABLE lpMAPITable,
		LPPROVIDERADMIN lpProviderAdmin);
	virtual ~CProviderTableDlg();

private:
	// Overrides from base class
	void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);
	HRESULT	OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, LPMAPIPROP* lppMAPIProp);

	// Menu items
	void OnOpenProfileSection();

	LPPROVIDERADMIN m_lpProviderAdmin;

	DECLARE_MESSAGE_MAP()
};