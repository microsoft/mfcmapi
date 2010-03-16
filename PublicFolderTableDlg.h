#pragma once
// PublicFolderTableDlg.h : header file

class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "ContentsTableDlg.h"

class CPublicFolderTableDlg : public CContentsTableDlg
{
public:
	CPublicFolderTableDlg(
		CParentWnd* pParentWnd,
		CMapiObjects* lpMapiObjects,
		LPCTSTR lpszServerName,
		LPMAPITABLE	lpMAPITable);
	virtual ~CPublicFolderTableDlg();

private:
	// Overrides from base class
	void CreateDialogAndMenu(UINT nIDMenuResource);
	void OnCreatePropertyStringRestriction();
	void OnDisplayItem();
	HRESULT OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, LPMAPIPROP* lppMAPIProp);

	LPTSTR m_lpszServerName;
};

