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
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		_In_ LPCTSTR lpszServerName,
		_In_ LPMAPITABLE	lpMAPITable);
	virtual ~CPublicFolderTableDlg();

private:
	// Overrides from base class
	void CreateDialogAndMenu(UINT nIDMenuResource);
	void OnCreatePropertyStringRestriction();
	void OnDisplayItem();
	_Check_return_ HRESULT OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp);

	LPTSTR m_lpszServerName;
};

