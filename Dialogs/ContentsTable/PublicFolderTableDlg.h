#pragma once
#include "ContentsTableDlg.h"

class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

class CPublicFolderTableDlg : public CContentsTableDlg
{
public:
	CPublicFolderTableDlg(
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		_In_ const wstring& lpszServerName,
		_In_ LPMAPITABLE lpMAPITable);
	virtual ~CPublicFolderTableDlg();

private:
	// Overrides from base class
	void CreateDialogAndMenu(UINT nIDMenuResource) override;
	void OnCreatePropertyStringRestriction() override;
	void OnDisplayItem() override;
	_Check_return_ HRESULT OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp) override;

	wstring m_lpszServerName;
};

