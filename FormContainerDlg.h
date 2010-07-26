#pragma once
// FormContainerDlg.h : header file

class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "ContentsTableDlg.h"

class CFormContainerDlg : public CContentsTableDlg
{
public:

	CFormContainerDlg(
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		_In_ LPMAPIFORMCONTAINER lpFormContainer);
	virtual ~CFormContainerDlg();

private:
	// Overrides from base class
	void HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ LPMAPICONTAINER lpContainer);
	void OnDeleteSelectedItem();
	_Check_return_ BOOL OnInitDialog();
	void OnInitMenu(_In_ CMenu* pMenu);
	void OnRefreshView();
	_Check_return_ HRESULT OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp);

	// Menu items
	void OnCalcFormPropSet();
	void OnGetDisplay();
	void OnInstallForm();
	void OnRemoveForm();
	void OnResolveMessageClass();
	void OnResolveMultipleMessageClasses();

	LPMAPIFORMCONTAINER m_lpFormContainer;

	DECLARE_MESSAGE_MAP()
};