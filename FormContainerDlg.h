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
		CParentWnd* pParentWnd,
		CMapiObjects* lpMapiObjects,
		LPMAPIFORMCONTAINER lpFormContainer);
	virtual ~CFormContainerDlg();

private:
	// Overrides from base class
	void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);
	void OnDeleteSelectedItem();
	BOOL OnInitDialog();
	void OnInitMenu(CMenu* pMenu);
	void OnRefreshView();
	HRESULT OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, LPMAPIPROP* lppMAPIProp);

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