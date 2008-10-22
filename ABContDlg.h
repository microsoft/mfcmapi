#pragma once
// AbContDlg.h : header file

#include "HierarchyTableDlg.h"

class CAbContDlg : public CHierarchyTableDlg
{
public:
	CAbContDlg(
		CParentWnd* pParentWnd,
		CMapiObjects* lpMapiObjects);
	virtual ~CAbContDlg();

private:
	// Overrides from base class
	void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);

	// Menu items
	void OnSetDefaultDir();
	void OnSetPAB();

	DECLARE_MESSAGE_MAP()
};