#pragma once
#include "HierarchyTableDlg.h"

class CAbContDlg : public CHierarchyTableDlg
{
public:
	CAbContDlg(
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects);
	virtual ~CAbContDlg();

private:
	// Overrides from base class
	void HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ LPMAPICONTAINER lpContainer) override;

	// Menu items
	void OnSetDefaultDir();
	void OnSetPAB();

	DECLARE_MESSAGE_MAP()
};