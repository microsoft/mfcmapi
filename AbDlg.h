#pragma once
// AbDlg.h : header file

#include "ContentsTableDlg.h"

class CAbDlg : public CContentsTableDlg
{
public:
	CAbDlg(
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		_In_ LPABCONT lpAdrBook);
	virtual ~CAbDlg();

private:
	// Overrides from base class
	void CreateDialogAndMenu(UINT nIDMenuResource);
	void HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ LPMAPICONTAINER lpContainer);
	void HandleCopy();
	_Check_return_ BOOL HandlePaste();
	void OnCreatePropertyStringRestriction();
	void OnDeleteSelectedItem();
	void OnDisplayDetails();
	void OnInitMenu(_In_ CMenu* pMenu);

	// Menu items
	void OnOpenContact();
	void OnOpenManager();
	void OnOpenOwner();

	DECLARE_MESSAGE_MAP()
};