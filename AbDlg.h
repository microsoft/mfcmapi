#pragma once
// AbDlg.h : header file

#include "ContentsTableDlg.h"

class CAbDlg : public CContentsTableDlg
{
public:
	CAbDlg(
		CParentWnd* pParentWnd,
		CMapiObjects* lpMapiObjects,
		LPABCONT lpAdrBook);
	virtual ~CAbDlg();

private:
	// Overrides from base class
	void CreateDialogAndMenu(UINT nIDMenuResource);
	void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);
	BOOL HandleCopy();
	BOOL HandlePaste();
	void OnCreatePropertyStringRestriction();
	void OnDeleteSelectedItem();
	void OnDisplayDetails();
	void OnInitMenu(CMenu* pMenu);

	// Menu items
	void OnOpenContact();
	void OnOpenManager();
	void OnOpenOwner();

	DECLARE_MESSAGE_MAP()
};