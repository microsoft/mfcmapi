#pragma once
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
	void CreateDialogAndMenu(UINT nIDMenuResource) override;
	void HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ LPMAPICONTAINER lpContainer) override;
	void HandleCopy() override;
	_Check_return_ bool HandlePaste() override;
	void OnCreatePropertyStringRestriction() override;
	void OnDeleteSelectedItem() override;
	void OnDisplayDetails();
	void OnInitMenu(_In_ CMenu* pMenu);

	// Menu items
	void OnOpenContact();
	void OnOpenManager();
	void OnOpenOwner();

	DECLARE_MESSAGE_MAP()
};