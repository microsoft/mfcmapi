#pragma once
// SingleRecipientDialog.h : header file

class CContentsTableListCtrl;

#include "BaseDialog.h"

class SingleRecipientDialog : public CBaseDialog
{
public:
	SingleRecipientDialog(
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		_In_opt_ LPMAILUSER lpMailUser);
	virtual ~SingleRecipientDialog();

protected:
	// Overrides from base class
	void CreateDialogAndMenu(UINT nIDMenuResource);
	BOOL OnInitDialog();

private:
	LPMAILUSER m_lpMailUser;

	// Menu items
	void OnRefreshView();

	DECLARE_MESSAGE_MAP()
};
