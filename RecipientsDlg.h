#pragma once
// RecipientsDlg.h : header file

class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "ContentsTableDlg.h"

class CRecipientsDlg : public CContentsTableDlg
{
public:
	CRecipientsDlg(
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		_In_ LPMAPITABLE	lpMAPITable,
		_In_ LPMESSAGE lpMessage);
	virtual ~CRecipientsDlg();

private:
	// Overrides from base class
	void OnDeleteSelectedItem();
	void OnInitMenu(_In_ CMenu* pMenu);

	// Menu items
	void OnModifyRecipients();
	void OnRecipOptions();
	void OnSaveChanges();

	LPMESSAGE m_lpMessage;

	DECLARE_MESSAGE_MAP()
};