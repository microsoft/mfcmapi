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
		CParentWnd* pParentWnd,
		CMapiObjects* lpMapiObjects,
		LPMAPITABLE	lpMAPITable,
		LPMESSAGE lpMessage);
	virtual ~CRecipientsDlg();

private:
	// Overrides from base class
	void OnDeleteSelectedItem();
	void OnInitMenu(CMenu* pMenu);

	// Menu items
	void OnModifyRecipients();
	void OnRecipOptions();
	void OnSaveChanges();

	LPMESSAGE m_lpMessage;

	DECLARE_MESSAGE_MAP()
};