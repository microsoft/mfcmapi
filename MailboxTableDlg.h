#pragma once
// MailboxTableDlg.h : header file

class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

#include "ContentsTableDlg.h"

class CMailboxTableDlg : public CContentsTableDlg
{
public:
	CMailboxTableDlg(
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		_In_z_ LPCTSTR lpszServerName,
		_In_ LPMAPITABLE lpMAPITable);
	virtual ~CMailboxTableDlg();

private:
	// Overrides from base class
	void CreateDialogAndMenu(UINT nIDMenuResource);
	void DisplayItem(ULONG ulFlags);
	void OnCreatePropertyStringRestriction();
	void OnDisplayItem();
	void OnInitMenu(_In_ CMenu* pMenu);

	// Menu items
	void OnOpenWithFlags();

	wstring m_lpszServerName;

	DECLARE_MESSAGE_MAP()
};