#pragma once
#include "ContentsTableDlg.h"

class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

class CMailboxTableDlg : public CContentsTableDlg
{
public:
	CMailboxTableDlg(
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		_In_ const std::wstring& lpszServerName,
		_In_ LPMAPITABLE lpMAPITable);
	virtual ~CMailboxTableDlg();

private:
	// Overrides from base class
	void CreateDialogAndMenu(UINT nIDMenuResource) override;
	void DisplayItem(ULONG ulFlags);
	void OnCreatePropertyStringRestriction() override;
	void OnDisplayItem() override;
	void OnInitMenu(_In_ CMenu* pMenu) override;

	// Menu items
	void OnOpenWithFlags();

	std::wstring m_lpszServerName;

	DECLARE_MESSAGE_MAP()
};