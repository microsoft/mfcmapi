#pragma once
#include "ContentsTableDlg.h"

class CContentsTableListCtrl;
class CSingleMAPIPropListCtrl;
class CParentWnd;
class CMapiObjects;

class CRecipientsDlg : public CContentsTableDlg
{
public:
	CRecipientsDlg(
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		_In_ LPMAPITABLE lpMAPITable,
		_In_ LPMESSAGE lpMessage);
	virtual ~CRecipientsDlg();

private:
	// Overrides from base class
	void OnDeleteSelectedItem() override;
	void OnInitMenu(_In_ CMenu* pMenu) override;
	_Check_return_ HRESULT OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp) override;

	// Menu items
	void OnModifyRecipients();
	void OnRecipOptions();
	void OnSaveChanges();
	void OnViewRecipientABEntry();

	LPMESSAGE m_lpMessage;
	bool m_bViewRecipientABEntry;

	DECLARE_MESSAGE_MAP()
};