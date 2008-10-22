#pragma once
// AttachmentsDlg.h : header file

#include "ContentsTableDlg.h"

class CAttachmentsDlg : public CContentsTableDlg
{
public:
	CAttachmentsDlg(
		CParentWnd* pParentWnd,
		CMapiObjects* lpMapiObjects,
		LPMAPITABLE	lpMAPITable,
		LPMESSAGE lpMessage);
	virtual ~CAttachmentsDlg();

private:
	// Overrides from base class
	void HandleAddInMenuSingle(
		LPADDINMENUPARAMS lpParams,
		LPMAPIPROP lpMAPIProp,
		LPMAPICONTAINER lpContainer);
	void OnDeleteSelectedItem();
	void OnInitMenu(CMenu* pMenu);
	HRESULT OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, LPMAPIPROP* lppMAPIProp);

	// Menu items
	void OnModifySelectedItem();
	void OnSaveChanges();
	void OnSaveToFile();
	void OnUseMapiModify();
	void OnViewEmbeddedMessageProps();

	LPATTACH	m_lpAttach;
	LPMESSAGE	m_lpMessage;
	BOOL		m_bDisplayAttachAsEmbeddedMessage;
	BOOL		m_bUseMapiModifyOnEmbeddedMessage;

	DECLARE_MESSAGE_MAP()
};