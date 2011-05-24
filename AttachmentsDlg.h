#pragma once
// AttachmentsDlg.h : header file

#include "ContentsTableDlg.h"

class CAttachmentsDlg : public CContentsTableDlg
{
public:
	CAttachmentsDlg(
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		_In_ LPMAPITABLE	lpMAPITable,
		_In_ LPMESSAGE lpMessage,
		bool bSaveMessageAtClose);
	virtual ~CAttachmentsDlg();

private:
	// Overrides from base class
	void HandleAddInMenuSingle(
		_In_ LPADDINMENUPARAMS lpParams,
		_In_ LPMAPIPROP lpMAPIProp,
		_In_ LPMAPICONTAINER lpContainer);
	void HandleCopy();
	_Check_return_ bool HandlePaste();
	void OnDeleteSelectedItem();
	void OnInitMenu(_In_ CMenu* pMenu);
	_Check_return_ HRESULT OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp);

	// Menu items
	void OnModifySelectedItem();
	void OnSaveChanges();
	void OnSaveToFile();
	void OnUseMapiModify();
	void OnViewEmbeddedMessageProps();
	void OnAttachmentProperties();
	void OnRecipientProperties();

	_Check_return_ HRESULT GetEmbeddedMessage(int iIndex, _Deref_out_opt_ LPMESSAGE *lppMessage);

	LPATTACH	m_lpAttach;
	LPMESSAGE	m_lpMessage;
	bool		m_bDisplayAttachAsEmbeddedMessage;
	bool		m_bUseMapiModifyOnEmbeddedMessage;
	bool		m_bSaveMessageAtClose;

	DECLARE_MESSAGE_MAP()
};