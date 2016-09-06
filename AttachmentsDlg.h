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
		_In_ LPMESSAGE lpMessage);
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
	void OnDisplayItem();
	void OnInitMenu(_In_ CMenu* pMenu);
	_Check_return_ LPATTACH OpenAttach(ULONG ulAttachNum) const;
	_Check_return_ LPMESSAGE CAttachmentsDlg::OpenEmbeddedMessage() const;
	_Check_return_ HRESULT OpenItemProp(int iSelectedItem, __mfcmapiModifyEnum bModify, _Deref_out_opt_ LPMAPIPROP* lppMAPIProp);

	// Menu items
	void OnModifySelectedItem();
	void OnSaveChanges();
	void OnSaveToFile();
	void OnViewEmbeddedMessageProps();
	void OnAddAttachment();

	LPATTACH	m_lpAttach; // Currently opened attachment
	ULONG		m_ulAttachNum; // Currently opened attachment number
	LPMESSAGE	m_lpMessage;
	bool		m_bDisplayAttachAsEmbeddedMessage;

	DECLARE_MESSAGE_MAP()
};