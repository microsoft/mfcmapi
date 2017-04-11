#pragma once
class CContentsTableListCtrl;

#include "BaseDialog.h"

class SingleMessageDialog : public CBaseDialog
{
public:
	SingleMessageDialog(
		_In_ CParentWnd* pParentWnd,
		_In_ CMapiObjects* lpMapiObjects,
		_In_opt_ LPMESSAGE lpMAPIProp);
	virtual ~SingleMessageDialog();

protected:
	// Overrides from base class
	BOOL OnInitDialog() override;

private:
	LPMESSAGE m_lpMessage;

	// Menu items
	void OnRefreshView() override;
	void OnAttachmentProperties();
	void OnRecipientProperties();
	void OnRTFSync();

	void OnTestEditBody();
	void OnTestEditHTML();
	void OnTestEditRTF();
	void OnSaveChanges();

	DECLARE_MESSAGE_MAP()
};
