// SingleMessageDialog.cpp : implementation file
//

#include "stdafx.h"
#include "SingleMessageDialog.h"
#include "SingleMAPIPropListCtrl.h"
#include "MAPIFunctions.h"
#include "MFCUtilityFunctions.h"

static TCHAR* CLASS = _T("SingleMessageDialog");

/////////////////////////////////////////////////////////////////////////////
// SingleMessageDialog dialog


SingleMessageDialog::SingleMessageDialog(
	_In_ CParentWnd* pParentWnd,
	_In_ CMapiObjects* lpMapiObjects,
	_In_opt_ LPMESSAGE lpMAPIProp):
CBaseDialog(
	pParentWnd,
	lpMapiObjects,
	NULL)
{
	TRACE_CONSTRUCTOR(CLASS);
	HRESULT hRes = S_OK;
	EC_B(m_szTitle.LoadString(IDS_MESSAGE));

	m_lpMessage = lpMAPIProp;
	if (m_lpMessage) m_lpMessage->AddRef();

	CBaseDialog::CreateDialogAndMenu(IDR_MENU_MESSAGE, NULL, NULL);
} // SingleMessageDialog::SingleMessageDialog

SingleMessageDialog::~SingleMessageDialog()
{
	TRACE_DESTRUCTOR(CLASS);
	if (m_lpMessage) m_lpMessage->Release();
	m_lpMessage = NULL;
} // SingleMessageDialog::~SingleMessageDialog

BOOL SingleMessageDialog::OnInitDialog()
{
	HRESULT hRes = S_OK;
	BOOL bRet = CBaseDialog::OnInitDialog();

	if (m_lpMessage)
	{
		// Get a property for the title bar
		m_szTitle = GetTitle(m_lpMessage);
	}

	UpdateTitleBarText(NULL);

	hRes = S_OK;
	EC_H(m_lpPropDisplay->SetDataSource(m_lpMessage, NULL, false));

	return bRet;
} // SingleMessageDialog::OnInitDialog

BEGIN_MESSAGE_MAP(SingleMessageDialog, CBaseDialog)
	ON_COMMAND(ID_REFRESHVIEW, OnRefreshView)
	ON_COMMAND(ID_ATTACHMENTPROPERTIES, OnAttachmentProperties)
	ON_COMMAND(ID_RECIPIENTPROPERTIES, OnRecipientProperties)
END_MESSAGE_MAP()

// Clear the current list and get a new one with whatever code we've got in LoadMAPIPropList
void SingleMessageDialog::OnRefreshView()
{
	if (!m_lpPropDisplay) return;
	(void) m_lpPropDisplay->RefreshMAPIPropList();
} // SingleMessageDialog::OnRefreshView

void SingleMessageDialog::OnAttachmentProperties()
{
	HRESULT hRes = S_OK;
	if (!m_lpMessage) return;

	EC_H(DisplayTable(m_lpMessage, PR_MESSAGE_ATTACHMENTS, otDefault, this));
} // SingleMessageDialog::OnAttachmentProperties

void SingleMessageDialog::OnRecipientProperties()
{
	HRESULT hRes = S_OK;
	if (!m_lpMessage) return;

	EC_H(DisplayTable(m_lpMessage, PR_MESSAGE_RECIPIENTS, otDefault, this));
} // SingleMessageDialog::OnRecipientProperties
