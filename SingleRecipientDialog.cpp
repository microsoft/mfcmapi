// SingleRecipientDialog.cpp : implementation file
//

#include "stdafx.h"
#include "SingleRecipientDialog.h"
#include "SingleMAPIPropListCtrl.h"
#include "MAPIFunctions.h"
#include "MFCUtilityFunctions.h"

static TCHAR* CLASS = _T("SingleRecipientDialog");

/////////////////////////////////////////////////////////////////////////////
// SingleRecipientDialog dialog


SingleRecipientDialog::SingleRecipientDialog(
	_In_ CParentWnd* pParentWnd,
	_In_ CMapiObjects* lpMapiObjects,
	_In_opt_ LPMAILUSER lpMAPIProp):
CBaseDialog(
	pParentWnd,
	lpMapiObjects,
	NULL)
{
	TRACE_CONSTRUCTOR(CLASS);
	HRESULT hRes = S_OK;
	EC_B(m_szTitle.LoadString(IDS_ADDRESS_BOOK_ENTRY));

	m_lpMailUser = lpMAPIProp;
	if (m_lpMailUser) m_lpMailUser->AddRef();

	CBaseDialog::CreateDialogAndMenu(NULL, NULL, NULL);
} // SingleRecipientDialog::SingleMessageDialog

SingleRecipientDialog::~SingleRecipientDialog()
{
	TRACE_DESTRUCTOR(CLASS);
	if (m_lpMailUser) m_lpMailUser->Release();
	m_lpMailUser = NULL;
} // SingleRecipientDialog::~SingleRecipientDialog

BOOL SingleRecipientDialog::OnInitDialog()
{
	HRESULT hRes = S_OK;
	BOOL bRet = CBaseDialog::OnInitDialog();

	if (m_lpMailUser)
	{
		// Get a property for the title bar
		m_szTitle = GetTitle(m_lpMailUser);
	}

	UpdateTitleBarText(NULL);

	hRes = S_OK;
	EC_H(m_lpPropDisplay->SetDataSource(m_lpMailUser, NULL, true));

	return bRet;
} // SingleRecipientDialog::OnInitDialog

BEGIN_MESSAGE_MAP(SingleRecipientDialog, CBaseDialog)
	ON_COMMAND(ID_REFRESHVIEW, OnRefreshView)
END_MESSAGE_MAP()

// Clear the current list and get a new one with whatever code we've got in LoadMAPIPropList
void SingleRecipientDialog::OnRefreshView()
{
	if (!m_lpPropDisplay) return;
	(void) m_lpPropDisplay->RefreshMAPIPropList();
} // SingleRecipientDialog::OnRefreshView