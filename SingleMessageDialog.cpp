// SingleMessageDialog.cpp : implementation file
//

#include "stdafx.h"
#include "SingleMessageDialog.h"
#include "SingleMAPIPropListCtrl.h"
#include "MAPIFunctions.h"
#include "MFCUtilityFunctions.h"
#include "Editor.h"
#include "StreamEditor.h"
#include "ExtraPropTags.h"

static TCHAR* CLASS = _T("SingleMessageDialog");

/////////////////////////////////////////////////////////////////////////////
// SingleMessageDialog dialog


SingleMessageDialog::SingleMessageDialog(
	_In_ CParentWnd* pParentWnd,
	_In_ CMapiObjects* lpMapiObjects,
	_In_opt_ LPMESSAGE lpMAPIProp) :
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
	ON_COMMAND(ID_RTFSYNC, OnRTFSync)
	ON_COMMAND(ID_TESTEDITBODY, OnTestEditBody)
	ON_COMMAND(ID_TESTEDITHTML, OnTestEditHTML)
	ON_COMMAND(ID_TESTEDITRTF, OnTestEditRTF)
	ON_COMMAND(ID_SAVECHANGES, OnSaveChanges)
END_MESSAGE_MAP()

// Clear the current list and get a new one with whatever code we've got in LoadMAPIPropList
void SingleMessageDialog::OnRefreshView()
{
	if (!m_lpPropDisplay) return;
	(void)m_lpPropDisplay->RefreshMAPIPropList();
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

void SingleMessageDialog::OnRTFSync()
{
	HRESULT hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	CEditor MyData(
		this,
		IDS_CALLRTFSYNC,
		IDS_CALLRTFSYNCPROMPT,
		1,
		CEDITOR_BUTTON_OK | CEDITOR_BUTTON_CANCEL);

	MyData.InitPane(0, CreateSingleLinePane(IDS_FLAGS, NULL, false));
	MyData.SetHex(0, RTF_SYNC_RTF_CHANGED);

	WC_H(MyData.DisplayDialog());
	if (S_OK == hRes)
	{
		BOOL bMessageUpdated = false;

		if (m_lpMessage)
		{
			DebugPrint(DBGGeneric, _T("Calling RTFSync on %p with flags 0x%X\n"), m_lpMessage, MyData.GetHex(0));
			EC_MAPI(RTFSync(
				m_lpMessage,
				MyData.GetHex(0),
				&bMessageUpdated));

			DebugPrint(DBGGeneric, _T("RTFSync returned %d\n"), bMessageUpdated);

			EC_MAPI(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));

			(void)m_lpPropDisplay->RefreshMAPIPropList();
		}
	}
}

void SingleMessageDialog::OnTestEditBody()
{
	HRESULT hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (m_lpMessage)
	{
		DebugPrint(DBGGeneric, _T("Editing body on %p\n"), m_lpMessage);

		CStreamEditor MyEditor(
			this,
			IDS_PROPEDITOR,
			IDS_STREAMEDITORPROMPT,
			m_lpMessage,
			PR_BODY_W,
			false, // No stream guessing
			m_bIsAB,
			false,
			false,
			0,
			0,
			0);
		MyEditor.DisableSave();

		WC_H(MyEditor.DisplayDialog());

		(void)m_lpPropDisplay->RefreshMAPIPropList();
	}
}

void SingleMessageDialog::OnTestEditHTML()
{
	HRESULT hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (m_lpMessage)
	{
		DebugPrint(DBGGeneric, _T("Testing HTML on %p\n"), m_lpMessage);

		CStreamEditor MyEditor(
			this,
			IDS_PROPEDITOR,
			IDS_STREAMEDITORPROMPT,
			m_lpMessage,
			PR_BODY_HTML,
			false, // No stream guessing
			m_bIsAB,
			false,
			false,
			0,
			0,
			0);
		MyEditor.DisableSave();

		WC_H(MyEditor.DisplayDialog());

		(void)m_lpPropDisplay->RefreshMAPIPropList();
	}
}

void SingleMessageDialog::OnTestEditRTF()
{
	HRESULT hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (m_lpMessage)
	{
		DebugPrint(DBGGeneric, _T("Testing body on %p\n"), m_lpMessage);

		CStreamEditor MyEditor(
			this,
			IDS_PROPEDITOR,
			IDS_STREAMEDITORPROMPT,
			m_lpMessage,
			PR_RTF_COMPRESSED,
			false, // No stream guessing
			m_bIsAB,
			true,
			false,
			0,
			0,
			0);
		MyEditor.DisableSave();

		WC_H(MyEditor.DisplayDialog());

		(void)m_lpPropDisplay->RefreshMAPIPropList();
	}
}

void SingleMessageDialog::OnSaveChanges()
{
	HRESULT hRes = S_OK;
	CWaitCursor Wait; // Change the mouse to an hourglass while we work.

	if (m_lpMessage)
	{
		DebugPrint(DBGGeneric, _T("Saving changes on %p\n"), m_lpMessage);

		EC_MAPI(m_lpMessage->SaveChanges(KEEP_OPEN_READWRITE));

		(void)m_lpPropDisplay->RefreshMAPIPropList();
	}
}