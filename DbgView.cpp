// DbgView.cpp : implementation file
//

#include "stdafx.h"
#include "Editor.h"
#include "DbgView.h"
#include "ParentWnd.h"
#include "MFCOutput.h"

class CDbgView : public CEditor
{
public:
	CDbgView(_In_ CParentWnd* pParentWnd);
	virtual ~CDbgView();
	void AppendText(_In_z_ LPCTSTR szMsg);

private:
	_Check_return_ ULONG HandleChange(UINT nID);
	void  OnEditAction1();
	void  OnEditAction2();

	void OnOK();
	void OnCancel();
	bool m_bPaused;
};

CDbgView* g_DgbView = NULL;

// Displays the debug viewer - only one may exist at a time
void DisplayDbgView(_In_ CParentWnd* pParentWnd)
{
	if (!g_DgbView) g_DgbView = new CDbgView(pParentWnd);
} // DisplayDbgView

void OutputToDbgView(_In_z_ LPCTSTR szMsg)
{
	if (!g_DgbView) return;
	g_DgbView->AppendText(szMsg);
} // OutputToDbgView

enum __DbgViewFields
{
	DBGVIEW_TAGS,
	DBGVIEW_PAUSE,
	DBGVIEW_VIEW,
	DBGVIEW_NUMFIELDS, // Must be last
};

static TCHAR* CLASS = _T("CDbgView");

CDbgView::CDbgView(_In_ CParentWnd* pParentWnd):
CEditor(pParentWnd,IDS_DBGVIEW,IDS_DBGVIEWPROMPT,0,CEDITOR_BUTTON_ACTION1|CEDITOR_BUTTON_ACTION2,IDS_CLEAR,IDS_CLOSE)
{
	TRACE_CONSTRUCTOR(CLASS);
	CreateControls(DBGVIEW_NUMFIELDS);
	InitSingleLine(DBGVIEW_TAGS,IDS_REGKEY_DEBUG_TAG,NULL,false);
	SetHex(DBGVIEW_TAGS,GetDebugLevel());
	InitCheck(DBGVIEW_PAUSE,IDS_PAUSE,false,false);
	InitMultiLine(DBGVIEW_VIEW,NULL,NULL,true);
	m_bPaused = false;
	DisplayParentedDialog(pParentWnd,800);
} // CDbgView::CDbgView

CDbgView::~CDbgView()
{
	TRACE_DESTRUCTOR(CLASS);
	g_DgbView = NULL;
} // CDbgView::~CDbgView

void CDbgView::OnOK()
{
	// Override does nothing except *not* call base OnOK
} // CDbgView::OnOK

void CDbgView::OnCancel()
{
	ShowWindow(SW_HIDE);
	delete this;
} // CDbgView::OnCancel

_Check_return_ ULONG CDbgView::HandleChange(UINT nID)
{
	ULONG i = CEditor::HandleChange(nID);

	if ((ULONG) -1 == i) return (ULONG) -1;

	switch (i)
	{
	case (DBGVIEW_TAGS):
		{
			ULONG ulTag = GetHexUseControl(DBGVIEW_TAGS);
			SetDebugLevel(ulTag); return true;
		}
		break;
	case (DBGVIEW_PAUSE):
		{
			m_bPaused = GetCheckUseControl(DBGVIEW_PAUSE);
		}
		break;

	default:
		break;
	}

	return i;
} // CDbgView::HandleChange

// Clear
void CDbgView::OnEditAction1()
{
	ClearView(DBGVIEW_VIEW);
} // CDbgView::OnEditAction1

// Close
void CDbgView::OnEditAction2()
{
	OnCancel();
} // CDbgView::OnEditAction1

void CDbgView::AppendText(_In_z_ LPCTSTR szMsg)
{
	if (m_bPaused) return;
	AppendString(DBGVIEW_VIEW,szMsg);
} // CDbgView::AppendText