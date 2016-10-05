// DbgView.cpp : implementation file
//

#include "stdafx.h"
#include <Dialogs/Editors/Editor.h>
#include "DbgView.h"
#include "ParentWnd.h"
#include "MFCOutput.h"
#include "String.h"

class CDbgView : public CEditor
{
public:
	CDbgView(_In_ CParentWnd* pParentWnd);
	virtual ~CDbgView();
	void AppendText(wstring szMsg) const;

private:
	_Check_return_ ULONG HandleChange(UINT nID) override;
	void  OnEditAction1() override;
	void  OnEditAction2() override;

	void OnOK() override;
	void OnCancel() override;
	bool m_bPaused;
};

CDbgView* g_DgbView = nullptr;

// Displays the debug viewer - only one may exist at a time
void DisplayDbgView(_In_ CParentWnd* pParentWnd)
{
	if (!g_DgbView) g_DgbView = new CDbgView(pParentWnd);
}

void OutputToDbgView(wstring szMsg)
{
	if (!g_DgbView) return;
	g_DgbView->AppendText(szMsg);
}

enum __DbgViewFields
{
	DBGVIEW_TAGS,
	DBGVIEW_PAUSE,
	DBGVIEW_VIEW,
	DBGVIEW_NUMFIELDS, // Must be last
};

static wstring CLASS = L"CDbgView";

CDbgView::CDbgView(_In_ CParentWnd* pParentWnd) :
	CEditor(pParentWnd, IDS_DBGVIEW, NULL, CEDITOR_BUTTON_ACTION1 | CEDITOR_BUTTON_ACTION2, IDS_CLEAR, IDS_CLOSE, NULL)
{
	TRACE_CONSTRUCTOR(CLASS);
	InitPane(DBGVIEW_TAGS, TextPane::CreateSingleLinePane(IDS_REGKEY_DEBUG_TAG, false));
	SetHex(DBGVIEW_TAGS, GetDebugLevel());
	InitPane(DBGVIEW_PAUSE, CheckPane::Create(IDS_PAUSE, false, false));
	InitPane(DBGVIEW_VIEW, TextPane::CreateMultiLinePane(NULL, true));
	m_bPaused = false;
	DisplayParentedDialog(pParentWnd, 800);
}

CDbgView::~CDbgView()
{
	TRACE_DESTRUCTOR(CLASS);
	g_DgbView = nullptr;
}

void CDbgView::OnOK()
{
	// Override does nothing except *not* call base OnOK
}

void CDbgView::OnCancel()
{
	ShowWindow(SW_HIDE);
	delete this;
}

_Check_return_ ULONG CDbgView::HandleChange(UINT nID)
{
	auto i = CEditor::HandleChange(nID);

	if (static_cast<ULONG>(-1) == i) return static_cast<ULONG>(-1);

	switch (i)
	{
	case DBGVIEW_TAGS:
	{
		auto ulTag = GetHexUseControl(DBGVIEW_TAGS);
		SetDebugLevel(ulTag);
		return true;
	}
	case DBGVIEW_PAUSE:
	{
		m_bPaused = GetCheckUseControl(DBGVIEW_PAUSE);
	}
	break;

	default:
		break;
	}

	return i;
}

// Clear
void CDbgView::OnEditAction1()
{
	if (IsValidEdit(DBGVIEW_VIEW))
	{
		auto lpPane = static_cast<TextPane*>(GetPane(DBGVIEW_VIEW));
		if (lpPane)
		{
			return lpPane->ClearView();
		}
	}
}

// Close
void CDbgView::OnEditAction2()
{
	OnCancel();
}

void CDbgView::AppendText(wstring szMsg) const
{
	if (m_bPaused) return;

	if (IsValidEdit(DBGVIEW_VIEW))
	{
		auto lpPane = static_cast<TextPane*>(GetPane(DBGVIEW_VIEW));
		if (lpPane)
		{
			lpPane->AppendString(szMsg);
		}
	}
}
