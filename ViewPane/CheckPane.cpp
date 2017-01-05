#include "stdafx.h"
#include "CheckPane.h"

static wstring CLASS = L"CheckPane";

CheckPane* CheckPane::Create(UINT uidLabel, bool bVal, bool bReadOnly)
{
	auto pane = new CheckPane();
	if (pane)
	{
		pane->m_bCheckValue = bVal;
		pane->SetLabel(uidLabel, bReadOnly);
	}

	return pane;
}

int CheckPane::GetMinWidth(_In_ HDC hdc)
{
	auto label = ViewPane::GetMinWidth(hdc);
	auto check = GetSystemMetrics(SM_CXMENUCHECK);
	auto edge = GetSystemMetrics(SM_CXEDGE);
	DebugPrint(DBGDraw, L"CheckPane::GetMinWidth Label:%d + check:%d + edge:%d = minwidth:%d\n",
		label,
		check,
		edge,
		label + edge + check);
	return label + edge + check;
}

int CheckPane::GetFixedHeight()
{
	return m_iButtonHeight;
}

int CheckPane::GetLines()
{
	return 0;
}

void CheckPane::Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC /*hdc*/)
{
	ViewPane::Initialize(iControl, pParent, nullptr);

	auto hRes = S_OK;

	EC_B(m_Check.Create(
		NULL,
		WS_TABSTOP
		| WS_CHILD
		| WS_CLIPSIBLINGS
		| WS_VISIBLE
		| BS_AUTOCHECKBOX
		| (m_bReadOnly ? WS_DISABLED : 0),
		CRect(0, 0, 0, 0),
		pParent,
		m_nID));
	m_Check.SetCheck(m_bCheckValue);
	SetWindowTextW(m_Check.m_hWnd, m_szLabel.c_str());

	m_bInitialized = true;
}

void CheckPane::SetWindowPos(int x, int y, int width, int height)
{
	auto hRes = S_OK;
	DebugPrint(DBGDraw, L"CheckPane::SetWindowPos x:%d width:%d \n",
		x,
		width);
	EC_B(m_Check.SetWindowPos(NULL, x, y, width, height, SWP_NOZORDER));
}

void CheckPane::CommitUIValues()
{
	m_bCheckValue = 0 != m_Check.GetCheck();
	m_bCommitted = true;
}

bool CheckPane::GetCheck() const
{
	if (!m_bCommitted) return 0 != m_Check.GetCheck();
	return m_bCheckValue;
}
