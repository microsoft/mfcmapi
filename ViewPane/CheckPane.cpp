#include "stdafx.h"
#include "..\stdafx.h"
#include "CheckPane.h"

static wstring CLASS = L"CheckPane";

ViewPane* CreateCheckPane(UINT uidLabel, bool bVal, bool bReadOnly)
{
	return new CheckPane(uidLabel, bReadOnly, bVal);
}

CheckPane::CheckPane(UINT uidLabel, bool bReadOnly, bool bCheck) :ViewPane(uidLabel, bReadOnly)
{
	m_bCheckValue = bCheck;
}

bool CheckPane::IsType(__ViewTypes vType)
{
	return CTRL_CHECKPANE == vType;
}

ULONG CheckPane::GetFlags()
{
	ULONG ulFlags = vpNone;
	if (m_bReadOnly) ulFlags |= vpReadonly;
	return ulFlags;
}

int CheckPane::GetMinWidth(_In_ HDC hdc)
{
	return ViewPane::GetMinWidth(hdc) + ::GetSystemMetrics(SM_CXMENUCHECK) + GetSystemMetrics(SM_CXEDGE);
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
	ViewPane::Initialize(iControl, pParent, NULL);

	HRESULT hRes = S_OK;

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
	HRESULT hRes = S_OK;
	EC_B(m_Check.SetWindowPos(NULL, x, y, width, height, SWP_NOZORDER));
}

void CheckPane::CommitUIValues()
{
	m_bCheckValue = GetCheckUseControl();
}

bool CheckPane::GetCheck()
{
	return m_bCheckValue;
}

bool CheckPane::GetCheckUseControl()
{
	return (0 != m_Check.GetCheck());
}