#include "stdafx.h"
#include <UI/ViewPane/CheckPane.h>
#include <UI/UIFunctions.h>

static wstring CLASS = L"CheckPane";

CheckPane* CheckPane::Create(UINT uidLabel, bool bVal, bool bReadOnly)
{
	auto pane = new (std::nothrow) CheckPane();
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
	auto edge = check / 5;
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

void CheckPane::Draw(_In_ HWND hWnd, _In_ HDC hDC, _In_ const RECT& rc, UINT itemState)
{
	WCHAR szButton[255];
	GetWindowTextW(hWnd, szButton, _countof(szButton));
	auto iState = ::SendMessage(hWnd, BM_GETSTATE, NULL, NULL);
	auto bGlow = (iState & BST_HOT) != 0;
	auto bChecked = (iState & BST_CHECKED) != 0;
	auto bDisabled = (itemState & CDIS_DISABLED) != 0;
	auto bFocused = (itemState & CDIS_FOCUS) != 0;

	auto lCheck = GetSystemMetrics(SM_CXMENUCHECK);
	auto lEdge = lCheck / 5;
	RECT rcCheck = { 0 };
	rcCheck.left = rc.left;
	rcCheck.right = rcCheck.left + lCheck;
	rcCheck.top = (rc.bottom - rc.top - lCheck) / 2;
	rcCheck.bottom = rcCheck.top + lCheck;

	FillRect(hDC, &rc, GetSysBrush(cBackground));
	FrameRect(hDC, &rcCheck, GetSysBrush(bDisabled ? cFrameUnselected : bGlow || bFocused ? cGlow : cFrameSelected));
	if (bChecked)
	{
		auto rcFill = rcCheck;
		auto deflate = lEdge;
		InflateRect(&rcFill, -deflate, -deflate);
		FillRect(hDC, &rcFill, GetSysBrush(cGlow));
	}

	auto rcLabel = rc;
	rcLabel.left = rcCheck.right + lEdge;
	rcLabel.right = rcLabel.left + GetTextExtentPoint32(hDC, szButton).cx;

	DebugPrint(DBGDraw, L"CheckButton::Draw left:%d width:%d checkwidth:%d space:%d labelwidth:%d (scroll:%d 2frame:%d), \"%ws\"\n",
		rc.left,
		rc.right - rc.left,
		rcCheck.right - rcCheck.left,
		lEdge,
		rcLabel.right - rcLabel.left,
		GetSystemMetrics(SM_CXVSCROLL),
		2 * GetSystemMetrics(SM_CXFIXEDFRAME),
		szButton);

	DrawSegoeTextW(
		hDC,
		szButton,
		bDisabled ? MyGetSysColor(cTextDisabled) : MyGetSysColor(cText),
		rcLabel,
		false,
		DT_SINGLELINE | DT_VCENTER);
}