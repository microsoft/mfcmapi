#include "stdafx.h"
#include "CollapsibleTextPane.h"

static wstring CLASS = L"CollapsibleTextPane";

CollapsibleTextPane* CollapsibleTextPane::Create(UINT uidLabel, bool bReadOnly)
{
	auto pane = new CollapsibleTextPane();
	if (pane)
	{
		pane->SetMultiline();
		pane->SetLabel(uidLabel, bReadOnly);
		pane->m_bCollapsible = true;
	}

	return pane;
}

void CollapsibleTextPane::Initialize(int iControl, _In_ CWnd* pParent, _In_ HDC hdc)
{
	TextPane::Initialize(iControl, pParent, hdc);
}

int CollapsibleTextPane::GetFixedHeight()
{
	auto iHeight = 0;
	if (0 != m_iControl) iHeight += m_iSmallHeightMargin; // Top margin

	// Our expand/collapse button
	iHeight += m_iButtonHeight;

	// A small margin between our button and the edit control, if we're not collapsed
	if (!m_bCollapsed)
	{
		iHeight += m_iSmallHeightMargin;
	}

	iHeight += m_iSmallHeightMargin; // Bottom margin

	return iHeight;
}

int CollapsibleTextPane::GetLines()
{
	if (m_bCollapsed)
	{
		return 0;
	}

	return LINES_MULTILINEEDIT;
}

void CollapsibleTextPane::SetWindowPos(int x, int y, int width, int height)
{
	auto hRes = S_OK;
	auto iVariableHeight = height - GetFixedHeight();
	if (0 != m_iControl)
	{
		y += m_iSmallHeightMargin;
		height -= m_iSmallHeightMargin;
	}

	ViewPane::SetWindowPos(x, y, width, height);

	y += m_iLabelHeight + m_iSmallHeightMargin;

	if (!m_bCollapsed)
	{
		EC_B(m_EditBox.ShowWindow(SW_SHOW));
		EC_B(m_EditBox.SetWindowPos(NULL, x, y, width, iVariableHeight, SWP_NOZORDER));
	}
	else
	{
		EC_B(m_EditBox.ShowWindow(SW_HIDE));
	}
}