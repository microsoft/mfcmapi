#include "stdafx.h"
#include "..\stdafx.h"
#include "CollapsibleTextPane.h"
#include "..\UIFunctions.h"

static TCHAR* CLASS = _T("CollapsibleTextPane");

ViewPane* CreateCollapsibleTextPane(UINT uidLabel, bool bReadOnly)
{
	CollapsibleTextPane* lpPane = new CollapsibleTextPane(uidLabel, bReadOnly);
	return lpPane;
}

CollapsibleTextPane::CollapsibleTextPane(UINT uidLabel, bool bReadOnly):TextPane(uidLabel, bReadOnly, true)
{
}

bool CollapsibleTextPane::IsType(__ViewTypes vType)
{
	return (CTRL_COLLAPSIBLETEXTPANE == vType) || TextPane::IsType(vType);
}

ULONG CollapsibleTextPane::GetFlags()
{
	return TextPane::GetFlags() | vpCollapsible;
}

int CollapsibleTextPane::GetFixedHeight()
{
	int iHeight = 0;
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
	else
	{
		return LINES_MULTILINEEDIT;
	}
}

void CollapsibleTextPane::SetWindowPos(int x, int y, int width, int height)
{
	HRESULT hRes = S_OK;
	int (iVariableHeight) = height - GetFixedHeight();
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