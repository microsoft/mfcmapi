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